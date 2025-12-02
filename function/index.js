const crypto = require('crypto');
const fetch = require('node-fetch');

const PRIVATE_KEY_PEM = process.env.GRAPH_SUB_PRIVATE_KEY || null;
const EVENTGRID_TOPIC_ENDPOINT = process.env.EVENTGRID_TOPIC_ENDPOINT || null;
const EVENTGRID_TOPIC_KEY = process.env.EVENTGRID_TOPIC_KEY || null;
const ENCRYPTION_CERT_ID = process.env.ENCRYPTION_CERT_ID || 'v1';

function rsaOaepDecrypt(encryptedBase64) {
  const encrypted = Buffer.from(encryptedBase64, 'base64');
  return crypto.privateDecrypt(
    {
      key: PRIVATE_KEY_PEM,
      oaepHash: 'sha1',
      padding: crypto.constants.RSA_PKCS1_OAEP_PADDING
    },
    encrypted
  );
}

function verifySignature(symmetricKey, dataB64, signatureB64) {
  const data = Buffer.from(dataB64, 'base64');
  const sig = Buffer.from(signatureB64, 'base64');
  const hmac = crypto.createHmac('sha256', symmetricKey);
  hmac.update(data);
  const computed = hmac.digest();
  return computed.length === sig.length && crypto.timingSafeEqual(computed, sig);
}

function aesCbcDecrypt(symmetricKey, dataB64) {
  const data = Buffer.from(dataB64, 'base64');
  const iv = symmetricKey.slice(0, 16);
  const decipher = crypto.createDecipheriv('aes-256-cbc', symmetricKey, iv);
  decipher.setAutoPadding(true);
  let out = decipher.update(data);
  out = Buffer.concat([out, decipher.final()]);
  return out.toString('utf8');
}

async function publishToEventGrid(events) {
  if (!EVENTGRID_TOPIC_ENDPOINT || !EVENTGRID_TOPIC_KEY) {
    throw new Error('EVENTGRID_TOPIC_ENDPOINT or EVENTGRID_TOPIC_KEY not set');
  }
  const res = await fetch(EVENTGRID_TOPIC_ENDPOINT, {
    method: 'POST',
    headers: {
      'aeg-sas-key': EVENTGRID_TOPIC_KEY,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(events)
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Publish failed ${res.status}: ${txt}`);
  }
}

module.exports = async function (context, req) {
  context.log('Webhook invoked');

  // Graph validation handshake may come as plain query param or in JSON body
  const validationToken = (req.query && req.query.validationToken) || (req.body && req.body.validationToken);
  if (validationToken) {
    context.log('Validation handshake request');
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'text/plain' },
      body: validationToken
    };
    return;
  }

  if (!req.body || !Array.isArray(req.body.value) && !Array.isArray(req.body)) {
    context.res = { status: 204 };
    return;
  }

  const notifications = Array.isArray(req.body.value) ? req.body.value : req.body;

  const eventsToPublish = [];

  for (const item of notifications) {
    try {
      let resourceObj = null;

      if (item.encryptedContent) {
        const enc = item.encryptedContent;
        if (!PRIVATE_KEY_PEM) {
          context.log.warn('Private key missing; cannot decrypt');
          continue;
        }
        if (enc.encryptionCertificateId !== ENCRYPTION_CERT_ID) {
          context.log.warn('Unexpected cert id', enc.encryptionCertificateId);
          continue;
        }
        const symmetricKey = rsaOaepDecrypt(enc.dataKey);
        if (!verifySignature(symmetricKey, enc.data, enc.dataSignature)) {
          context.log.warn('Signature mismatch - skip');
          continue;
        }
        const decrypted = aesCbcDecrypt(symmetricKey, enc.data);
        resourceObj = JSON.parse(decrypted);
      } else {
        resourceObj = item.resourceData || { id: item.resource };
      }

      const egEvent = {
        id: `${item.subscriptionId}-${resourceObj.id || item.resource || crypto.randomBytes(6).toString('hex')}`,
        subject: item.resource || '/unknown',
        eventType: `Microsoft.Graph.CalendarEvent.${(item.changeType || 'unknown').toUpperCase()}`,
        eventTime: new Date().toISOString(),
        data: {
          changeType: item.changeType,
          subscriptionId: item.subscriptionId,
          event: resourceObj
        },
        dataVersion: '1.0'
      };

      eventsToPublish.push(egEvent);
    } catch (err) {
      context.log.error('Error processing notification', err);
    }
  }

  if (eventsToPublish.length > 0) {
    try {
      await publishToEventGrid(eventsToPublish);
      context.log(`Published ${eventsToPublish.length} events`);
    } catch (err) {
      context.log.error('Publish error', err);
      context.res = { status: 500, body: 'Publish failed' };
      return;
    }
  }

  context.res = { status: 202, body: 'Accepted' };
};
