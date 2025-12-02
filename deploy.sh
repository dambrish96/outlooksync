#!/usr/bin/env bash
set -euo pipefail

########################################
# EDIT THESE VALUES
########################################

SUBSCRIPTION_ID="<YOUR-SUBSCRIPTION-ID>"
RESOURCE_GROUP="rg-graph-eg-test"
LOCATION="eastus"

# Function App / Storage names (must be unique globally)
FUNCTION_APP_NAME="fn-graph-webhook-$RANDOM"
STORAGE_ACCOUNT_NAME="stgraphwebhook$RANDOM"
APP_SERVICE_PLAN="asp-graph-$RANDOM"
INSIGHTS_NAME="${FUNCTION_APP_NAME}-ai"

# ARM template file
ARM_TEMPLATE="azuredeploy.json"

# Event Grid partner placeholder names
PARTNER_REG_NAME="graphPartnerReg"
PARTNER_TOPIC_NAME="graphPartnerTopic"

# >>> IMPORTANT: Mailbox to monitor <<<
MAILBOX_USER_ID="user@domain.com"
########################################



echo "--------------------------------------------"
echo " Setting Azure subscription"
echo "--------------------------------------------"
az account set --subscription "$SUBSCRIPTION_ID"


echo "--------------------------------------------"
echo " Creating Resource Group"
echo "--------------------------------------------"
az group create -n "$RESOURCE_GROUP" -l "$LOCATION"


echo "--------------------------------------------"
echo " Deploying ARM template"
echo "--------------------------------------------"
az deployment group create \
  --resource-group "$RESOURCE_GROUP" \
  --template-file "$ARM_TEMPLATE" \
  --parameters functionAppName="$FUNCTION_APP_NAME" \
               storageAccountName="$STORAGE_ACCOUNT_NAME" \
               appServicePlanName="$APP_SERVICE_PLAN" \
               aiName="$INSIGHTS_NAME" \
               partnerRegistrationName="$PARTNER_REG_NAME" \
               partnerTopicName="$PARTNER_TOPIC_NAME" \
               location="$LOCATION"


echo "--------------------------------------------"
echo " Creating Azure AD App Registration"
echo "--------------------------------------------"
APP_DISPLAY_NAME="graph-calendar-webhook-$RANDOM"
APP_ID=$(az ad app create \
    --display-name "$APP_DISPLAY_NAME" \
    --query appId -o tsv)

echo "App Registered:"
echo "  App ID: $APP_ID"


echo "--------------------------------------------"
echo " Creating Client Secret"
echo "--------------------------------------------"
CLIENT_SECRET=$(az ad app credential reset \
    --id "$APP_ID" \
    --query password -o tsv)

TENANT_ID=$(az account show --query tenantId -o tsv)

echo "Client Secret created."


echo "--------------------------------------------"
echo " Adding Graph API Permissions"
echo "--------------------------------------------"
az ad app permission add \
  --id "$APP_ID" \
  --api 00000003-0000-0000-c000-000000000000 \
  --api-permissions "Calendars.Read=Scope" "Calendars.ReadWrite=Scope"

echo ""
echo "!!! ACTION REQUIRED !!!"
echo "Open Azure Portal → App Registrations → $APP_DISPLAY_NAME"
echo "Go to API Permissions → click **Grant admin consent**"
echo ""
read -p "Press ENTER after granting consent..."


echo "--------------------------------------------"
echo " Getting Function Hostname"
echo "--------------------------------------------"
FUNC_HOSTNAME=$(az webapp show -g "$RESOURCE_GROUP" -n "$FUNCTION_APP_NAME" --query defaultHostName -o tsv)
echo "Function Hostname: $FUNC_HOSTNAME"


echo "--------------------------------------------"
echo " Getting Function Master Key"
echo "--------------------------------------------"
MASTER_KEY=$(az functionapp keys list \
    --resource-group "$RESOURCE_GROUP" \
    --name "$FUNCTION_APP_NAME" \
    --query masterKey -o tsv)

FUNCTION_URL="https://${FUNC_HOSTNAME}/api/HttpTrigger?code=${MASTER_KEY}"

echo "Function URL: $FUNCTION_URL"


echo "--------------------------------------------"
echo " Creating Dev Event Grid Topic"
echo "--------------------------------------------"
DEV_TOPIC="evtTopic$RANDOM"
az eventgrid topic create -g "$RESOURCE_GROUP" -n "$DEV_TOPIC" -l "$LOCATION"

TOPIC_ENDPOINT=$(az eventgrid topic show -g "$RESOURCE_GROUP" -n "$DEV_TOPIC" --query endpoint -o tsv)
TOPIC_KEY=$(az eventgrid topic key list -g "$RESOURCE_GROUP" --name "$DEV_TOPIC" --query key1 -o tsv)

echo "Event Grid Topic: $DEV_TOPIC"
echo "Topic Endpoint: $TOPIC_ENDPOINT"


echo "--------------------------------------------"
echo " Setting Function App Settings"
echo "--------------------------------------------"
az webapp config appsettings set \
  -g "$RESOURCE_GROUP" \
  -n "$FUNCTION_APP_NAME" \
  --settings \
    GRAPH_APP_CLIENT_ID="$APP_ID" \
    GRAPH_APP_CLIENT_SECRET="$CLIENT_SECRET" \
    GRAPH_TENANT_ID="$TENANT_ID" \
    EVENTGRID_TOPIC_ENDPOINT="$TOPIC_ENDPOINT" \
    EVENTGRID_TOPIC_KEY="$TOPIC_KEY" \
    ENCRYPTION_CERT_ID="v1"


echo "--------------------------------------------"
echo " Deploying Function Code"
echo "--------------------------------------------"
zip -r function.zip function > /dev/null
az functionapp deployment source config-zip \
  -g "$RESOURCE_GROUP" \
  -n "$FUNCTION_APP_NAME" \
  --src function.zip


echo "--------------------------------------------"
echo " Generating Graph Token"
echo "--------------------------------------------"
TOKEN=$(curl -s -X POST \
  -H "Content-Type: application/x-www-form-urlencoded" \
  -d "client_id=$APP_ID&scope=https://graph.microsoft.com/.default&client_secret=$CLIENT_SECRET&grant_type=client_credentials" \
  "https://login.microsoftonline.com/$TENANT_ID/oauth2/v2.0/token" | jq -r .access_token)


echo "--------------------------------------------"
echo " Creating Graph Subscription for ${MAILBOX_USER_ID}"
echo "--------------------------------------------"

RESOURCE="/users/${MAILBOX_USER_ID}/events"
EXPIRY=$(date -u -d "48 hours" +"%Y-%m-%dT%H:%M:%SZ")

curl -X POST \
  "https://graph.microsoft.com/v1.0/subscriptions" \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d "{
    \"changeType\": \"created,updated,deleted\",
    \"notificationUrl\": \"${FUNCTION_URL}\",
    \"resource\": \"${RESOURCE}\",
    \"expirationDateTime\": \"${EXPIRY}\",
    \"clientState\": \"secretStateValue\"
  }" | jq .


echo "--------------------------------------------"
echo " DEPLOYMENT COMPLETE"
echo "--------------------------------------------"
echo "Function URL: $FUNCTION_URL"
echo "Mailbox subscription created for: $MAILBOX_USER_ID"
echo "Monitor logs using:"
echo "az webapp log tail -g $RESOURCE_GROUP -n $FUNCTION_APP_NAME"
