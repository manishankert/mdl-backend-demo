#!/usr/bin/env bash
set -euo pipefail

# --- Your subscription & resource group (from you) ---
SUBSCRIPTION_ID="5721f014-3c71-4a1e-a4c0-69eecafe8ded"
RG_NAME="rg-aip-dev-eus"
LOCATION="eastus"

# --- Editable names (must be unique where noted) ---
STG="stmdl$RANDOM"                 # storage account: lowercase, 3-24 chars, globally unique
CONT_OUTPUT="mdl-output"           # container for generated MDL docs
CONT_TEMPLATES="mdl-templates"     # optional: if you store templates in blob

PLAN="plan-mdl-dev"
APP="mdl-aip-dev-eus"              # must be unique within azurewebsites.net
PY="3.11"

echo ">> Set subscription"
az account set --subscription "$SUBSCRIPTION_ID"

echo ">> Ensure resource group exists: $RG_NAME"
az group create -n "$RG_NAME" -l "$LOCATION" 1>/dev/null

echo ">> Create storage account: $STG"
az storage account create -n "$STG" -g "$RG_NAME" -l "$LOCATION" --sku Standard_LRS --kind StorageV2 1>/dev/null

echo ">> Create blob containers"
az storage container create --name "$CONT_OUTPUT"   --account-name "$STG" --auth-mode login 1>/dev/null
az storage container create --name "$CONT_TEMPLATES" --account-name "$STG" --auth-mode login 1>/dev/null || true

echo ">> Create Linux App Service plan"
az appservice plan create -g "$RG_NAME" -n "$PLAN" --sku B1 --is-linux 1>/dev/null

echo ">> Create Web App (Python)"
az webapp create -g "$RG_NAME" -p "$PLAN" -n "$APP" --runtime "PYTHON:${PY}" 1>/dev/null

echo ">> Enable System Assigned Managed Identity"
az webapp identity assign -g "$RG_NAME" -n "$APP" 1>/dev/null

# Grant MI access to the storage account
PRINCIPAL_ID=$(az webapp show -g "$RG_NAME" -n "$APP" --query identity.principalId -o tsv)
STG_ID=$(az storage account show -g "$RG_NAME" -n "$STG" --query id -o tsv)
echo ">> Assign 'Storage Blob Data Contributor' to MI on storage"
az role assignment create --assignee "$PRINCIPAL_ID" --role "Storage Blob Data Contributor" --scope "$STG_ID" 1>/dev/null || true

# --- App settings you’ll use in code ---
echo ">> Configure App Settings"
az webapp config appsettings set -g "$RG_NAME" -n "$APP" --settings \
  APP_ENV=dev \
  AZURE_STORAGE_ACCOUNT="$STG" \
  BLOB_CONTAINER="$CONT_OUTPUT" \
  USE_MANAGED_IDENTITY=true \
  TREASURY_CONTACT_EMAIL=ORP_SingleAudits@treasury.gov \
  OPENAI_API_KEY="" \
  FAC_API_KEY="" \
  DOCX_API_KEY="" 1>/dev/null

# Startup command for gunicorn + uvicorn worker (adjust module if not 'main:app')
echo ">> Configure startup command"
az webapp config set -g "$RG_NAME" -n "$APP" \
  --startup-file "gunicorn -k uvicorn.workers.UvicornWorker main:app --timeout 180 --workers 2 --threads 4 --bind=0.0.0.0" 1>/dev/null

echo ">> Done"
echo "Web App URL: https://${APP}.azurewebsites.net"
