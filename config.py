# config.py
import os

# FAC API Configuration
FAC_BASE = os.getenv("FAC_API_BASE", "https://api.fac.gov")
FAC_KEY = os.getenv("FAC_API_KEY")

# Azure Storage Configuration
AZURE_CONTAINER = os.getenv("AZURE_BLOB_CONTAINER", "mdl-output")
AZURE_CONN_STR = os.getenv("AZURE_STORAGE_CONNECTION_STRING")  # optional
AZURITE_SAS_VERSION = os.getenv("AZURITE_SAS_VERSION", "2021-08-06")

# Local Storage Configuration
LOCAL_SAVE_DIR = os.getenv("LOCAL_SAVE_DIR", "./_out")
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "http://localhost:8000")

# Template Configuration
MDL_TEMPLATE_PATH = os.getenv("MDL_TEMPLATE_PATH")

# Treasury Programs Mapping
TREASURY_PROGRAMS = {
    "21.027": "Coronavirus State and Local Fiscal Recovery Funds (SLFRF)",
    "21.023": "Emergency Rental Assistance Program (ERA)",
    "21.026": "Homeowner Assistance Fund (HAF)",
    "21.029": "Capital Projects Fund (CPF)",
    "21.031": "State Small Business Credit Initiative (SSBCI)",
    "21.032": "Local Assistance and Tribal Consistency Fund (LATCF)",
}

# Default lowercase words for name formatting
LOWERCASE_WORDS = {"and", "of", "the", "for", "to", "in", "on", "at", "by", "with", "from"}
