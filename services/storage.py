# services/storage.py
import os
from typing import Dict, Optional
from datetime import datetime, timedelta
from urllib.parse import quote

from azure.storage.blob import BlobServiceClient, BlobSasPermissions, generate_blob_sas

from config import AZURE_CONN_STR, AZURITE_SAS_VERSION, LOCAL_SAVE_DIR, PUBLIC_BASE_URL


def parse_conn_str(conn: str) -> Dict[str, Optional[str]]:
    if not conn:
        return {"AccountName": None, "AccountKey": None, "BlobEndpoint": None}

    if "UseDevelopmentStorage=true" in conn:
        return {
            "AccountName": "devstoreaccount1",
            "AccountKey": (
                "Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsu"
                "Fq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw=="
            ),
            "BlobEndpoint": "http://127.0.0.1:10000/devstoreaccount1",
        }

    parts = dict(p.split("=", 1) for p in conn.split(";") if "=" in p)
    return {
        "AccountName": parts.get("AccountName"),
        "AccountKey": parts.get("AccountKey"),
        "BlobEndpoint": parts.get("BlobEndpoint"),
    }


def blob_service_client():
    if not AZURE_CONN_STR:
        raise RuntimeError("AZURE_STORAGE_CONNECTION_STRING not set")
    info = parse_conn_str(AZURE_CONN_STR)
    if info.get("BlobEndpoint") and info.get("AccountKey"):
        return BlobServiceClient(account_url=info["BlobEndpoint"], credential=info["AccountKey"])
    return BlobServiceClient.from_connection_string(AZURE_CONN_STR)


def upload_and_sas(container: str, blob_name: str, data: bytes, ttl_minutes: int = 120) -> str:
    if not AZURE_CONN_STR:
        raise RuntimeError("AZURE_STORAGE_CONNECTION_STRING not set")

    info = parse_conn_str(AZURE_CONN_STR)
    account_name = info["AccountName"]
    account_key = info["AccountKey"]
    blob_endpoint = info.get("BlobEndpoint")  # e.g. https://<acct>.blob.core.windows.net

    bsc = blob_service_client()
    cc = bsc.get_container_client(container)
    try:
        cc.create_container()
    except Exception:
        pass

    cc.upload_blob(name=blob_name, data=data, overwrite=True)

    # Build SAS (no extra encoding)
    sas_kwargs = dict(
        account_name=account_name,
        account_key=account_key,
        container_name=container,
        blob_name=blob_name,
        permission=BlobSasPermissions(read=True),
        # allow 5 min clock skew
        start=datetime.utcnow() - timedelta(minutes=5),
        expiry=datetime.utcnow() + timedelta(minutes=ttl_minutes),
    )

    # Only force http/version when running against Azurite
    if blob_endpoint and ("127.0.0.1" in blob_endpoint or "localhost" in blob_endpoint):
        sas_kwargs["protocol"] = "http"               # ok for Azurite
        if AZURITE_SAS_VERSION:
            sas_kwargs["version"] = AZURITE_SAS_VERSION

    sas = generate_blob_sas(**sas_kwargs)

    base = blob_endpoint.rstrip("/") if blob_endpoint else f"https://{account_name}.blob.core.windows.net"
    # Important: DO NOT quote/encode the SAS. It is already correctly encoded.
    # Optionally quote the blob path in case of spaces or special chars.
    return f"{base}/{container}/{quote(blob_name, safe='/')}?{sas}"


def save_local_and_url(blob_name: str, data: bytes) -> str:
    full_path = os.path.join(LOCAL_SAVE_DIR, blob_name)
    os.makedirs(os.path.dirname(full_path), exist_ok=True)
    with open(full_path, "wb") as f:
        f.write(data)
    return f"{PUBLIC_BASE_URL}/local/{blob_name}"
