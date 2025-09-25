from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from datetime import datetime, timedelta
from io import BytesIO
import os, re, json
from html2docx import html2docx
from docx import Document
from azure.storage.blob import (
    BlobServiceClient, BlobSasPermissions, generate_blob_sas
)

app = FastAPI(title="MDL DOCX Builder")

# ---- Azure helpers ----
def _parse_conn_str(conn: str):
    # Extract AccountName/AccountKey from standard connection string
    parts = dict([p.split("=", 1) for p in conn.split(";") if "=" in p])
    return parts.get("AccountName"), parts.get("AccountKey")

def upload_and_sas(container: str, blob_name: str, data: bytes, ttl_minutes: int = 120):
    conn = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
    if not conn:
        raise RuntimeError("AZURE_STORAGE_CONNECTION_STRING not set")

    bsc = BlobServiceClient.from_connection_string(conn)
    cc = bsc.get_container_client(container)
    try:
        cc.create_container()
    except Exception:
        pass  # already exists

    cc.upload_blob(name=blob_name, data=data, overwrite=True)

    account_name, account_key = _parse_conn_str(conn)
    if not (account_name and account_key):
        raise RuntimeError("Could not parse account name/key from connection string")

    sas = generate_blob_sas(
        account_name=account_name,
        account_key=account_key,
        container_name=container,
        blob_name=blob_name,
        permission=BlobSasPermissions(read=True),
        expiry=datetime.utcnow() + timedelta(minutes=ttl_minutes),
    )
    url = f"https://{account_name}.blob.core.windows.net/{container}/{blob_name}?{sas}"
    return url

def sanitize(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_")

# ---- Models ----
class BuildDocx(BaseModel):
    auditee_name: str
    ein: str
    audit_year: int
    body_html: str | None = None  # preferred
    dest_path: str | None = None  # e.g., "mdl/2024/"
    filename: str | None = None   # optional override

class ComposeAndBuild(BaseModel):
    auditee_name: str
    ein: str
    audit_year: int
    mdl: dict         # normalized JSON from the flow (Block 6 output)
    treasury_template: str | None = None  # optional (server-side compose)
    dest_path: str | None = None
    filename: str | None = None

# ---- Minimal JSON->HTML (server-side) ----
def mdl_json_to_html(mdl: dict) -> str:
    aud = mdl.get("auditee", {})
    findings = mdl.get("normalized_findings", []) or []
    has_findings = mdl.get("has_findings", False)

    def esc(s): return (s or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

    rows = []
    for f in findings:
        rows.append(f"""
          <tr>
            <td>{esc(f.get('finding_ref'))}</td>
            <td>{esc(f.get('program_name') or 'N/A')}</td>
            <td>{esc(f.get('type_requirement') or '')}</td>
            <td>{esc(", ".join(f.get('severity_flags') or []))}</td>
            <td>{esc(f.get('summary') or '')}</td>
            <td>{esc(f.get('cap_excerpt') or '')}</td>
            <td>FAC findings_text {esc(f.get('finding_ref') or '')}</td>
          </tr>
        """)

    summary = (f"{len(findings)} MDL-relevant finding(s) identified."
               if has_findings else "No MDL-relevant findings identified per FAC records.")
    today = datetime.utcnow().date().isoformat()

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>MDL - {esc(aud.get('name') or '')} - {esc(aud.get('ein') or '')} - {aud.get('audit_year') or ''}</title></head>
<body>
<h1>Master Decision Letter (MDL)</h1>
<p><b>Auditee:</b> {esc(aud.get('name') or '')}<br/>
<b>EIN:</b> {esc(aud.get('ein') or '')} &nbsp; <b>Audit Year:</b> {esc(str(aud.get('audit_year') or ''))}<br/>
<b>FAC Report ID:</b> {esc(aud.get('report_id') or 'N/A')} &nbsp; <b>FAC Accepted:</b> {esc(aud.get('fac_accepted_date') or 'N/A')}<br/>
<b>Date:</b> {today}</p>

<h2>Executive Summary</h2>
<p>{esc(summary)}</p>

<h2>Findings</h2>
<table border="1" cellspacing="0" cellpadding="6">
  <thead>
    <tr>
      <th>Ref #</th><th>Program</th><th>Requirement Type</th><th>Severity</th><th>Summary</th><th>CAP</th><th>Sources</th>
    </tr>
  </thead>
  <tbody>
    {''.join(rows)}
  </tbody>
</table>

<div style="page-break-after: always;"></div>
<h2>Appendix A — Data sources and keys</h2>
<p>Key: {esc(aud.get('name') or '')} | {esc(aud.get('ein') or '')} | {esc(str(aud.get('audit_year') or ''))}</p>
<p>Data derived from FAC API tables: general, findings, findings_text, corrective_action_plans.</p>
</body>
</html>"""
    return html

# ---- Routes ----
@app.get("/healthz")
def healthz():
    return {"ok": True, "time": datetime.utcnow().isoformat()}

@app.post("/build-docx")
def build_docx(req: BuildDocx):
    if not (req.body_html and req.body_html.strip()):
        raise HTTPException(400, "body_html is required (use /compose-and-build for JSON input)")

    # Convert HTML -> DOCX
    docx_io = BytesIO()
    document = Document()
    html2docx(req.body_html, document)
    document.save(docx_io)
    data = docx_io.getvalue()

    # Naming & upload
    container = os.getenv("AZURE_BLOB_CONTAINER", "mdl-output")
    folder = (req.dest_path or "").lstrip("/")
    base = req.filename or f"MDL-{sanitize(req.auditee_name)}-{sanitize(req.ein)}-{req.audit_year}.docx"
    blob_name = f"{folder}{base}" if folder else base

    url = upload_and_sas(container, blob_name, data)
    return {"url": url, "blob_path": f"{container}/{blob_name}", "size_bytes": len(data)}

@app.post("/compose-and-build")
def compose_and_build(req: ComposeAndBuild):
    # Compose HTML from normalized JSON (and optional template hint)
    try:
        html = mdl_json_to_html(req.mdl)
    except Exception as e:
        raise HTTPException(400, f"Failed to compose HTML from JSON: {e}")

    # Reuse build pipeline
    build = BuildDocx(
        auditee_name=req.auditee_name,
        ein=req.ein,
        audit_year=req.audit_year,
        body_html=html,
        dest_path=req.dest_path,
        filename=req.filename
    )
    return build_docx(build)

