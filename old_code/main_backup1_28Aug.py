# ---------- NEW: Build DOCX directly from raw FAC arrays (no LLM needed) ----------
from pydantic import BaseModel
from typing import List, Dict, Any, Optional

class BuildFromFAC(BaseModel):
    auditee_name: str
    ein: str
    audit_year: int
    fac_general: List[Dict[str, Any]] = []
    fac_findings: List[Dict[str, Any]] = []
    fac_findings_text: List[Dict[str, Any]] = []
    fac_caps: List[Dict[str, Any]] = []
    federal_awards: List[Dict[str, Any]] = []
    dest_path: Optional[str] = None
    filename: Optional[str] = None

def _short(s: Optional[str], limit: int) -> str:
    if not s:
        return ""
    s = s.strip()
    return (s[:limit] + "…") if len(s) > limit else s

def _compose_html_from_fac(req: BuildFromFAC) -> str:
    # Take the first general row if present
    g = (req.fac_general[0] if req.fac_general else {})
    rid = g.get("report_id") or "N/A"
    fac_date = g.get("fac_accepted_date") or "N/A"

    # Maps for quick joins
    text_map = { (t.get("finding_ref_number") or t.get("reference_number")): (t.get("finding_text") or "") 
                 for t in (req.fac_findings_text or []) }
    cap_map  = { (c.get("finding_ref_number") or c.get("reference_number")): (c.get("planned_action") or "") 
                 for c in (req.fac_caps or []) }
    prog_map = { a.get("award_reference"): a.get("federal_program_name") 
                 for a in (req.federal_awards or []) if a.get("award_reference") }

    # Keep only "MDL-relevant" findings (any flag true)
    def relevant(row: Dict[str, Any]) -> bool:
        flags = [
            row.get("is_material_weakness"),
            row.get("is_significant_deficiency"),
            row.get("is_questioned_costs"),
            row.get("is_modified_opinion"),
            row.get("is_other_findings"),
            row.get("is_other_matters"),
            row.get("is_repeat_finding"),
        ]
        return any(bool(x) for x in flags)

    findings = [f for f in (req.fac_findings or []) if relevant(f)]
    # Sort & cap to avoid huge docs
    findings.sort(key=lambda x: str(x.get("reference_number") or ""))
    findings = findings[:50]

    # Build table rows
    rows_html = []
    for f in findings:
        ref = str(f.get("reference_number") or "")
        program = prog_map.get(f.get("award_reference")) or "N/A"
        req_type = f.get("type_requirement") or ""
        sev = []
        if f.get("is_material_weakness"): sev.append("material_weakness")
        if f.get("is_significant_deficiency"): sev.append("significant_deficiency")
        if f.get("is_questioned_costs"): sev.append("questioned_costs")
        if f.get("is_modified_opinion"): sev.append("modified_opinion")
        if f.get("is_repeat_finding"): sev.append("repeat_finding")
        if f.get("is_other_matters"): sev.append("other_matters")
        if f.get("is_other_findings"): sev.append("other_findings")
        sev_str = ", ".join(sev) if sev else "—"

        summary = _short(text_map.get(ref, ""), 900)   # keep it readable
        cap_txt = _short(cap_map.get(ref, ""), 400)

        rows_html.append(f"""
        <tr>
          <td>{ref}</td>
          <td>{program}</td>
          <td>{req_type}</td>
          <td>{sev_str}</td>
          <td>{summary}</td>
          <td>{cap_txt or ''}</td>
        </tr>
        """)

    total = len(findings)
    exec_summary = (f"{total} MDL-relevant finding(s) identified."
                    if total > 0 else "No MDL-relevant findings identified per FAC records.")

    html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>MDL (Preview) - {req.auditee_name} - {req.ein} - {req.audit_year}</title>
</head>
<body>
  <h1>Master Decision Letter – Preview (No-LLM)</h1>
  <p>
    <b>Auditee:</b> {req.auditee_name}<br/>
    <b>EIN:</b> {req.ein} &nbsp; <b>Audit Year:</b> {req.audit_year}<br/>
    <b>FAC Report ID:</b> {rid} &nbsp; <b>FAC Accepted:</b> {fac_date}<br/>
    <b>Date:</b> {datetime.utcnow().date().isoformat()}
  </p>

  <h2>Executive Summary</h2>
  <p>{exec_summary}</p>

  <h2>Findings (first {total} shown)</h2>
  <table border="1" cellspacing="0" cellpadding="6">
    <thead>
      <tr>
        <th>Ref #</th><th>Program</th><th>Requirement Type</th>
        <th>Severity</th><th>Summary (trimmed)</th><th>CAP (trimmed)</th>
      </tr>
    </thead>
    <tbody>
      {''.join(rows_html) if rows_html else '<tr><td colspan="6">None</td></tr>'}
    </tbody>
  </table>

  <div style="page-break-after: always;"></div>
  <h2>Appendix A — Raw counts</h2>
  <ul>
    <li>general rows: {len(req.fac_general or [])}</li>
    <li>findings rows (MDL-relevant): {len(findings)}</li>
    <li>findings_text rows: {len(req.fac_findings_text or [])}</li>
    <li>corrective_action_plans rows: {len(req.fac_caps or [])}</li>
    <li>federal_awards rows: {len(req.federal_awards or [])}</li>
  </ul>
</body>
</html>"""
    return html

@app.post("/build-docx-from-fac")
def build_docx_from_fac(req: BuildFromFAC):
    # Compose HTML from raw arrays
    html = _compose_html_from_fac(req)

    # Convert HTML -> DOCX (reuse your existing logic)
    docx_io = BytesIO()
    document = Document()
    html2docx(html, document)
    document.save(docx_io)
    data = docx_io.getvalue()

    container = os.getenv("AZURE_BLOB_CONTAINER", "mdl-output")
    folder = (req.dest_path or "").lstrip("/")
    base = req.filename or f"MDL-Preview-{sanitize(req.auditee_name)}-{sanitize(req.ein)}-{req.audit_year}.docx"
    blob_name = f"{folder}{base}" if folder else base

    # Prefer Azure/Azurite if configured, else save locally if you added the local fallback
    if os.getenv("AZURE_STORAGE_CONNECTION_STRING"):
        url = upload_and_sas(container, blob_name, data)
    elif 'save_local_and_url' in globals():
        url = save_local_and_url(blob_name, data)  # uses the optional local fallback if you added it earlier
    else:
        raise HTTPException(500, "No storage configured. Set AZURE_STORAGE_CONNECTION_STRING or add local fallback.")

    return {"url": url, "blob_path": f"{container}/{blob_name}", "size_bytes": len(data)}

