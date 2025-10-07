# main.py
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Tuple
from io import BytesIO
from urllib.parse import quote
import os, re, base64, html, requests
from lxml import etree
import os, openpyxl
from docx.oxml.ns import qn
import re
import json
# DOCX / HTML
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from bs4 import BeautifulSoup
from html2docx import HTML2Docx
import os, json, requests
import logging
logging.basicConfig(level=logging.INFO)

# Azure
from azure.storage.blob import BlobServiceClient, BlobSasPermissions, generate_blob_sas

# ------------------------------------------------------------------------------
# FastAPI app
# ------------------------------------------------------------------------------
app = FastAPI(title="MDL DOCX Builder")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ------------------------------------------------------------------------------
# Environment
# ------------------------------------------------------------------------------
FAC_BASE = os.getenv("FAC_API_BASE", "https://api.fac.gov")
FAC_KEY  = os.getenv("FAC_API_KEY")

AZURE_CONTAINER = os.getenv("AZURE_BLOB_CONTAINER", "mdl-output")
AZURE_CONN_STR  = os.getenv("AZURE_STORAGE_CONNECTION_STRING")  # optional
AZURITE_SAS_VERSION = os.getenv("AZURITE_SAS_VERSION", "2021-08-06")

LOCAL_SAVE_DIR = os.getenv("LOCAL_SAVE_DIR", "./_out")
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "http://localhost:8000")

MDL_TEMPLATE_PATH = os.getenv("MDL_TEMPLATE_PATH")

# ------------------------------------------------------------------------------
# Utilities
# ------------------------------------------------------------------------------
def sanitize(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name or "").strip("_")

def _short(s: Optional[str], limit: int = 900) -> str:
    if not s:
        return ""
    s = re.sub(r"\s+", " ", s.strip())
    return (s[: limit - 1] + "…") if len(s) > limit else s

def _norm_ref(x: Optional[str]) -> str:
    return re.sub(r"\s+", "", (x or "")).upper()

def _shade_cell(cell, hex_fill="E7E6E6"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)

def _set_col_widths(table: Table, widths):
    for col_idx, w in enumerate(widths):
        for cell in table.columns[col_idx].cells:
            cell.width = w

def _tight_paragraph(p: Paragraph):
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)

def _as_oxml(el):
    """Return underlying oxml element for Paragraph/Table/raw CT_* safely."""
    if hasattr(el, "_p"):   # Paragraph
        return el._p
    if hasattr(el, "_tbl"): # Table
        return el._tbl
    if hasattr(el, "_element"):
        return el._element
    return el  # assume already oxml

def _insert_after(anchor, new_block):
    """Insert new_block (Paragraph/Table or raw oxml) after anchor (Paragraph/Table or raw oxml)."""
    a = _as_oxml(anchor)
    n = _as_oxml(new_block)
    a.addnext(n)

def _apply_grid_borders(tbl: Table):
    """Ensure visible borders regardless of style availability."""
    tbl_el = tbl._tbl
    tblPr = tbl_el.tblPr or tbl_el.get_or_add_tblPr()
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "6")     # 0.5pt
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        borders.append(e)
    tblPr.append(borders)

def _title_case(s: str) -> str:
    if not s:
        return ""
    return s.title()

def _remove_paragraph(p):
    # safe remove of a docx paragraph
    p._element.getparent().remove(p._element)

def _norm_txt(s: str) -> str:
    if not s:
        return ""
    # normalize NBSP, dashes, whitespace
    s = s.replace("\u00A0", " ").replace("\xa0", " ")
    s = s.replace("–", "-").replace("—", "-")
    return " ".join(s.split())

def _title_with_article(name: str) -> str:
    if not name:
        return ""
    return name if name.lower().startswith("the ") else f"The {name}"

def _from_fac_general(gen: List[Dict[str, Any]]) -> Dict[str, str]:
    """
    Pull best-effort defaults from FAC 'general' row.
    We tolerate missing columns—return what we can.
    """
    if not gen:
        return {}
    g = gen[0] or {}

    # FAC fields vary slightly across vintages; try common variants.
    addr1 = g.get("auditee_address_line_1") or g.get("auditee_address1") or ""
    city  = g.get("auditee_city") or g.get("city") or ""
    state = g.get("auditee_state") or g.get("state") or ""
    zipc  = (g.get("auditee_zip") or g.get("zip_code") or "").strip()

    auditor = g.get("auditor_firm_name") or g.get("auditor_name") or ""

    # Period end text if present; fall back to just year elsewhere
    fy_end = g.get("fy_end_text") or g.get("fy_end_date") or g.get("fiscal_year_end") or ""

    return {
        "street_address": addr1,
        "city": city,
        "state": state,
        "zip_code": zipc,
        "auditor_name": auditor,
        "period_end_text": fy_end
    }

def _cleanup_post_table_narrative(doc, model):
    """
    Remove the repeated narrative paragraphs that appear after the program table(s):
      - Lines starting with the finding id (e.g., '2024-002 – ...')
      - Auditor Description..., Auditor Recommendation., Responsible Person:, Corrective Action., Anticipated Completion Date:
      - Lines duplicating the raw finding summary text
    """
    # Collect IDs and summaries to match
    finding_ids = set()
    summaries = set()
    combos = set()
    for prog in (model.get("programs") or []):
        for f in (prog.get("findings") or []):
            fid = (f.get("finding_id") or "").strip()
            summ = (f.get("summary") or "").strip()
            combo = (f.get("compliance_and_summary") or "").strip()
            if fid: finding_ids.add(fid)
            if summ: summaries.add(_norm_txt(summ))
            if combo: combos.add(_norm_txt(combo))

    # Regex patterns that match the repeated narrative blocks in the body
    starts = [
        r"^\d{4}-\d{3}\s*-\s*",                        # e.g., 2024-002 -
        r"^\d{4}-\d{3}\s*[–—]\s*",                     # e.g., 2024-002 – (en/em dash)
        r"^Auditor\s+Description\s+of\s+Condition",    # Auditor Description of Condition...
        r"^Auditor\s+Recommendation\.?",               # Auditor Recommendation.
        r"^Responsible\s+Person\s*:",                  # Responsible Person:
        r"^Corrective\s+Action\.?",                    # Corrective Action.
        r"^Anticipated\s+Completion\s+Date\s*:",       # Anticipated Completion Date:
    ]
    patt = re.compile("|".join(starts), re.IGNORECASE)

    # Remove paragraphs that match any of the above
    for p in list(doc.paragraphs):
        t = _norm_txt("".join(r.text for r in p.runs))
        if not t:
            continue

        # Exact/contains matches
        if any(fid in t for fid in finding_ids):
            _remove_paragraph(p); continue

        if patt.search(t):
            _remove_paragraph(p); continue

        nt = _norm_txt(t)
        if any(s and s.lower() in nt.lower() for s in summaries):
            _remove_paragraph(p); continue

        if any(c and c.lower() in nt.lower() for c in combos):
            _remove_paragraph(p); continue

# def _pluralize_text(doc, total_findings: int):
#     """
#     Replace tokens like finding(s), CAP(s), issue(s), violate(s) etc.
#     with singular or plural forms depending on total_findings.
#     """
#     singular = (total_findings == 1)

#     replacements = {
#         "audit finding(s)": "audit finding" if singular else "audit findings",
#         "finding(s)": "finding" if singular else "findings",
#         "issue(s)": "issue" if singular else "issues",
#         "violate(s)": "violates" if singular else "violate",
#         "CAP(s)": "CAP" if singular else "CAPs",
#         "address(es)": "addresses" if singular else "address",
#         "date(s)": "date" if singular else "dates",
#         "corrective action(s)": "corrective action" if singular else "corrective actions",
#         "appear(s)": "appears" if singular else "appear",
#     }

#     def _replace_in_para(p):
#         for run in p.runs:
#             text = run.text
#             for k, v in replacements.items():
#                 if k in text:
#                     run.text = text.replace(k, v)

#     for p in doc.paragraphs:
#         _replace_in_para(p)

#     # also fix inside tables if tokens appear there
#     for tbl in doc.tables:
#         for row in tbl.rows:
#             for cell in row.cells:
#                 for p in cell.paragraphs:
#                     _replace_in_para(p)

import re as _re

def _rewrite_para_text(p, new_text: str):
    """Clear all runs in a paragraph and set to new_text."""
    for r in list(p.runs):
        r.clear()  # python-docx 1.1+; if older, do r._element.getparent().remove(r._element)
    p._element.clear_content()  # older-safe: remove content, keep properties
    p.add_run(new_text)

def _get_para_text(p) -> str:
    return "".join(r.text for r in p.runs)

def _pluralize_string(s: str, singular: bool) -> str:
    mapping_singular = {
        "audit finding(s) sustained": "The audit finding is sustained",
        "issue(s) violate(s)": "issue violates",
        "cap(s), if implemented,  responsive": "The CAP, if implemented, is responsive",
        "address(es) the cause": "addresses the cause",
        "date(s) indicated": "date indicated",
        "corrective action(s)  subject": "The corrective action is subject",
        "audit finding(s) appear(s)": "audit finding appears",
    }

    mapping_plural = {
        "audit finding(s) sustained": "The audit findings are sustained",
        "issue(s) violate(s)": "issues violate",
        "cap(s), if implemented,  responsive": "The CAPs, if implemented, are responsive",
        "address(es) the cause": "address the causes",
        "date(s) indicated": "dates indicated",
        "corrective action(s)  subject": "The corrective actions are subject",
        "audit finding(s) appear(s)": "audit findings appear",
    }

    mapping = mapping_singular if singular else mapping_plural

    out = s
    for k, v in mapping.items():
        if k in out:
            out = out.replace(k, v)
    return out

def _rewrite_para_text(p, new_text: str):
    # Clear all runs and set new clean text
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.add_run(new_text)

def _para_text(p) -> str:
    return "".join(r.text for r in p.runs)

def _looks_like_optional_plural_text(s: str) -> bool:
    """Find paragraphs that still have '(s)' or '(es)' style tokens or subject-verb '(s)'. """
    s = (s or "")
    return any(t in s for t in ["(s)", "(es)", "violate(s)", "address(es)", "appear(s)"]) or " audit finding" in s.lower() or " corrective action" in s.lower()

def _pluralize_with_openai(text: str, total_findings: int) -> Optional[str]:
    """
    Use OpenAI to convert optional-plural boilerplate into grammatically correct text.
    Returns rewritten string, or None on failure.
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        return None

    singular = (total_findings == 1)
    style_hint = (
        "Use singular grammar (is/addresses/appears; finding, issue, CAP, date, corrective action)."
        if singular else
        "Use plural grammar (are/address/appear; findings, issues, CAPs, dates, corrective actions)."
    )

    system = (
        "You are revising boilerplate text in a U.S. government letter. "
        "Rewrite the provided text to be grammatically correct and natural, "
        "resolving any optional plural tokens like '(s)' or '(es)' and fixing subject–verb agreement. "
        "Preserve meaning and tone; do not add or remove content beyond grammar and number agreement. "
        "Return only the final sentence(s) with no quotes."
    )
    user = (
        f"{style_hint}\n\n"
        "Rewrite the text below to be grammatically correct. Resolve all '(s)' / '(es)' tokens and subject–verb forms. "
        "Keep the same information, formal tone, and punctuation.\n\n"
        f"Text:\n{text}"
    )

    try:
        resp = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            data=json.dumps({
                "model": "gpt-4o-mini",
                "messages": [
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
                "temperature": 0,
            }),
            timeout=15,
        )
        resp.raise_for_status()
        out = resp.json()
        rewritten = (out.get("choices", [{}])[0].get("message", {}).get("content") or "").strip()
        return rewritten or None
    except Exception:
        return None

def _ai_fix_pluralization_in_doc(doc, total_findings: int):
    """
    Find paragraphs with '(s)/(es)' style text or affected phrases and fix them via OpenAI.
    Falls back silently if API not available.
    """
    candidates = []
    # Scan body paragraphs
    for p in doc.paragraphs:
        t = _para_text(p)
        if _looks_like_optional_plural_text(t):
            candidates.append(p)
    # Also scan header/footer just in case
    for sec in doc.sections:
        for container in (sec.header, sec.footer):
            for p in container.paragraphs:
                t = _para_text(p)
                if _looks_like_optional_plural_text(t):
                    candidates.append(p)

    # Rewrite each candidate via OpenAI; if it fails, leave as-is
    for p in candidates:
        original = _para_text(p).strip()
        if not original:
            continue
        rewritten = _pluralize_with_openai(original, total_findings)
        if rewritten and rewritten != original:
            _rewrite_para_text(p, rewritten)


# === Finding types & summary mapping (from Excel) ===
def _load_finding_mappings(xlsx_path: Optional[str]):
    """
    Returns:
      - type_map: {'I': 'Procurement and suspension and debarment', ...}
      - summary_labels: ['Lack of evidence of suspension and debarment verification', ...]
    Tolerant to header naming; no-op if workbook missing.
    """
    type_map, summary_labels = {}, []
    if not xlsx_path:
        return type_map, summary_labels
    try:
        import openpyxl
    except Exception:
        return type_map, summary_labels

    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    except Exception:
        return type_map, summary_labels

    def _find_sheet(*names):
        for ws in wb.worksheets:
            nm = (ws.title or "").strip().lower().replace(" ", "_")
            for want in names:
                if want in nm:
                    return ws
        return None

    # 1) Finding Types sheet
    ws_types = _find_sheet("finding_types", "findingtype", "finding_types_sheet", "types")
    if ws_types and ws_types.max_row >= 2:
        hdrs = [ (c.value or "") for c in ws_types[1] ]
        hl = [str(h).strip().lower() for h in hdrs]

        def colidx(cands):
            for i,h in enumerate(hl):
                for c in cands:
                    c = c.lower()
                    if h == c or c in h:
                        return i
            return None

        i_code = colidx(["code","compliance type","compliance_type","ctype"])
        i_name = colidx(["name","description","label","type name","type"])
        if i_code is not None and i_name is not None:
            for row in ws_types.iter_rows(min_row=2, values_only=True):
                code = (row[i_code] or "")
                name = (row[i_name] or "")
                code = str(code).strip().upper()
                name = str(name).strip()
                if code and name:
                    type_map[code] = name

    # 2) Finding_summaries sheet
    ws_summ = _find_sheet("finding_summaries", "finding_summ", "summaries", "summary")
    if ws_summ and ws_summ.max_row >= 2:
        # Use first non-empty cell in each row (or a column named 'summary' / 'label')
        hdrs = [ (c.value or "") for c in ws_summ[1] ]
        hl = [str(h).strip().lower() for h in hdrs]
        def cidx(cands):
            for i,h in enumerate(hl):
                for c in cands:
                    c = c.lower()
                    if h == c or c in h:
                        return i
            return None
        i_lbl = cidx(["summary","label","finding summary","finding_label"]) or 0
        for row in ws_summ.iter_rows(min_row=2, values_only=True):
            cell = row[i_lbl] if i_lbl < len(row) else None
            txt = str(cell or "").strip()
            if txt:
                summary_labels.append(txt)

    return type_map, summary_labels


def _best_summary_label(summary: str, labels: List[str]) -> Optional[str]:
    """
    Offline fuzzy match: pick the label with highest similarity to the summary.
    """
    if not summary or not labels:
        return None
    import difflib
    cand = difflib.get_close_matches(summary, labels, n=1, cutoff=0.0)
    if cand:
        return cand[0]
    return None

# ------------------------------------------------------------------------------
# Storage helpers
# ------------------------------------------------------------------------------
def _parse_conn_str(conn: str) -> Dict[str, Optional[str]]:
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

def _blob_service_client():
    if not AZURE_CONN_STR:
        raise RuntimeError("AZURE_STORAGE_CONNECTION_STRING not set")
    info = _parse_conn_str(AZURE_CONN_STR)
    if info.get("BlobEndpoint") and info.get("AccountKey"):
        return BlobServiceClient(account_url=info["BlobEndpoint"], credential=info["AccountKey"])
    return BlobServiceClient.from_connection_string(AZURE_CONN_STR)

def _best_summary_label_openai(summary: str, labels: List[str]) -> Optional[str]:
    import os, json, requests
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key or not labels:
        return None
    prompt = {
        "summary": summary,
        "labels": labels,
        "task": "Pick exactly one label from 'labels' that best matches 'summary'. Respond with just the label text."
    }
    try:
        r = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            data=json.dumps({
                "model": "gpt-4o-mini",
                "messages": [{"role":"user","content": json.dumps(prompt)}],
                "temperature": 0
            }),
            timeout=12,
        )
        out = r.json()
        txt = (out.get("choices",[{}])[0].get("message",{}).get("content") or "").strip()
        if txt in labels:
            return txt
    except Exception:
        pass
    return None

# def upload_and_sas(container: str, blob_name: str, data: bytes, ttl_minutes: int = 120) -> str:
#     if not AZURE_CONN_STR:
#         raise RuntimeError("AZURE_STORAGE_CONNECTION_STRING not set")

#     info = _parse_conn_str(AZURE_CONN_STR)
#     account_name = info["AccountName"]
#     account_key  = info["AccountKey"]
#     blob_endpoint = info.get("BlobEndpoint")

#     bsc = _blob_service_client()
#     cc = bsc.get_container_client(container)
#     try:
#         cc.create_container()
#     except Exception:
#         pass
#     cc.upload_blob(name=blob_name, data=data, overwrite=True)

#     proto = "http" if (blob_endpoint and ("127.0.0.1" in blob_endpoint or "localhost" in blob_endpoint)) else None
#     sas = generate_blob_sas(
#         account_name=account_name,
#         account_key=account_key,
#         container_name=container,
#         blob_name=blob_name,
#         permission=BlobSasPermissions(read=True),
#         expiry=datetime.utcnow() + timedelta(minutes=ttl_minutes),
#         version=AZURITE_SAS_VERSION,
#         protocol=proto,
#     )
#     sas_q = quote(sas, safe="=&")

#     base = blob_endpoint.rstrip("/") if blob_endpoint else f"https://{account_name}.blob.core.windows.net"
#     return f"{base}/{container}/{blob_name}?{sas_q}"


from urllib.parse import quote
from datetime import datetime, timedelta
from azure.storage.blob import BlobSasPermissions, generate_blob_sas

def upload_and_sas(container: str, blob_name: str, data: bytes, ttl_minutes: int = 120) -> str:
    if not AZURE_CONN_STR:
        raise RuntimeError("AZURE_STORAGE_CONNECTION_STRING not set")

    info = _parse_conn_str(AZURE_CONN_STR)
    account_name  = info["AccountName"]
    account_key   = info["AccountKey"]
    blob_endpoint = info.get("BlobEndpoint")  # e.g. https://<acct>.blob.core.windows.net

    bsc = _blob_service_client()
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

def _title_with_acronyms(s: str, keep_all_caps=True) -> str:
    # simple title-caser with stop-words; preserves ALL-CAPS tokens and acronyms in ( )
    lowers = {"and","or","the","of","for","to","in","on","by","with","a","an"}
    parts = []
    for word in s.split():
        base = word.strip()
        if keep_all_caps and base.isupper() and len(base) > 1:
            parts.append(base)
        else:
            w = base.lower()
            if w in lowers:
                parts.append(w)
            else:
                parts.append(w.capitalize())
    return " ".join(parts)

@app.get("/local/{path:path}")
def get_local_file(path: str):
    full = os.path.join(LOCAL_SAVE_DIR, path)
    if not os.path.isfile(full):
        raise HTTPException(404, "Not found")
    return FileResponse(
        full,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# ------------------------------------------------------------------------------
# FAC helpers
# ------------------------------------------------------------------------------
def _fac_headers():
    key = os.getenv("FAC_API_KEY")
    if not key:
        raise HTTPException(500, "FAC_API_KEY not configured on the docx service")
    return {"X-Api-Key": key}

def _fac_get(path: str, params: Dict[str, Any]) -> Any:
    try:
        r = requests.get(f"{FAC_BASE.rstrip('/')}/{path.lstrip('/')}",
                         headers=_fac_headers(), params=params, timeout=20)
        r.raise_for_status()
        return r.json()
    except requests.HTTPError as e:
        raise HTTPException(r.status_code if 'r' in locals() else 500,
                            f"FAC GET {path} failed: {getattr(r,'text','')}") from e
    except Exception as e:
        raise HTTPException(500, f"FAC GET {path} failed: {e}") from e

def _or_param(field: str, values: List[str]) -> str:
    inner = ",".join([f"{field}.eq.{v}" for v in values])
    return f"({inner})"

# ------------------------------------------------------------------------------
# HTML → DOCX (preview)
# ------------------------------------------------------------------------------
def _apply_inline_formatting(paragraph, node):
    def add_text(text, bold=False, italic=False, underline=False):
        if text is None:
            return
        r = paragraph.add_run(text)
        r.bold = bool(bold)
        r.italic = bool(italic)
        r.underline = bool(underline)

    for child in getattr(node, "children", []):
        name = getattr(child, "name", None)
        if name is None:
            add_text(str(child))
        elif name in ("b", "strong"):
            add_text(child.get_text(), bold=True)
        elif name in ("i", "em"):
            add_text(child.get_text(), italic=True)
        elif name == "u":
            add_text(child.get_text(), underline=True)
        elif name == "br":
            paragraph.add_run().add_break()
        else:
            if hasattr(child, "children"):
                _apply_inline_formatting(paragraph, child)

def _basic_html_to_docx(doc: Document, html_str: str):
    soup = BeautifulSoup(html_str, "html.parser")
    body = soup.body or soup

    for element in body.children:
        if getattr(element, "name", None) is None:
            txt = str(element).strip()
            if txt:
                doc.add_paragraph(txt)
            continue

        tag = element.name.lower()
        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(tag[1])
            text = element.get_text(strip=False)
            doc.add_heading(text, level=min(max(level, 1), 6))
            continue

        if tag == "p":
            p = doc.add_paragraph()
            _apply_inline_formatting(p, element)
            continue

        if tag in ("ul", "ol"):
            style = "List Bullet" if tag == "ul" else "List Number"
            for li in element.find_all("li", recursive=False):
                p = doc.add_paragraph(style=style)
                _apply_inline_formatting(p, li)
            continue

        if tag == "table":
            rows = element.find_all("tr", recursive=False)
            if not rows: continue
            first_cells = rows[0].find_all(["th","td"], recursive=False)
            cols = max(1, len(first_cells))
            first_is_header = any(c.name == "th" for c in first_cells)

            tbl = doc.add_table(rows=len(rows), cols=cols)
            try:
                tbl.style = "Table Grid"
            except Exception:
                pass
            _apply_grid_borders(tbl)

            sect = doc.sections[0]
            content_width = sect.page_width - sect.left_margin - sect.right_margin
            col_w = int(content_width / cols)
            _set_col_widths(tbl, [col_w]*cols)

            for r_idx, tr in enumerate(rows):
                cells = tr.find_all(["th","td"], recursive=False)
                for c_idx in range(cols):
                    cell = tbl.cell(r_idx, c_idx)
                    if not cell.paragraphs:
                        cell.add_paragraph()
                    p = cell.paragraphs[0]
                    if hasattr(p, "clear"):
                        p.clear()
                    if c_idx < len(cells):
                        _apply_inline_formatting(p, cells[c_idx])
                    else:
                        p.text = ""
                    _tight_paragraph(p)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

            if first_is_header:
                for c in tbl.rows[0].cells:
                    _shade_cell(c, "E7E6E6")
                    for r in c.paragraphs[0].runs:
                        r.bold = True
                    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue

        if tag in ("div","section","article"):
            p = doc.add_paragraph()
            _apply_inline_formatting(p, element)
            continue

        txt = element.get_text(strip=True)
        if txt:
            doc.add_paragraph(txt)

def html_to_docx_bytes(html_str: str, *, force_basic: bool = False) -> bytes:
    doc = Document()
    if not force_basic:
        try:
            HTML2Docx().add_html_to_document(html_str or "", doc)
        except Exception:
            _basic_html_to_docx(doc, html_str or "")
    else:
        _basic_html_to_docx(doc, html_str or "")

    if len(doc.paragraphs) == 0 and len(doc.tables) == 0:
        doc.add_paragraph("⚠️ HTML result is empty.")
    bio = BytesIO(); doc.save(bio)
    return bio.getvalue()

# ------------------------------------------------------------------------------
# MDL model & renderer (preview)
# ------------------------------------------------------------------------------
def summarize_finding_text(raw: str, max_chars: int = 1000) -> str:
    if not raw: return ""
    text = re.sub(r"\s+", " ", raw).strip()
    parts = re.split(r"(?<=[.?!])\s+", text)
    picked = []
    for p in parts:
        if len(picked) >= 3: break
        if re.search(r"\b(Assistance Listing|Award Period|Federal Program|Identification Number|CFDA)\b", p, re.I):
            continue
        picked.append(p)
    out = " ".join(picked) or text
    return _short(out, max_chars)

def format_letter_date(date_iso: Optional[str] = None) -> Tuple[str, str]:
    dt = datetime.fromisoformat(date_iso) if date_iso else datetime.utcnow()
    return dt.strftime("%Y-%m-%d"), dt.strftime("%B %d, %Y")

def render_mdl_html(model: Dict[str, Any]) -> str:
    letter_date_iso = model.get("letter_date_iso")
    _, letter_date_long = format_letter_date(letter_date_iso)

    auditee_name = model.get("auditee_name", "Recipient")
    ein = model.get("ein", "")
    address_lines = model.get("address_lines", [])
    attention_line = model.get("attention_line")
    period_end_text = model.get("period_end_text", str(model.get("audit_year", "")))
    include_no_qc_line = model.get("include_no_qc_line", True)
    treasury_contact_email = model.get("treasury_contact_email", "ORP_SingleAudits@treasury.gov")
    address_block = "<br>".join(html.escape(x) for x in address_lines) if address_lines else ""
    attention_block = f"<p><strong>{html.escape(attention_line)}</strong></p>" if attention_line else ""

    def _render_program_table(p: Dict[str, Any]) -> str:
        rows_html = []
        for f in p.get("findings", []):
            rows_html.append(f"""
              <tr>
                <td>{html.escape(f.get('finding_id',''))}</td>
                <td>{html.escape(f.get('compliance_and_summary') or f"{f.get('compliance_type','')} - {f.get('summary','')}".strip(" -"))}</td>                <td>{html.escape(f.get('summary',''))}</td>
                <td>{html.escape(f.get('audit_determination',''))}</td>
                <td>{html.escape(f.get('questioned_cost_determination',''))}</td>
                <td>{html.escape(f.get('cap_determination',''))}</td>
              </tr>
            """)
        if not rows_html:
            rows_html.append("""
              <tr><td colspan="5"><em>No MDL-relevant findings identified for this program.</em></td></tr>
            """)
        table = f"""
          <table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse; width:100%; font-size:10.5pt;">
            <tr>
              <th>Audit<br>Finding #</th>
              <th>Compliance Type -<br>Audit Finding</th>
              <th>Audit Finding<br>Determination</th>
              <th>Questioned Cost<br>Determination</th>
              <th>CAP<br>Determination</th>
            </tr>
            {''.join(rows_html)}
          </table>
        """
        cap_blocks = []
        for f in p.get("findings", []):
            cap_text = f.get("cap_text")
            if cap_text:
                cap_blocks.append(f"""
                  <h4>Corrective Action Plan – {html.escape(f.get('finding_id',''))}</h4>
                  <p>{html.escape(cap_text)}</p>
                """)
        return table + ("\n".join(cap_blocks) if cap_blocks else "")

    programs = model.get("programs", [])
    programs_html = "\n".join(_render_program_table(p) for p in programs) if programs else "<p><em>No MDL-relevant findings identified per FAC records.</em></p>"

    not_sustained_notes = model.get("not_sustained_notes", [])
    not_sustained_html = ""
    if not_sustained_notes:
        notes_paras = "\n".join(f"<p>{html.escape(n)}</p>" for n in not_sustained_notes if n)
        not_sustained_html = f"<h3>FINDINGS NOT SUSTAINED</h3>\n{notes_paras}"

    chunks = []
    chunks.append(f'<p style="text-align:right; margin:0 0 12pt 0;">{html.escape(letter_date_long)}</p>')
    chunks.append("""
      <p style="margin:0 0 12pt 0;">
        <strong>DEPARTMENT OF THE TREASURY</strong><br>
        WASHINGTON, D.C.
      </p>
    """)
    chunks.append(f"""
      <p style="margin:0 0 12pt 0;">
        <strong>{html.escape(auditee_name)}</strong><br>
        EIN: {html.escape(ein)}<br>
        {address_block}
      </p>
    """)
    if attention_block:
        chunks.append(attention_block)
    chunks.append(f"""
      <p style="margin:12pt 0 12pt 0;">
        <strong>Subject:</strong> U.S. Department of the Treasury’s Management Decision Letter (MDL) for Single Audit Report for the period ending on {html.escape(period_end_text)}
      </p>
    """)
    chunks.append("""
      <p>
        In accordance with 2 C.F.R. § 200.521(b), the U.S. Department of the Treasury (Treasury)
        is required to issue a management decision for single audit findings pertaining to awards under
        Treasury’s programs. Treasury’s review as part of its responsibilities under 2 C.F.R § 200.513(c)
        includes an assessment of Treasury’s award recipients’ single audit findings, corrective action plans (CAPs),
        and questioned costs, if any.
      </p>
    """)
    chunks.append(f"""
      <p>
        Treasury has reviewed the single audit report for {html.escape(auditee_name)}.
        Treasury has made the following determinations regarding the audit finding(s) and CAP(s) listed below.
      </p>
    """)
    if include_no_qc_line:
        chunks.append("<p>No questioned costs are included in this single audit report.</p>")

    chunks.append(programs_html)
    if not_sustained_html:
        chunks.append(not_sustained_html)

    # Email sentence removed per feedback.
    chunks.append("""
      <p>
        Please note, the corrective action(s) are subject to review during the recipient’s next annual single audit
        or program-specific audit, as applicable, to determine adequacy. If the same audit finding(s) appear in a future single
        audit report for this recipient, its current or future award funding under Treasury’s programs may be adversely impacted.
      </p>
      <p>
        For questions regarding the audit finding(s), please email us at
        <a href="mailto:{html.escape(treasury_contact_email)}">{html.escape(treasury_contact_email)}</a>.
        Thank you.
    </p>
      <p style="margin-top:18pt;">Sincerely,<br><br>
      Audit and Compliance Resolution Team<br>
      Office of Capital Access<br>
      U.S. Department of the Treasury</p>
    """)

    return f'<div style="font-family: Calibri, Arial, sans-serif; font-size:11pt; line-height:1.4;">{"".join(chunks)}</div>'

def build_mdl_model_from_fac(
    *,
    auditee_name: str,
    ein: str,
    audit_year: int,
    fac_general: List[Dict[str, Any]],
    fac_findings: List[Dict[str, Any]],
    fac_findings_text: List[Dict[str, Any]],
    fac_caps: List[Dict[str, Any]],
    federal_awards: List[Dict[str, Any]],
    period_end_text: Optional[str] = None,
    address_lines: Optional[List[str]] = None,
    attention_line: Optional[str] = None,
    only_flagged: bool = False,
    max_refs: int = 10,
    auto_cap_determination: bool = True,
    include_no_qc_line: bool = False,
    include_no_cap_line: bool = False,
    treasury_listings: Optional[List[str]] = None,
    aln_reference_xlsx: Optional[str] = None,
    aln_overrides_by_finding: Optional[Dict[str, str]] = None,
    **_
) -> Dict[str, Any]:
    """
    Builds the MDL model.

    Point C addressed:
      - Ensures ALN is not "Unknown" by:
        1) Loading Excel mapping (ALN -> canonical name; name -> ALN).
        2) Canonicalizing each grouped program after grouping:
           - If ALN known, use canonical label.
           - Else map by program name -> ALN.
           - Else apply Treasury heuristics (SLFRF/ERA/HAF/CPF/SSBCI/LATCF).
        3) THEN apply treasury_listings filter.
    """

    # --------- helpers ----------
    def _derive_assistance_listing(program_name: str, fallback: str = "Unknown") -> str:
        m = re.search(r"\b\d{2}\.\d{3}\b", program_name or "")
        return m.group(0) if m else fallback

    def _title_with_acronyms(s: str) -> str:
        """Title-case but preserve ALL-CAPS tokens (e.g., SLFRF) and common stop words."""
        if not s:
            return ""
        lowers = {"and","or","the","of","for","to","in","on","by","with","a","an"}
        out = []
        for tok in str(s).split():
            if tok.isupper() and len(tok) > 1:
                out.append(tok)  # keep acronym
            else:
                w = tok.lower()
                out.append(w if w in lowers else w.capitalize())
        return " ".join(out)

    def _load_aln_mapping(xlsx_path: Optional[str]):
        """
        Returns:
          aln_to_label: {'21.027': 'Coronavirus State and Local Fiscal Recovery Funds (SLFRF)', ...}
          name_to_aln:  {'coronavirus state and local fiscal recovery funds': '21.027', ...}
        Tolerant to header naming; safe no-op if workbook missing or openpyxl not installed.
        """
        aln_to_label, name_to_aln = {}, {}
        if not xlsx_path:
            return aln_to_label, name_to_aln
        try:
            import openpyxl
        except Exception:
            return aln_to_label, name_to_aln
        try:
            wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        except Exception:
            return aln_to_label, name_to_aln

        # Heuristic: find the first sheet that looks like a mapping
        for ws in wb.worksheets:
            hdrs = [(c.value.strip() if isinstance(c.value, str) else (c.value or "")) for c in ws[1]]
            if not any(hdrs):
                continue
            hl = [str(h).strip().lower() for h in hdrs]

            def _find_col(candidates: List[str]) -> Optional[int]:
                for i, h in enumerate(hl):
                    for cand in candidates:
                        cand = cand.lower()
                        if h == cand or cand in h:
                            return i
                return None

            i_aln = _find_col(["aln", "assistance listing", "assistance listing number", "cfda", "cfda number"])
            i_prog = _find_col(["program", "program name", "assistance listing title", "program title"])
            i_acr = _find_col(["acronym", "short", "short name", "abbrev", "abbreviation"])

            if i_prog is None or (i_aln is None and i_acr is None):
                continue  # not a mapping sheet

            for row in ws.iter_rows(min_row=2, values_only=True):
                raw_aln = (row[i_aln] if i_aln is not None else "") or ""
                raw_prog = (row[i_prog] or "")
                raw_acr = (row[i_acr] if i_acr is not None else "") or ""

                aln = str(raw_aln).strip()
                prog_name = str(raw_prog).strip()
                acr = str(raw_acr).strip()
                if not prog_name:
                    continue

                canonical_name = _title_with_acronyms(prog_name)
                if acr:
                    canonical_name = f"{canonical_name} ({acr})"

                if aln:
                    aln_to_label[aln] = canonical_name
                name_to_aln[prog_name.lower()] = aln  # may be "" if not provided
            break  # first matching sheet is enough

        return aln_to_label, name_to_aln

    def _apply_canonicalization_after_grouping(
        group: Dict[str, Any],
        aln_to_label: Dict[str, str],
        name_to_aln: Dict[str, str],
    ):
        """
        Normalize 'assistance_listing' and 'program_name' in-place for a group:
          - Use Excel ALN→label if ALN known.
          - Else map by name→ALN.
          - Else apply Treasury heuristics.
          - Always tidy program name casing.
        """
        cur_aln  = (group.get("assistance_listing") or "").strip()
        cur_name = (group.get("program_name") or "").strip()

        # 1) If we already have a valid ALN, use canonical label
        if cur_aln and cur_aln in aln_to_label:
            group["assistance_listing"] = cur_aln
            group["program_name"] = aln_to_label[cur_aln]
            return

        # 2) If ALN missing/Unknown, try via name → ALN
        guess_aln = name_to_aln.get(cur_name.lower())
        if (not cur_aln or cur_aln == "Unknown") and guess_aln:
            group["assistance_listing"] = guess_aln
            group["program_name"] = aln_to_label.get(guess_aln, _title_with_acronyms(cur_name or "Unknown Program"))
            return

        # 3) Treasury heuristics (common programs) — last resort
        nm = cur_name.lower()
        heuristics = [
            ("slfrf", ("21.027", "Coronavirus State and Local Fiscal Recovery Funds (SLFRF)")),
            ("fiscal recovery", ("21.027", "Coronavirus State and Local Fiscal Recovery Funds (SLFRF)")),
            ("emergency rental assistance", ("21.023", "Emergency Rental Assistance Program (ERA)")),
            ("homeowner assistance", ("21.026", "Homeowner Assistance Fund (HAF)")),
            ("capital projects fund", ("21.029", "Capital Projects Fund (CPF)")),
            ("state small business credit", ("21.031", "State Small Business Credit Initiative (SSBCI)")),
            ("local assistance and tribal consistency", ("21.032", "Local Assistance and Tribal Consistency Fund (LATCF)")),
        ]
        for key, (aln_guess, label) in heuristics:
            if key in nm:
                group["assistance_listing"] = aln_guess
                group["program_name"] = label
                break

        # 4) Final tidy for program name casing if still raw/all-caps
        final_name = (group.get("program_name") or "").strip()
        if final_name.isupper() or not final_name:
            group["program_name"] = _title_with_acronyms(final_name or cur_name or "Unknown Program")
        if not group.get("assistance_listing"):
            group["assistance_listing"] = "Unknown"

    # --------- load mapping once ----------
    aln_to_label, name_to_aln = _load_aln_mapping(aln_reference_xlsx)
    # ADD THIS:
    type_map, summary_labels = _load_finding_mappings(aln_reference_xlsx)
    # --------- award lookups from FAC ----------
    award2meta: Dict[str, Dict[str, str]] = {}
    for a in (federal_awards or []):
        ref = a.get("award_reference")
        pname = (a.get("federal_program_name") or "").strip()
        explicit_aln = (a.get("assistance_listing") or a.get("assistance_listing_number") or "").strip()
        derived_aln = _derive_assistance_listing(pname, fallback="")
        aln = explicit_aln or derived_aln or "Unknown"

        # Prefer Excel canonical if we have the ALN
        if aln != "Unknown" and aln in aln_to_label:
            canonical_name = aln_to_label[aln]
        else:
            # Try map by name -> ALN
            mapped_aln = name_to_aln.get(pname.lower())
            if mapped_aln:
                aln = mapped_aln
                canonical_name = aln_to_label.get(mapped_aln, _title_with_acronyms(pname or "Unknown Program"))
            else:
                canonical_name = _title_with_acronyms(pname or "Unknown Program")

        if ref:
            award2meta[ref] = {
                "program_name": canonical_name or "Unknown Program",
                "assistance_listing": aln or "Unknown",
            }

    # --------- text / CAP lookups ----------
    text_by_ref = {
        _norm_ref(t.get("finding_ref_number")): (t.get("finding_text") or "").strip()
        for t in (fac_findings_text or [])
    }
    cap_by_ref = {
        _norm_ref(c.get("finding_ref_number")): (c.get("planned_action") or "").strip()
        for c in (fac_caps or [])
    }

    def _is_flagged(f: dict) -> bool:
        return any([
            f.get("is_material_weakness") is True,
            f.get("is_significant_deficiency") is True,
            f.get("is_questioned_costs") is True,
            f.get("is_modified_opinion") is True,
            f.get("is_other_findings") is True,
            f.get("is_other_matters") is True,
            f.get("is_repeat_finding") is True,
        ])

    # --------- choose finding refs ----------
    base_refs: List[str] = []
    for f in (fac_findings or []):
        if only_flagged and not _is_flagged(f):
            continue
        r = f.get("reference_number")
        if r:
            base_refs.append(r)

    if not base_refs and fac_findings_text:
        base_refs = [t.get("finding_ref_number") for t in fac_findings_text if t.get("finding_ref_number")]

    seen = set()
    norm_refs: List[Tuple[str, str]] = []
    for r in base_refs:
        if not r:
            continue
        k = _norm_ref(r)
        if k not in seen:
            seen.add(k)
            norm_refs.append((r, k))
    norm_refs = norm_refs[: max_refs or 10]
    chosen_keys = {kn for _, kn in norm_refs}

    # --------- group findings under award_reference ----------
    programs_map: Dict[str, Dict[str, Any]] = {}
    for f in (fac_findings or []):
        r = f.get("reference_number")
        if not r:
            continue
        k = _norm_ref(r)
        if k not in chosen_keys:
            continue

        award_ref = f.get("award_reference") or "UNKNOWN"
        meta = award2meta.get(award_ref, {})
        group = programs_map.setdefault(award_ref, {
            "assistance_listing": meta.get("assistance_listing", "Unknown"),
            "program_name": meta.get("program_name", "Unknown Program"),
            "findings": []
        })
        
        # If ALN is Unknown, try to fill from finding-level override (XLSX)
        if group.get("assistance_listing") in (None, "", "Unknown"):
            # use the original finding reference if available
            orig_ref = f.get("reference_number") or ""
            # try both raw and normalized keys
            cand_aln = None
            if aln_overrides_by_finding:
                cand_aln = (aln_overrides_by_finding.get(orig_ref)
                            or aln_overrides_by_finding.get(_norm_ref(orig_ref)))
            if cand_aln:
                group["assistance_listing"] = cand_aln
                # if we have an ALN→label map, upgrade program_name too
                if cand_aln in aln_to_label:
                    group["program_name"] = aln_to_label[cand_aln]
        summary  = summarize_finding_text(text_by_ref.get(k, ""))
        cap_text = cap_by_ref.get(k)

        qcost_det = "No questioned costs identified" if include_no_qc_line else "None"
        cap_det   = (
            "Accepted" if (auto_cap_determination and cap_text)
            else ("No CAP required" if include_no_cap_line else "Not Applicable")
        )

        ctype_code = (f.get("type_requirement") or "").strip().upper()[:1]
        ctype_label = type_map.get(ctype_code) or ctype_code or ""
        matched_label = (_best_summary_label_openai(summary, summary_labels)
                 or _best_summary_label(summary, summary_labels)
                 or summary)
        print(f"Matched label: {matched_label}")
        print("\n")
        print(f"Compliance type: {ctype_label}")
        print(f"Summary: {summary}")
        print(" for finding {f.get('reference_number')}")
        print(ctype_label, summary, cap_text, qcost_det, cap_det)
        group["findings"].append({
            "finding_id": f.get("reference_number") or "",
            "compliance_type": ctype_label,  # use the full label, not just 'I'
            "summary": summary,
            "compliance_and_summary": f"{ctype_label} - {matched_label}".strip(" -"),
            "audit_determination": "Sustained",
            "questioned_cost_determination": qcost_det,
            "disallowed_cost_determination": "None",
            "cap_determination": cap_det,
            "cap_text": cap_text,
        })
    # If nothing grouped but we have refs, emit a catch-all
    if not programs_map and norm_refs:
        catchall = {"assistance_listing": "Unknown", "program_name": "Unknown Program", "findings": []}
        ctype_code = (f.get("type_requirement") or "").strip().upper()[:1]
        ctype_label = type_map.get(ctype_code) or ctype_code or ""
        matched_label = _best_summary_label(summary, summary_labels) or summary
        for orig, key in norm_refs:
            cap_text = cap_by_ref.get(key)
            qcost_det = "No questioned costs identified" if include_no_qc_line else "None"
            cap_det   = (
                "Accepted" if (auto_cap_determination and cap_text)
                else ("No CAP required" if include_no_cap_line else "Not Applicable")
            )
            catchall["findings"].append({
                "finding_id": orig,
                "compliance_type": ctype_label, # use the full label not just 'I'
                "summary": summarize_finding_text(text_by_ref.get(key, "")),
                "compliance_and_summary": f"{ctype_label} - {matched_label}".strip(" -"),
                "audit_determination": "Sustained",
                "questioned_cost_determination": qcost_det,
                "disallowed_cost_determination": "None",
                "cap_determination": cap_det,
                "cap_text": cap_text,
            })
        programs_map["UNKNOWN"] = catchall

    # ----- Canonicalize program fields using the Excel mapping (now that groups exist)
    for grp in programs_map.values():
        _apply_canonicalization_after_grouping(grp, aln_to_label, name_to_aln)

    # ----- Apply Treasury ALN filter AFTER canonicalization
    if treasury_listings:
        allowed = {(aln or "").strip() for aln in treasury_listings if aln}
        programs_map = {k: v for k, v in programs_map.items() if v.get("assistance_listing") in allowed}

    # ----- Build final model
    model = {
        "letter_date_iso": datetime.utcnow().strftime("%Y-%m-%d"),
        "auditee_name": auditee_name,
        "ein": f"{ein[:2]}-{ein[2:]}" if ein and ein.isdigit() and len(ein) == 9 else ein,
        "address_lines": address_lines or [],
        "attention_line": attention_line or "",
        "period_end_text": period_end_text or f"June 30, {audit_year}",
        "audit_year": audit_year,
        "programs": list(programs_map.values()),
        "not_sustained_notes": [],
    }
    return model

# ------------------------------------------------------------------------------
# Template helpers
# ------------------------------------------------------------------------------
def _para_text(p: Paragraph) -> str:
    return "".join(run.text for run in p.runs)

def _clear_runs(p: Paragraph):
    for r in list(p.runs):
        r._element.getparent().remove(r._element)

def _replace_in_paragraph_run_aware(p: Paragraph, mapping: Dict[str, str]) -> bool:
    original = _para_text(p)
    if not original:
        return False
    new_text = original
    for k, v in mapping.items():
        if k in new_text:
            new_text = new_text.replace(k, v)
    if new_text != original:
        _clear_runs(p)
        p.add_run(new_text)
        return True
    return False

def _iter_cells_in_table(tbl: Table):
    for row in tbl.rows:
        for cell in row.cells:
            yield cell

def _iter_all_paragraphs_in_container(container) -> list[Paragraph]:
    items = []
    if hasattr(container, "paragraphs"):
        items.extend(container.paragraphs)
    if hasattr(container, "tables"):
        for t in container.tables:
            for c in _iter_cells_in_table(t):
                items.extend(c.paragraphs)
                for nt in c.tables:
                    for nc in _iter_cells_in_table(nt):
                        items.extend(nc.paragraphs)
    return items

def _replace_placeholders_docwide(doc: Document, mapping: Dict[str, str]):
    for p in _iter_all_paragraphs_in_container(doc):
        _replace_in_paragraph_run_aware(p, mapping)
    for sec in doc.sections:
        for p in _iter_all_paragraphs_in_container(sec.header):
            _replace_in_paragraph_run_aware(p, mapping)
        for p in _iter_all_paragraphs_in_container(sec.footer):
            _replace_in_paragraph_run_aware(p, mapping)

def _find_anchor_paragraph(doc: Document, anchor: str) -> Optional[Paragraph]:
    for p in _iter_all_paragraphs_in_container(doc):
        if anchor in _para_text(p):
            return p
    for sec in doc.sections:
        for p in _iter_all_paragraphs_in_container(sec.header):
            if anchor in _para_text(p):
                return p
        for p in _iter_all_paragraphs_in_container(sec.footer):
            if anchor in _para_text(p):
                return p
    return None

def _delete_immediate_next_table(anchor_para: Paragraph):
    """If the template has a placeholder table immediately after the anchor paragraph, delete it."""
    p_el = anchor_para._p
    nxt = p_el.getnext()
    if nxt is not None and nxt.tag.endswith("tbl"):
        parent = nxt.getparent()
        parent.remove(nxt)

def _pick_table_style(doc: Document) -> Optional[str]:
    if getattr(doc, "tables", None):
        for t in doc.tables:
            try:
                if t.style and t.style.name:
                    return t.style.name
            except Exception:
                pass
    try:
        _ = doc.styles["Table Grid"]
        return "Table Grid"
    except KeyError:
        pass
    for st in doc.styles:
        try:
            if st.type == WD_STYLE_TYPE.TABLE:
                return st.name
        except Exception:
            continue
    return None

def _build_program_table(doc: Document, program: Dict[str, Any]) -> Table:
    findings = program.get("findings", []) or []
    rows = max(1, len(findings)) + 1

    tbl = doc.add_table(rows=rows, cols=6)
    _style = _pick_table_style(doc)
    if _style:
        try:
            tbl.style = _style
        except Exception:
            pass
    _apply_grid_borders(tbl)  # ensure borders even without style

    headers = [
        "Audit\nFinding #",
        "Compliance Type -\nAudit Finding",
        "Audit Finding\nDetermination",
        "Questioned Cost\nDetermination",
        "CAP\nDetermination",
    ]
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        _clear_runs(cell.paragraphs[0])
        cell.paragraphs[0].add_run(h)
        _shade_cell(cell, "E7E6E6")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        _tight_paragraph(cell.paragraphs[0])

    if findings:
        for r, f in enumerate(findings, start=1):
            vals = [
                f.get("finding_id", ""),
                f.get("compliance_type", ""),
                f.get("summary", ""),
                f.get("audit_determination", "Sustained"),
                f.get("questioned_cost_determination", "None"),
                f.get("cap_determination", "Not Applicable"),
            ]
            for c, val in enumerate(vals):
                cell = tbl.cell(r, c)
                _clear_runs(cell.paragraphs[0])
                cell.paragraphs[0].add_run(str(val))
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                _tight_paragraph(cell.paragraphs[0])
    else:
        cell = tbl.cell(1, 0)
        _clear_runs(cell.paragraphs[0])
        cell.paragraphs[0].add_run("—")

    return tbl

def _insert_program_tables_at_anchor(doc: Document, anchor_para: Paragraph, programs: List[Dict[str, Any]]):
    # Clean anchor text and delete any placeholder table immediately following it
    text = _para_text(anchor_para).replace("[[PROGRAM_TABLES]]", "")
    _clear_runs(anchor_para)
    if text.strip():
        anchor_para.add_run(text)

    _delete_immediate_next_table(anchor_para)

    # Order programs by ALN
    def _al_key(p):
        return (p.get("assistance_listing") or "99.999")
    programs_sorted = sorted(programs or [], key=_al_key)

    last = anchor_para
    for p in programs_sorted:
        al = p.get("assistance_listing", "Unknown")
        name = p.get("program_name", "Unknown Program")
        heading = f"Assistance Listing Number/Program Name: {al} / {name}"
        heading_para = doc.add_paragraph()
        _clear_runs(heading_para); heading_para.add_run(heading)

        # splice heading after 'last'
        heading_el = heading_para._p
        heading_el.getparent().remove(heading_el)
        _insert_after(last, heading_el)
        last = heading_el

        # table
        tbl = _build_program_table(doc, p)
        tbl_el = tbl._tbl
        tbl_el.getparent().remove(tbl_el)
        _insert_after(last, tbl_el)
        last = tbl_el

        # CAPs
        for f in p.get("findings", []):
            cap_text = (f or {}).get("cap_text")
            if cap_text:
                cap_title = doc.add_paragraph()
                _clear_runs(cap_title); cap_title.add_run(f"Corrective Action Plan – {f.get('finding_id','')}")
                cap_text_para = doc.add_paragraph()
                _clear_runs(cap_text_para); cap_text_para.add_run(cap_text)

                cap_title_el = cap_title._p; cap_text_el = cap_text_para._p
                cap_title_el.getparent().remove(cap_title_el)
                cap_text_el.getparent().remove(cap_text_el)
                _insert_after(last, cap_title_el)
                _insert_after(cap_title_el, cap_text_el)
                last = cap_text_el

        # spacer
        spacer = doc.add_paragraph()
        spacer_el = spacer._p
        spacer_el.getparent().remove(spacer_el)
        _insert_after(last, spacer_el)
        last = spacer_el

def _remove_watermarks(doc):
    """
    Removes WordArt/VML watermark shapes (e.g., 'PowerPlusWaterMarkObject', 'DRAFT')
    from headers/footers/body. Works with typical Word 'DRAFT' watermarks.
    """
    for part in doc.part.package.parts:
        if not hasattr(part, "element"):
            continue
        root = part.element
        ns = root.nsmap
        # Remove shapes named like watermark or with textpath 'DRAFT'
        for pict in list(root.xpath(".//*[local-name()='pict']")):
            kill_pict = False
            for shp in pict.xpath(".//*[local-name()='shape']"):
                sid = (shp.get("id") or "") + " " + (shp.get("{urn:schemas-microsoft-com:office:office}spid") or "")
                if "PowerPlusWaterMarkObject" in sid or "WaterMark" in sid:
                    kill_pict = True
                    break
                for tx in shp.xpath(".//*[local-name()='textpath'][@string]"):
                    if "DRAFT" in (tx.get("string") or "").upper():
                        kill_pict = True
                        break
            if kill_pict:
                parent = pict.getparent()
                parent.remove(pict)

def _insert_paragraph_after(p: Paragraph, text: str = "") -> Paragraph:
    new_p = p._element.addnext(p._element.__class__())
    # Create a real python-docx Paragraph around that element
    from docx.text.paragraph import Paragraph as _Para
    new_par = _Para(new_p, p._parent)
    if text:
        new_par.add_run(text)
    return new_par

def _find_para_by_contains(doc: Document, needle: str) -> Optional[Paragraph]:
    def _norm(s: str) -> str:
        s = (s or "").replace("\u00A0"," ").replace("\xa0"," ")
        s = s.replace("\u200b","").replace("\u200c","").replace("\u200d","")
        return " ".join(s.split())
    N = _norm(needle)
    for p in _iter_all_paragraphs_in_container(doc):
        if N in _norm(_para_text(p)):
            return p
    for sec in doc.sections:
        for p in _iter_all_paragraphs_in_container(sec.header):
            if N in _norm(_para_text(p)):
                return p
        for p in _iter_all_paragraphs_in_container(sec.footer):
            if N in _norm(_para_text(p)):
                return p
    return None

def build_docx_from_template(model: Dict[str, Any], *, template_path: str) -> bytes:
    """
    Open a .docx template and:
      1) Replace placeholders across the whole document (headers/footers too)
      2) Insert program tables at the [[PROGRAM_TABLES]] anchor
    """
    if not os.path.isfile(template_path):
        raise HTTPException(400, f"Template not found: {template_path}")

    doc = Document(template_path)
    _remove_watermarks(doc)  # remove DRAFT/Watermark shapes immediately

    # Dates
    _, letter_date_long = format_letter_date(model.get("letter_date_iso"))

    # Header fields (defaults -> empty so placeholders never leak through)
    auditee = (model.get("auditee_name")
               or model.get("recipient_name")
               or "")
    if not auditee.lower().startswith("the "):
        auditee = "The " + auditee
    ein     = model.get("ein", "") or ""
    street  = model.get("street_address", "") or ""
    city    = model.get("city", "") or ""
    state   = model.get("state", "") or ""
    zipc    = model.get("zip_code", "") or ""
    poc     = model.get("poc_name", "") or ""
    poc_t   = model.get("poc_title", "") or ""
    auditor = model.get("auditor_name", "") or ""
    if auditor and not auditor.lower().startswith("the "):
        auditor = "the " + auditor
    print(f"Auditor: {auditor}")
    print(f"Auditee: {auditee}")
    fy_end  = (model.get("period_end_text")
               or str(model.get("audit_year", ""))) or ""
    # Treasury contact email
    treasury_contact_email = "ORP_SingleAudits@treasury.gov"

    # Map BOTH styles of placeholders used by the template
    mapping = {
        # date stub used in some templates
        "Date XX, 2025": letter_date_long,

        # [bracket] style
        "[Recipient Name]": auditee,
        "[EIN]": ein,
        "[Street Address]": street,
        "[City]": city,
        "[State]": state,
        "[Zip Code]": zipc,
        "[Point of Contact]": poc,
        "[Point of Contact Title]": poc_t,
        "[Auditor Name]": auditor,
        "[Fiscal Year End Date]": fy_end,
        "[The]": "The", 
        "[the]": "the",

        # ${curly} style
        "${recipient_name}": auditee,
        "${ein}": ein,
        "${street_address}": street,
        "${city}": city,
        "${state}": state,
        "${zip_code}": zipc,
        "${poc_name}": poc,
        "${poc_title}": poc_t,
        "${auditor_name}": auditor,
        "${fy_end_text}": fy_end,
    }
    email = (model.get("treasury_contact_email") or "ORP_SingleAudits@treasury.gov").strip()

    mapping.update({
        # bracket style used by template
        "[treasury_contact_email]": email,
        # curly style just in case
        "${treasury_contact_email}": email,
    })
    # Ensure no None values sneak in
    mapping = {k: (v if v is not None else "") for k, v in mapping.items()}

    # 1) Replace placeholders everywhere (body + headers/footers + nested tables)
    _replace_placeholders_docwide(doc, mapping)
    _fix_treasury_email(doc, model.get("treasury_contact_email") or "ORP_SingleAudits@treasury.gov")
    _strip_leading_token_artifacts(doc)
    _unset_all_caps_everywhere(doc)

    # 2) Insert program tables at the anchor (do this BEFORE stripping bracketed tokens,
    # because cleanup would otherwise delete the [[PROGRAM_TABLES]] marker)
    anchor = _find_anchor_paragraph(doc, "[[PROGRAM_TABLES]]")
    if not anchor:
        raise HTTPException(400, "Template does not contain the [[PROGRAM_TABLES]] anchor paragraph.")
    programs = model.get("programs", []) or []
    # Find the visible label paragraph and fill it with ALN/Program from the first program
    try:
        label_p = _find_para_by_contains(doc, "Assistance Listing Number/Program Name")
        progs = model.get("programs") or []
        if label_p is not None and progs:
            first = progs[0]
            aln = (first.get("assistance_listing") or "").strip()
            pname = (first.get("program_name") or "").strip()
            # Title-case the program if it somehow stayed all-caps
            def _fix_case(s: str) -> str:
                if s.isupper():
                    lowers = {"and","or","the","of","for","to","in","on","by","with","a","an"}
                    parts = []
                    for w in s.split():
                        lw = w.lower()
                        parts.append(lw if lw in lowers else lw.capitalize())
                    return " ".join(parts)
                return s
            pname = _fix_case(pname)
            _clear_runs(label_p)
            label_p.add_run(f"Assistance Listing Number/Program Name: {aln} / {pname}")
    except Exception:
        pass
    _insert_program_tables_at_anchor(doc, anchor, programs)
    # Remove duplicate narrative blocks under the table
    # Remove duplicate narrative that appears below the table
    try:
        _cleanup_post_table_narrative(doc, model)
    except Exception:
        pass
    # Compute total findings and pluralize boilerplate
    # try:
    #     total_findings = sum(len(p.get("findings") or []) for p in (model.get("programs") or []))
    #     _pluralize_text_everywhere(doc, total_findings)
    # except Exception:
    #     pass
    # Grammar-fix optional plurals via OpenAI (if key set)
    try:
        total_findings = sum(len(prog.get("findings") or []) for prog in (model.get("programs") or []))
        _ai_fix_pluralization_in_doc(doc, total_findings)
    except Exception:
        pass

    # Final tidy: strip any *remaining* token patterns like ${...} or [...]
    def _strip_leftovers_in_container(container):
        for p in _iter_all_paragraphs_in_container(container):
            t = _para_text(p)
            if not t:
                continue
            new_t = t
            if "${" in new_t:
                new_t = re.sub(r"\$\{[^}]+\}", "", new_t)
            if "[" in new_t and "]" in new_t:
                new_t = re.sub(r"\[[^\]]+\]", "", new_t)
            if new_t != t:
                _clear_runs(p)
                p.add_run(new_t)

    _strip_leftovers_in_container(doc)
    for sec in doc.sections:
        _strip_leftovers_in_container(sec.header)
        _strip_leftovers_in_container(sec.footer)
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()


# --- placeholder cleanup helpers ---
PLACEHOLDER_RE = re.compile(r"^\s*\$\{[^}]+\}\s*$")

def _none_if_placeholder(v):
    """Return None if value looks like an unresolved ${var} placeholder."""
    return None if isinstance(v, str) and PLACEHOLDER_RE.match(v.strip()) else v

def _str_or_default(v, default=""):
    """If v is placeholder/blank/None return default, else v."""
    v = _none_if_placeholder(v)
    if isinstance(v, str) and v.strip():
        return v
    return default

import requests
from io import BytesIO

def _read_headers(ws):
    return [ (c.value or "").strip() if isinstance(c.value, str) else (c.value or "") for c in ws[1] ]

def _find_col(headers, candidates):
    hl = [str(h).strip().lower() for h in headers]
    for i, h in enumerate(hl):
        for cand in candidates:
            cl = cand.strip().lower()
            if h == cl or cl in h:
                return i
    return None

def _aln_overrides_from_summary(report_id: str):
    """
    Returns (aln_by_award, aln_by_finding) by parsing the public FAC summary XLSX.
    """
    url = f"https://app.fac.gov/dissemination/summary-report/xlsx/{report_id}"
    r = requests.get(url, timeout=20)
    r.raise_for_status()

    import openpyxl
    wb = openpyxl.load_workbook(BytesIO(r.content), data_only=True)

    aln_by_award, aln_by_finding = {}, {}
    # Look for a sheet with findings
    for ws in wb.worksheets:
        headers = _read_headers(ws)
        if not any(headers):
            continue
        i_findref = _find_col(headers, ["finding_ref_number", "finding reference number", "reference_number"])
        i_award   = _find_col(headers, ["award_reference", "award reference"])
        i_aln     = _find_col(headers, ["assistance listing", "assistance listing number", "aln", "cfda", "cfda number"])
        if i_aln is None:
            continue

        for row in ws.iter_rows(min_row=2, values_only=True):
            findref = (row[i_findref] if i_findref is not None else "") or ""
            award   = (row[i_award]   if i_award   is not None else "") or ""
            aln     = (row[i_aln]     if i_aln     is not None else "") or ""
            aln = str(aln).strip()
            if not aln:
                continue
            if award:
                aln_by_award[str(award).strip()] = aln
            if findref:
                aln_by_finding[str(findref).strip()] = aln
        break  # first matching sheet is enough

    return aln_by_award, aln_by_finding

# ------------------------------------------------------------------------------
# Schemas
# ------------------------------------------------------------------------------
class BuildDocx(BaseModel):
    auditee_name: str
    ein: str
    audit_year: int
    body_html: Optional[str] = None
    body_html_b64: Optional[str] = None
    dest_path: Optional[str] = None
    filename: Optional[str] = None

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

class BuildByReport(BaseModel):
    auditee_name: str
    ein: str
    audit_year: int
    report_id: str
    dest_path: Optional[str] = None
    only_flagged: bool = False
    max_refs: int = 15
    include_awards: bool = True

class BuildByReportTemplated(BuildByReport):
    auditor_name: Optional[str] = None
    fy_end_text: Optional[str] = None
    recipient_name: Optional[str] = None
    street_address: Optional[str] = None
    city: Optional[str] = None
    state: Optional[str] = None
    zip_code: Optional[str] = None
    poc_name: Optional[str] = None
    poc_title: Optional[str] = None
    template_path: Optional[str] = None
    treasury_listings: Optional[List[str]] = None

# ------------------------------------------------------------------------------
# Routes
# ------------------------------------------------------------------------------
@app.get("/healthz")
def healthz():
    logging.info(f"Incoming payload to healthsz endpoint")
    return {"ok": True, "time": datetime.utcnow().isoformat()}

@app.post("/echo")
def echo(payload: Dict[str, Any]):
    logging.info(f"Incoming payload to echo endpoint: {payload}")
    #logging.info(f"Incoming payload: {payload.dict()}")
    return {"received": payload, "ts": datetime.utcnow().isoformat()}

@app.get("/debug/env")
def debug_env():
    key = os.getenv("FAC_API_KEY") or ""
    masked = (key[:4] + "…" + key[-2:]) if key else None
    return {"fac_api_key_present": bool(key), "fac_api_key_masked": masked}

@app.get("/debug/storage")
def debug_storage():
    info = _parse_conn_str(AZURE_CONN_STR) if AZURE_CONN_STR else {}
    return {"using_storage": bool(AZURE_CONN_STR), "account": info.get("AccountName"), "blob_endpoint": info.get("BlobEndpoint")}

@app.get("/debug/sas")
def debug_sas():
    if not AZURE_CONN_STR:
        raise HTTPException(400, "Set AZURE_STORAGE_CONNECTION_STRING to test SAS.")
    url = upload_and_sas(AZURE_CONTAINER, "debug/hello.txt", b"hi", ttl_minutes=5)
    return {"url": url}

@app.post("/build-docx-demo")
def build_docx_demo():
    document = Document()
    document.add_heading("Hello from the DOCX demo ✅", level=1)
    document.add_paragraph("If you can read this, your write/upload pipeline is good.")
    bio = BytesIO(); document.save(bio); data = bio.getvalue()
    blob_name = "demo/hello.docx"
    url = upload_and_sas(AZURE_CONTAINER, blob_name, data) if AZURE_CONN_STR else save_local_and_url(blob_name, data)
    return {"url": url, "size_bytes": len(data)}

@app.post("/build-docx")
def build_docx(req: BuildDocx):
    html_str = (req.body_html or "").strip()
    if (not html_str) and req.body_html_b64:
        try:
            html_str = base64.b64decode(req.body_html_b64).decode("utf-8", errors="ignore")
        except Exception:
            raise HTTPException(400, "Invalid base64 in body_html_b64")
    if not html_str:
        raise HTTPException(400, "body_html (or body_html_b64) is required")

    data = html_to_docx_bytes(html_str, force_basic=False)
    folder = (req.dest_path or "").lstrip("/")
    base = req.filename or f"MDL-{sanitize(req.auditee_name)}-{sanitize(req.ein)}-{req.audit_year}.docx"
    blob_name = f"{folder}{base}" if folder else base
    url = upload_and_sas(AZURE_CONTAINER, blob_name, data) if AZURE_CONN_STR else save_local_and_url(blob_name, data)
    return {"url": url, "blob_path": f"{AZURE_CONTAINER}/{blob_name}", "size_bytes": len(data)}

@app.post("/build-docx-from-fac")
def build_docx_from_fac(req: BuildFromFAC):
    model = build_mdl_model_from_fac(
        auditee_name=req.auditee_name,
        ein=req.ein,
        audit_year=req.audit_year,
        fac_general=req.fac_general,
        fac_findings=req.fac_findings,
        fac_findings_text=req.fac_findings_text,
        fac_caps=req.fac_caps,
        federal_awards=req.federal_awards,
        only_flagged=False,
        max_refs=25,
        include_no_qc_line=True,
    )
    html_str = render_mdl_html(model)
    data = html_to_docx_bytes(html_str, force_basic=True)

    folder = (req.dest_path or "").lstrip("/")
    base = req.filename or f"MDL-Preview-{sanitize(req.auditee_name)}-{sanitize(req.ein)}-{req.audit_year}.docx"
    blob_name = f"{folder}{base}" if folder else base
    url = upload_and_sas(AZURE_CONTAINER, blob_name, data) if AZURE_CONN_STR else save_local_and_url(blob_name, data)
    return {"url": url, "blob_path": f"{AZURE_CONTAINER}/{blob_name}", "size_bytes": len(data)}



@app.post("/build-docx-by-report")
def build_docx_by_report(req: BuildByReport):
    fac_general = _fac_get("general", {
        "report_id": f"eq.{req.report_id}",
        "select": "report_id,fac_accepted_date",
        "limit": 1
    })

    findings_params = {
        "report_id": f"eq.{req.report_id}",
        "select": "reference_number,award_reference,type_requirement,"
                  "is_material_weakness,is_significant_deficiency,is_questioned_costs,"
                  "is_modified_opinion,is_other_findings,is_other_matters,is_repeat_finding",
        "order": "reference_number.asc",
        "limit": str(req.max_refs)
    }
    if req.only_flagged:
        flagged = [
            "is_material_weakness","is_significant_deficiency","is_questioned_costs",
            "is_modified_opinion","is_other_findings","is_other_matters","is_repeat_finding"
        ]
        findings_params["or"] = "(" + ",".join([f"{f}.eq.true" for f in flagged]) + ")"
    fac_findings = _fac_get("findings", findings_params)

    refs = [row.get("reference_number") for row in fac_findings if row.get("reference_number")]
    refs = refs[: req.max_refs]

    if refs:
        fac_findings_text = _fac_get("findings_text", {
            "report_id": f"eq.{req.report_id}",
            "select": "finding_ref_number,finding_text",
            "order": "finding_ref_number.asc",
            "limit": str(len(refs)),
            "or": _or_param("finding_ref_number", refs)
        })
        fac_caps = _fac_get("corrective_action_plans", {
            "report_id": f"eq.{req.report_id}",
            "select": "finding_ref_number,planned_action",
            "order": "finding_ref_number.asc",
            "limit": str(len(refs)),
            "or": _or_param("finding_ref_number", refs)
        })
    else:
        fac_findings_text, fac_caps = [], []

    federal_awards = []
    if req.include_awards:
        federal_awards = _fac_get("federal_awards", {
            "report_id": f"eq.{req.report_id}",
            "select": "award_reference,federal_program_name",
            "order": "award_reference.asc",
            "limit": "200"
        })

    model = build_mdl_model_from_fac(
        auditee_name=req.auditee_name,
        ein=req.ein,
        audit_year=req.audit_year,
        fac_general=fac_general,
        fac_findings=fac_findings,
        fac_findings_text=fac_findings_text,
        fac_caps=fac_caps,
        federal_awards=federal_awards,
        only_flagged=req.only_flagged,
        max_refs=req.max_refs,
        include_no_qc_line=True,
    )
    html_str = render_mdl_html(model)
    data = html_to_docx_bytes(html_str, force_basic=True)

    folder = (req.dest_path or "").lstrip("/")
    base = f"MDL-Preview-{sanitize(req.auditee_name)}-{sanitize(req.ein)}-{req.audit_year}.docx"
    blob_name = f"{folder}{base}" if folder else base
    url = upload_and_sas(AZURE_CONTAINER, blob_name, data) if AZURE_CONN_STR else save_local_and_url(blob_name, data)
    return {"url": url, "blob_path": f"{AZURE_CONTAINER}/{blob_name}", "size_bytes": len(data)}

from fastapi.responses import JSONResponse

@app.post("/build-mdl-docx-by-report-templated")
def build_mdl_docx_by_report_templated(req: BuildByReportTemplated):
    try:
        # 1) General
        fac_general = _fac_get("general", {
            "report_id": f"eq.{req.report_id}",
            "select": "report_id,fac_accepted_date",
            "limit": 1
        })

        # 2) Findings
        findings_params = {
            "report_id": f"eq.{req.report_id}",
            "select": "reference_number,award_reference,type_requirement,"
                      "is_material_weakness,is_significant_deficiency,is_questioned_costs,"
                      "is_modified_opinion,is_other_findings,is_other_matters,is_repeat_finding",
            "order": "reference_number.asc",
            "limit": str(req.max_refs)
        }
        if req.only_flagged:
            flagged = [
                "is_material_weakness","is_significant_deficiency","is_questioned_costs",
                "is_modified_opinion","is_other_findings","is_other_matters","is_repeat_finding"
            ]
            findings_params["or"] = "(" + ",".join([f"{f}.eq.true" for f in flagged]) + ")"
        fac_findings = _fac_get("findings", findings_params)

        # 3) refs
        refs = [row.get("reference_number") for row in fac_findings if row.get("reference_number")]
        refs = refs[: req.max_refs]

        # 4) texts & CAPs
        if refs:
            fac_findings_text = _fac_get("findings_text", {
                "report_id": f"eq.{req.report_id}",
                "select": "finding_ref_number,finding_text",
                "order": "finding_ref_number.asc",
                "limit": str(len(refs)),
                "or": _or_param("finding_ref_number", refs)
            })
            fac_caps = _fac_get("corrective_action_plans", {
                "report_id": f"eq.{req.report_id}",
                "select": "finding_ref_number,planned_action",
                "order": "finding_ref_number.asc",
                "limit": str(len(refs)),
                "or": _or_param("finding_ref_number", refs)
            })
        else:
            fac_findings_text, fac_caps = [], []

        # 5) awards
        federal_awards = []
        if req.include_awards:
            federal_awards = _fac_get("federal_awards", {
                "report_id": f"eq.{req.report_id}",
                "select": "award_reference,federal_program_name",
                "order": "award_reference.asc",
                "limit": "200"
            })

        # 6) model
        mdl_model = build_mdl_model_from_fac(
            auditee_name=req.auditee_name,
            ein=req.ein,
            audit_year=req.audit_year,
            fac_general=fac_general,
            fac_findings=fac_findings,
            fac_findings_text=fac_findings_text,
            fac_caps=fac_caps,
            federal_awards=federal_awards,
            only_flagged=req.only_flagged,
            max_refs=req.max_refs,
            include_no_qc_line=True,
        )

        # 7) header autofill (if provided)
        if req.fy_end_text:     mdl_model["period_end_text"] = req.fy_end_text
        if req.auditor_name:    mdl_model["auditor_name"]    = req.auditor_name
        if req.recipient_name:  mdl_model["auditee_name"]    = req.recipient_name
        if req.street_address:  mdl_model["street_address"]  = req.street_address
        if req.city:            mdl_model["city"]            = req.city
        if req.state:           mdl_model["state"]           = req.state
        if req.zip_code:        mdl_model["zip_code"]        = req.zip_code
        if req.poc_name:        mdl_model["poc_name"]        = req.poc_name
        if req.poc_title:       mdl_model["poc_title"]       = req.poc_title

        # 8) template path
        template_path = req.template_path or MDL_TEMPLATE_PATH
        if not template_path:
            return JSONResponse(status_code=200, content={"ok": False, "message": "Template path not provided (template_path or MDL_TEMPLATE_PATH)."})

        # 9) build docx
        try:
            data = build_docx_from_template(mdl_model, template_path=template_path)
        except HTTPException as e:
            return JSONResponse(status_code=200, content={"ok": False, "message": f"Template error: {e.detail}"})
        except Exception as e:
            return JSONResponse(status_code=200, content={"ok": False, "message": f"Unexpected template error: {e}"})

        # 10) upload
        folder = (req.dest_path or "").lstrip("/")
        base = f"MDL-{sanitize(req.auditee_name)}-{sanitize(req.ein)}-{req.audit_year}.docx"
        blob_name = f"{folder}{base}" if folder else base
        url = upload_and_sas(AZURE_CONTAINER, blob_name, data) if AZURE_CONN_STR else save_local_and_url(blob_name, data)
        return {"ok": True, "url": url, "blob_path": f"{AZURE_CONTAINER}/{blob_name}", "size_bytes": len(data)}

    except HTTPException as e:
        # Convert any other HTTPException into a soft error so AnythingLLM can surface the message
        return JSONResponse(status_code=200, content={"ok": False, "message": f"{e.status_code}: {e.detail}"})
    except Exception as e:
        return JSONResponse(status_code=200, content={"ok": False, "message": f"Unhandled error: {e}"})

# --- NEW: single-call builder that looks up report_id first -------------------
from pydantic import BaseModel

# --- Pydantic input model (replace your BuildAuto) ---
class BuildAuto(BaseModel):
    # required
    auditee_name: str
    ein: str
    audit_year: int

    # options (all optional)
    dest_path: Optional[str] = None
    max_refs: int = 15
    only_flagged: bool = False
    include_awards: bool = True
    treasury_listings: Optional[List[str]] = None

    # header overrides (all optional)
    recipient_name: Optional[str] = None
    fy_end_text: Optional[str] = None
    auditor_name: Optional[str] = None
    street_address: Optional[str] = None
    city: Optional[str] = None
    state: Optional[str] = None
    zip_code: Optional[str] = None
    poc_name: Optional[str] = None
    poc_title: Optional[str] = None

    # template & mappings (optional)
    template_path: Optional[str] = None
    aln_reference_xlsx: Optional[str] = None
    treasury_contact_email: Optional[str] = None

    # optional flags
    include_no_qc_line: bool = True
    include_no_cap_line: bool = False


from fastapi.responses import JSONResponse

def _from_fac_general(gen_rows):
    """
    Pulls reasonable defaults from FAC /general for headers.
    Expects select to include:
      auditee_address_line_1,auditee_city,auditee_state,auditee_zip,
      auditor_firm_name,fy_end_date
    """
    if not gen_rows:
        return {}
    g = gen_rows[0]
    # fy_end_date can be 'YYYY-MM-DD' -> convert to 'Month DD, YYYY' if present
    per_end = None
    try:
        from datetime import datetime
        if g.get("fy_end_date"):
            dt = datetime.fromisoformat(g["fy_end_date"])
            per_end = dt.strftime("%B %-d, %Y") if hasattr(dt, "strftime") else None
    except Exception:
        pass

    return {
        "street_address": g.get("auditee_address_line_1") or "",
        "city": g.get("auditee_city") or "",
        "state": g.get("auditee_state") or "",
        "zip_code": g.get("auditee_zip") or "",
        "auditor_name": g.get("auditor_firm_name") or "",
        "period_end_text": per_end,
        "poc_name":  g.get("auditee_contact_name")  or "",
        "poc_title": g.get("auditee_contact_title") or "",
    }

def _unset_all_caps_everywhere(doc):
    # body paragraphs
    for p in doc.paragraphs:
        for r in p.runs:
            r.font.all_caps = False
            r.font.small_caps = False
    # tables
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.all_caps = False
                        r.font.small_caps = False
    # headers/footers
    for sec in doc.sections:
        for container in (sec.header, sec.footer):
            for p in container.paragraphs:
                for r in p.runs:
                    r.font.all_caps = False
                    r.font.small_caps = False

def _rewrite_paragraph(p, text):
    _clear_runs(p); p.add_run(text)

def _iter_all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for sec in doc.sections:
        for container in (sec.header, sec.footer):
            for p in container.paragraphs:
                yield p

def _fix_treasury_email(doc, email: str):
    if not email:
        return
    email = email.strip()

    found_token = False
    # 1) Replace tokens anywhere (body, tables, hdr/ftr)
    for p in _iter_all_paragraphs(doc):
        t = _para_text(p)
        if "${treasury_contact_email}" in t:
            found_token = True
            new_t = t.replace("${treasury_contact_email}", email)
            # tidy double spaces and missing space before next sentence
            new_t = new_t.replace(f"{email}.The", f"{email}. The")
            new_t = re.sub(r"\s{2,}", " ", new_t)
            _rewrite_paragraph(p, new_t)

    if found_token:
        return

    # 2) No token in template → inject into the “For questions…” line only
    target = _find_para_by_contains(doc, "For questions regarding the audit finding")
    if target:
        t = _para_text(target)
        if email not in t:
            new_t = re.sub(r"(?i)(please email us at)(\s*)", rf"\1 {email}. ", t, count=1)
            _rewrite_paragraph(target, new_t)
def _strip_leading_token_artifacts(doc):
    pat = re.compile(r"^\s*\$\{[^}]+\}\.?\s*")
    for p in _iter_all_paragraphs(doc):
        t = _para_text(p)
        if not t:
            continue
        new = pat.sub("", t)
        if new != t:
            _clear_runs(p); p.add_run(new)

@app.post("/build-mdl-docx-auto")
def build_mdl_docx_auto(req: BuildAuto):
    try:
        # 1) Find newest report_id for EIN/year
        gen = _fac_get("general", {
            "audit_year": f"eq.{req.audit_year}",
            "auditee_ein": f"eq.{req.ein}",
            "select": "report_id, fac_accepted_date, auditee_address_line_1, auditee_city, auditee_state, auditee_zip, auditor_firm_name, fy_end_date, auditee_contact_name,auditee_contact_title",
            "order": "fac_accepted_date.desc",
            "limit": 1
        })
        if not gen:
            return {"ok": False, "message": f"No FAC report found for EIN {req.ein} in {req.audit_year}."}

        report_id = gen[0]["report_id"]
        try:
            aln_by_award, aln_by_finding = _aln_overrides_from_summary(report_id)
        except Exception:
            aln_by_award, aln_by_finding = {}, {}

        # 2) (unchanged) fetch findings / texts / caps / awards ...
        #    ... your existing code here ...

        # 2) Fetch findings / finding text / CAPs (ALWAYS initialize lists)
        findings_params = {
            "report_id": f"eq.{report_id}",
            "select": ("reference_number,award_reference,type_requirement,"
                    "is_material_weakness,is_significant_deficiency,is_questioned_costs,"
                    "is_modified_opinion,is_other_findings,is_other_matters,is_repeat_finding"),
            "order": "reference_number.asc",
            "limit": str(req.max_refs or 15),
        }
        if req.only_flagged:
            flagged = [
                "is_material_weakness", "is_significant_deficiency", "is_questioned_costs",
                "is_modified_opinion", "is_other_findings", "is_other_matters", "is_repeat_finding",
            ]
            findings_params["or"] = "(" + ",".join(f"{f}.eq.true" for f in flagged) + ")"

        try:
            fac_findings = _fac_get("findings", findings_params) or []
        except Exception:
            fac_findings = []

        # refs are safe even if no findings
        refs = [r.get("reference_number") for r in fac_findings if r.get("reference_number")]
        refs = refs[: (req.max_refs or 15)]

        # Always define these
        if refs:
            try:
                fac_findings_text = _fac_get("findings_text", {
                    "report_id": f"eq.{report_id}",
                    "select": "finding_ref_number,finding_text",
                    "order": "finding_ref_number.asc",
                    "limit": str(len(refs)),
                    "or": _or_param("finding_ref_number", refs),
                }) or []
            except Exception:
                fac_findings_text = []
            try:
                fac_caps = _fac_get("corrective_action_plans", {
                    "report_id": f"eq.{report_id}",
                    "select": "finding_ref_number,planned_action",
                    "order": "finding_ref_number.asc",
                    "limit": str(len(refs)),
                    "or": _or_param("finding_ref_number", refs),
                }) or []
            except Exception:
                fac_caps = []
        else:
            fac_findings_text, fac_caps = [], []

        # Awards (optional, still safe)
        federal_awards = []
        if req.include_awards:
            try:
                federal_awards = _fac_get("federal_awards", {
                    "report_id": f"eq.{report_id}",
                    "select": "award_reference,federal_program_name,assistance_listing",
                    "order": "award_reference.asc",
                    "limit": "200",
                }) or []
            except Exception:
                federal_awards = []
        if federal_awards:
            for a in federal_awards:
                if not (a.get("assistance_listing") or "").strip():
                    ar = (a.get("award_reference") or "").strip()
                    if ar and ar in aln_by_award:
                        a["assistance_listing"] = aln_by_award[ar]
        # ---------- NEW: build the model -------------
        mdl_model = build_mdl_model_from_fac(
            auditee_name=req.auditee_name,
            ein=req.ein,
            audit_year=req.audit_year,
            fac_general=gen,
            fac_findings=fac_findings,
            fac_findings_text=fac_findings_text,
            fac_caps=fac_caps,
            federal_awards=federal_awards,
            only_flagged=req.only_flagged,
            max_refs=req.max_refs,
            include_no_qc_line=True,
            treasury_listings=req.treasury_listings,
            aln_reference_xlsx=req.aln_reference_xlsx,
            aln_overrides_by_finding=aln_by_finding,
        )

        # ---------- NEW: enrich headers from FAC + defaults ----------
        fac_defaults = _from_fac_general(gen)

        def _normalize_auditor_name(name: str) -> str:
            if not name:
                return ""
            clean = name.strip()
            return clean if clean.lower().startswith("the ") else f"the {clean}"

        recipient = _title_with_article(req.recipient_name or req.auditee_name)

        header_overrides = {
            # recipient & period end
            "recipient_name": recipient,
            "period_end_text": req.fy_end_text or fac_defaults.get("period_end_text") or mdl_model.get("period_end_text"),

            # address (title case street + city, uppercase state, keep zip as-is)
            "street_address": _title_case(req.street_address or fac_defaults.get("street_address")),
            "city": _title_case(req.city or fac_defaults.get("city")),
            "state": (req.state or fac_defaults.get("state") or "").upper(),
            "zip_code": req.zip_code or fac_defaults.get("zip_code") or "",

            # auditor
            "auditor_name": _normalize_auditor_name(req.auditor_name or fac_defaults.get("auditor_name") or ""),

            # POC (title case name + title)
            "poc_name": _title_case(req.poc_name or fac_defaults.get("poc_name")),
            "poc_title": _title_case(req.poc_title or fac_defaults.get("poc_title")),
        }

        # apply non-empty values only
        for k, v in header_overrides.items():
            if v:
                mdl_model[k] = v

        # ------------- sensible defaults for things the caller omitted -------------
        # Treasury listings: if not provided, use the SLFRF + common Treasury programs for demo
        if not req.treasury_listings:
            req.treasury_listings = ["21.027", "21.023", "21.026"]

        # Template defaults if not provided
        template_path = _none_if_placeholder(req.template_path) or "templates/MDL_Template_Data_Mapping_Comments.docx"
        aln_xlsx      = _none_if_placeholder(req.aln_reference_xlsx) or "templates/Additional_Reference_Documentation_MDLs.xlsx"

        # Pass mapping workbook path into the model builder via existing parameter if you support it
        # (If build_mdl_model_from_fac already accepts aln_reference_xlsx, we've passed it above.)

        # Destination folder defaults
        dest_folder = _str_or_default(req.dest_path, f"mdl/{req.audit_year}/").lstrip("/")

        # 4) Build DOCX (unchanged except variable names)
        try:
            data = build_docx_from_template(mdl_model, template_path=template_path)
        except HTTPException as e:
            return {"ok": False, "message": f"Template error: {e.detail}"}
        except Exception as e:
            return {"ok": False, "message": f"Unexpected template error: {e}"}

        # 5) Upload (unchanged)
        base = f"MDL-{sanitize(req.auditee_name)}-{sanitize(req.ein)}-{req.audit_year}.docx"
        blob_name = f"{dest_folder}{base}" if dest_folder else base
        url = upload_and_sas(AZURE_CONTAINER, blob_name, data) if AZURE_CONN_STR else save_local_and_url(blob_name, data)

        return {"ok": True, "url": url, "blob_path": f"{AZURE_CONTAINER}/{blob_name}"}

    except HTTPException as e:
        return JSONResponse(status_code=200, content={"ok": False, "message": f"{e.status_code}: {e.detail}"})
    except Exception as e:
        return JSONResponse(status_code=200, content={"ok": False, "message": f"Unhandled error: {e}"})

from fastapi import Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse

@app.middleware("http")
async def log_requests(request: Request, call_next):
    if request.url.path.endswith("/build-mdl-docx-auto"):
        raw = await request.body()
        try:
            logging.info("== /build-mdl-docx-auto RAW BODY ==")
            logging.info(raw.decode("utf-8", errors="ignore"))
        except Exception:
            pass
        # re-create the request stream for downstream
        request._receive = (lambda b=raw: {"type": "http.request", "body": b, "more_body": False})
    return await call_next(request)

@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    logging.info("== Pydantic Validation Errors ==")
    logging.info(exc.errors())
    return JSONResponse(status_code=422, content={"ok": False, "errors": exc.errors()})        