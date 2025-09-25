import csv, json, os, re, requests, pathlib

DOCX_BASE = os.getenv("DOCX_BASE", "http://127.0.0.1:8000")
OUTDIR = pathlib.Path("./out"); OUTDIR.mkdir(exist_ok=True)

def sanitize(s): return re.sub(r'[^A-Za-z0-9_.-]+', '_', s or '')

def get_report_id(ein: str, year: int):
    url = f"https://api.fac.gov/general?audit_year=eq.{year}&auditee_ein=eq.{ein}&select=report_id,fac_accepted_date&order=fac_accepted_date.desc&limit=1"
    r = requests.get(url, timeout=30); r.raise_for_status()
    data = r.json()
    return (data[0]["report_id"] if data else None)

def build_docx(name: str, ein: str, year: int, rid: str):
    body = {
        "auditee_name": name,
        "ein": ein,
        "audit_year": year,
        "report_id": rid,
        "dest_path": f"mdl/{year}/",
        "max_refs": 2,
        "only_flagged": False,       # set True if you only want flagged items
        "include_awards": False
    }
    r = requests.post(f"{DOCX_BASE}/build-docx-by-report",
                      json=body,
                      headers={"X-Flow-Trace": "batch"},
                      timeout=120)
    r.raise_for_status()
    return r.json()

with open("cases.csv") as f:
    for name, ein, year in csv.reader(f):
        name = name.strip()
        ein  = re.sub(r'[^0-9]', '', ein)
        year = int(year.strip())
        print(f"----\nAuditee: {name} | EIN: {ein} | Year: {year}")

        rid = get_report_id(ein, year)
        if not rid:
            print("No FAC report_id found — skipping.")
            continue
        print("report_id:", rid)

        resp = build_docx(name, ein, year, rid)
        url  = resp.get("url")
        size = resp.get("size_bytes")

        if not url:
            print("Build failed:", resp); continue

        outfile = OUTDIR / f"MDL-{sanitize(name)}-{ein}-{year}.docx"
        with requests.get(url, stream=True, timeout=120) as dl:
            dl.raise_for_status()
            with open(outfile, "wb") as fh:
                for chunk in dl.iter_content(8192):
                    fh.write(chunk)
        print(f"Saved: {outfile} ({size} bytes)")
