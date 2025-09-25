# save as fix_py.sh and run:  bash fix_py.sh
set -euo pipefail

# Prefer Homebrew ARM Python if present, else python.org universal2, else system python3
CANDIDATES=(
  /opt/homebrew/bin/python3.12
  /opt/homebrew/bin/python3.11
  /Library/Frameworks/Python.framework/Versions/3.12/bin/python3
  /Library/Frameworks/Python.framework/Versions/3.11/bin/python3
  "$(command -v python3 || true)"
)

PY=""
for c in "${CANDIDATES[@]}"; do
  [ -x "$c" ] || continue
  A=$(ARCHPREFERENCE=arm64 "$c" -c 'import platform; print(platform.machine())' 2>/dev/null || true)
  if [ "$A" = "arm64" ]; then PY="$c"; break; fi
done

if [ -z "$PY" ]; then
  echo "No ARM64 Python found. You can either:
  - Install Homebrew Python:   /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\" && brew install python@3.12
  - Or install python.org 3.11: https://www.python.org/downloads/macos/
Then re-run this script." >&2
  exit 1
fi

echo "Using Python: $PY"
"$PY" -V

# Rebuild venv cleanly
deactivate 2>/dev/null || true
rm -rf .venv
"$PY" -m venv .venv
source .venv/bin/activate

# Sanity check: must be arm64
python - <<'PY'
import platform, sys
print("Python:", sys.version)
print("Arch:   ", platform.machine())
assert platform.machine() == "arm64", "Not running on arm64 interpreter"
PY

# Fresh native installs
pip install --upgrade pip setuptools wheel
pip cache purge

# Force-reinstall pydantic core as native ARM64
pip install --no-cache-dir --force-reinstall pydantic-core pydantic

# Your app deps (add/remove as needed)
pip install --no-cache-dir fastapi uvicorn requests python-docx html2docx beautifulsoup4 azure-storage-blob

# Verify pydantic_core wheel is arm64
python - <<'PY'
import pydantic_core, subprocess, sys
print("pydantic_core file:", pydantic_core.__file__)
subprocess.run(["file", pydantic_core.__file__], check=False)
PY

echo "✅ Done. Activate with:  source .venv/bin/activate"
echo "Run server:            uvicorn main:app --host 0.0.0.0 --port 8000 --reload"
