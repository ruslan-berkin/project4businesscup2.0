"""Refresh the standalone app fallback from the canonical checked-in assets."""
import base64
import json
import zlib
from pathlib import Path

root = Path(__file__).resolve().parents[1]
assets = {}
for folder in ("templates", "static"):
    for path in sorted((root / folder).rglob("*")):
        if path.is_file():
            assets[path.relative_to(root).as_posix()] = base64.b64encode(zlib.compress(path.read_bytes())).decode()
app = root / "app.py"
source = app.read_text()
start = source.index("_EMBEDDED_FILES = {")
end = source.index("\n\n\ndef _ensure_runtime_files", start)
app.write_text(source[:start] + "_EMBEDDED_FILES = " + json.dumps(assets, indent=4) + source[end:])
