#!/usr/bin/env python3
# Reconstrueert build_v3.py en genereert Vermogensmonitor_Premium.xlsx (v3).
# Vereist: build_v2.py in dezelfde map (staat in deze repo) + openpyxl.
# Gebruik:  python rebuild_build_v3.py
import base64, gzip, runpy, pathlib, hashlib

SHA256_SURGERY = "cb443e1f3903ae691a86ab99f62f113b2aefac4bbe8fd089cef075487d2ec01e"

b64 = pathlib.Path("build_v3_surgery.gz.b64").read_text().strip()
data = gzip.decompress(base64.b64decode(b64))
if hashlib.sha256(data).hexdigest() != SHA256_SURGERY:
    raise SystemExit("Integriteitscheck mislukt: build_v3_surgery.gz.b64 is beschadigd.")
pathlib.Path("build_v3_surgery.py").write_bytes(data)
print("build_v3_surgery.py hersteld (SHA-256 OK).")
runpy.run_path("build_v3_surgery.py")   # leest build_v2.py, schrijft build_v3.py
runpy.run_path("build_v3.py")           # schrijft Vermogensmonitor_Premium.xlsx
print("Klaar: build_v3.py en Vermogensmonitor_Premium.xlsx gegenereerd.")
