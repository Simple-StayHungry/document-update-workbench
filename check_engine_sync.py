from pathlib import Path
import hashlib,json,sys
root=Path(__file__).resolve().parent
m=json.loads((root/"engine_manifest.json").read_text(encoding="utf-8"))
bad=[]
for rec in m["files"]:
    p=root/rec["path"]
    got=hashlib.sha256(p.read_bytes()).hexdigest() if p.exists() else None
    if got!=rec["sha256"]: bad.append((rec["path"],rec["sha256"],got))
if bad:
    print("ENGINE DRIFT")
    for x in bad: print(*x)
    sys.exit(1)
print("ENGINE SNAPSHOT OK",len(m["files"]),"files",m.get("build",""))
