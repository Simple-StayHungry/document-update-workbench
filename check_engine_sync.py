from pathlib import Path
import hashlib,json,sys
root=Path(__file__).resolve().parent
sys.path.insert(0,str(root))
from workbench.version import BUILD,VERSION

m=json.loads((root/"engine_manifest.json").read_text(encoding="utf-8"))
bad=[]
for rec in m["files"]:
    p=root/rec["path"]
    got=hashlib.sha256(p.read_bytes()).hexdigest() if p.exists() else None
    if got!=rec["sha256"]: bad.append((rec["path"],rec["sha256"],got))
if bad:
    print("ENGINE DRIFT")
    for x in bad: print(*x)
    print("先说明每个文件为什么变，再用 tools/write_engine_manifest.py 重新生成；不要手改 hash。")
    sys.exit(1)

# The version lives in exactly one place (workbench/version.py). Everything else
# reads it, so the manifest disagreeing with it is a real defect, not a warning.
if m.get("build")!=BUILD or m.get("version")!=VERSION:
    print("VERSION DRIFT")
    print("  manifest         ",m.get("build"),m.get("version"))
    print("  workbench/version",BUILD,VERSION)
    print("运行 tools/write_engine_manifest.py 重新生成。")
    sys.exit(1)

print("ENGINE SNAPSHOT OK",len(m["files"]),"files",BUILD)
