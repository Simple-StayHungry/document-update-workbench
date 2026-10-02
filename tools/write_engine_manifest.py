#!/usr/bin/env python3
"""Regenerate engine_manifest.json from the current engine files.

The manifest is a generated artefact. Never hand-edit it, and never "fix" a
drift by pasting in new hashes without first explaining why the file changed.

Usage:
    python tools/write_engine_manifest.py            # show drift, then write
    python tools/write_engine_manifest.py --check     # exit 1 if out of date
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "engine_manifest.json"

sys.path.insert(0, str(ROOT))
from workbench.version import BUILD, VERSION  # noqa: E402  (path set above)

# Deltas that the public release intentionally introduces relative to the
# production build the manifest was originally frozen against. Each entry is a
# fact about a past edit, not something that can be recomputed, so it is declared
# here and preserved verbatim by the generator.
PUBLIC_RELEASE_DELTAS = [
    {
        "path": "workbench/engine.py",
        "baseline_sha256": "2773ce4376e328182b7ba1655759c2bcfca889845bc2bc7a78e844ab42052f66",
        "delta": "用户可见提示文案：老师定稿 → 示例定稿",
        "impact": "仅错误提示的措辞，不参与任何判定逻辑",
        "reason": "公开版不出现内部称呼",
    },
    {
        "path": "workbench/model.py",
        "baseline_sha256": "3ff4b294963b84037a9b9bed570f8c1024a86c26768694559cc8b9964a905092",
        "delta": "正则字面量：经广发证券对 → 经示例证券对",
        "impact": "该替换收窄了「独立核查程序与结论保留」保护规则的触发面；是否改为泛化写法（如 经\\S{2,10}证券对）待定",
        "reason": "公开版不出现具体承销商名称",
    },
]


def engine_files() -> list[Path]:
    """Every module under workbench/, sorted — the engine surface we freeze."""
    return sorted((ROOT / "workbench").glob("*.py"))


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def build_manifest() -> dict:
    return {
        "build": BUILD,
        "version": VERSION,
        "version_source": "workbench/version.py",
        "generator": "tools/write_engine_manifest.py",
        "files": [
            {"path": str(p.relative_to(ROOT)).replace("\\", "/"), "sha256": digest(p)}
            for p in engine_files()
        ],
        "public_release_deltas": PUBLIC_RELEASE_DELTAS,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="exit 1 if the manifest is out of date")
    args = ap.parse_args()

    fresh = build_manifest()
    text = json.dumps(fresh, ensure_ascii=False, indent=2) + "\n"

    old = None
    if MANIFEST.exists():
        try:
            old = json.loads(MANIFEST.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            old = None

    if old == fresh:
        print(f"manifest is up to date: {BUILD} ({len(fresh['files'])} files)")
        return 0

    if old:
        before = {e["path"]: e["sha256"] for e in old.get("files", [])}
        after = {e["path"]: e["sha256"] for e in fresh["files"]}
        for path in sorted(set(before) | set(after)):
            if before.get(path) != after.get(path):
                b = (before.get(path) or "(新增)")[:12]
                a = (after.get(path) or "(移除)")[:12]
                print(f"  changed  {path:<34} {b} -> {a}")
        if old.get("build") != fresh["build"]:
            print(f"  build    {old.get('build')!r} -> {fresh['build']!r}")

    if args.check:
        print("engine_manifest.json is out of date; run tools/write_engine_manifest.py")
        return 1

    MANIFEST.write_text(text, encoding="utf-8")
    print(f"wrote {MANIFEST.relative_to(ROOT)}: {BUILD} ({len(fresh['files'])} files)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
