#!/usr/bin/env python3
"""check_skill_sync.py — PRESENT_PPT와 REPORT_PPT의 공용 파일이 어긋났는지 본다.

    python3 tools/check_skill_sync.py

두 스킬은 모듈·점검기·설계 문서를 같은 내용으로 갖는다. 코워크에 올릴 때는 각 스킬
폴더만 zip으로 나가므로 저장소 밖에 원본을 둘 수 없고, 두 벌을 같게 유지해야 한다.
한쪽만 고치면 이 스크립트가 잡는다. 커밋 전에 돌릴 것.
"""
import hashlib
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent / ".claude" / "skills"
SKILLS = ("PRESENT_PPT", "REPORT_PPT")
SHARED = (
    "references/diagrams.md",
    "assets/ksa_mono.js",
    "assets/photo_placeholder.png",
    "scripts/postprocess.py",
    "scripts/check_layout.py",
    "scripts/measure_head.py",
    "references/headline.md",
    "references/design-system.md",
)


def digest(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()[:12] if p.exists() else None


def main():
    bad = []
    for rel in SHARED:
        ds = {sk: digest(ROOT / sk / rel) for sk in SKILLS}
        if None in ds.values():
            missing = [sk for sk, d in ds.items() if d is None]
            bad.append(f"없음  {rel} — {', '.join(missing)}")
        elif len(set(ds.values())) > 1:
            bad.append(f"불일치 {rel} — " + " / ".join(f"{sk} {d}" for sk, d in ds.items()))
        else:
            print(f"같음  {rel}  {next(iter(ds.values()))}")
    if bad:
        print()
        for b in bad:
            print("FAIL " + b)
        print(f"\n{len(bad)}건 어긋남 — 한쪽을 다른 쪽으로 복사할 것")
        return 1
    print(f"\n공용 파일 {len(SHARED)}건 모두 같음")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
