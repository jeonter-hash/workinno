#!/usr/bin/env python3
"""apply_gradient.py — Hero Gradient 후처리.

pptxgenjs는 그라디언트 채우기를 지원하지 않는다. 헬퍼가 센티넬 색(0A0B0C)으로
채워 둔 도형을 찾아 벡터 gradFill(linear-gradient(135deg, #1456f0 0%, #3b82f6 50%,
#60a5fa 100%))로 바꾼다. 이미지로 굽지 않으므로 확대·인쇄 품질이 유지된다.

    python3 apply_gradient.py deck.pptx [-o out.pptx]

옵션 없이 쓰면 제자리에서 교체한다. 장표당 1개·덱 전체 3개 상한을 검사해서
넘으면 실패시킨다(디자인 시스템 규칙). XML은 문자열 치환만 하므로 네임스페이스
접두사가 변형될 위험이 없다.
"""
import argparse
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

SENTINEL = "0A0B0C"
SOLID_RE = re.compile(
    r'<a:solidFill>\s*<a:srgbClr val="%s"\s*/>\s*</a:solidFill>' % SENTINEL, re.I
)
GRADIENT = (
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:srgbClr val="1456F0"/></a:gs>'
    '<a:gs pos="50000"><a:srgbClr val="3B82F6"/></a:gs>'
    '<a:gs pos="100000"><a:srgbClr val="60A5FA"/></a:gs>'
    '</a:gsLst><a:lin ang="2700000" scaled="0"/></a:gradFill>'
)
MAX_PER_SLIDE = 1
MAX_PER_DECK = 3


def slide_no(name: str) -> int:
    m = re.search(r"slide(\d+)\.xml$", name)
    return int(m.group(1)) if m else 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out", help="출력 경로 (생략 시 제자리 교체)")
    ap.add_argument("--allow-over-limit", action="store_true",
                    help="상한 초과를 경고로만 처리 (기본은 실패)")
    args = ap.parse_args()

    src = Path(args.pptx)
    if not src.exists():
        print(f"파일 없음: {src}", file=sys.stderr)
        return 2
    dst = Path(args.out) if args.out else src

    total = 0
    per_slide = {}
    tmp = Path(tempfile.mkdtemp()) / "out.pptx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/slides/slide") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                text, n = SOLID_RE.subn(GRADIENT, text)
                if n:
                    per_slide[slide_no(item.filename)] = n
                    total += n
                    data = text.encode("utf-8")
            zout.writestr(item, data)

    over = [f"slide{k}({v}개)" for k, v in sorted(per_slide.items()) if v > MAX_PER_SLIDE]
    problems = []
    if over:
        problems.append("장표당 Hero Gradient 1개 초과: " + ", ".join(over))
    if total > MAX_PER_DECK:
        problems.append(f"덱 전체 Hero Gradient {total}개 — 상한 {MAX_PER_DECK}개 초과")
    if problems and not args.allow_over_limit:
        for p in problems:
            print("FAIL: " + p, file=sys.stderr)
        print("디자인 시스템 위반이다. 그라디언트를 줄이고 다시 생성할 것.", file=sys.stderr)
        return 1
    for p in problems:
        print("WARN: " + p)

    shutil.move(str(tmp), str(dst))
    if total == 0:
        print(f"센티넬({SENTINEL}) 도형 없음 — 교체할 그라디언트가 없다: {dst}")
    else:
        detail = ", ".join(f"slide{k}:{v}" for k, v in sorted(per_slide.items()))
        print(f"Hero Gradient {total}개 적용 ({detail}) → {dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
