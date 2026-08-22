#!/usr/bin/env python3
"""postprocess.py — pptxgenjs 산출물의 차트 결함 보정. **선택이 아니라 필수다.**

pptxgenjs 4.x는 차트에 두 가지 결함을 남긴다.
  1) 선언하지 않은 세 번째 <c:axId>를 plot 안에 써 넣는다. 존재하지 않는 축을 가리키므로
     PowerPoint가 파일을 손상으로 판단하고 "복구하시겠습니까" 대화상자를 띄운다.
     LibreOffice는 조용히 무시하므로 렌더 QA로는 드러나지 않는다.
  2) <c:ser> 안에서 <c:dPt>를 <c:dLbls> 뒤에 쓴다 (ISO 자식 순서 위반).
또한 차트 파트에는 <a:latin>만 쓰므로 한글 축·데이터 라벨이 테마 폰트로 떨어질 수 있다.
같은 서체의 <a:ea>를 주입한다.

    python3 postprocess.py deck.pptx [-o out.pptx]
"""
import argparse
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

LATIN_RE = re.compile(r'<a:latin typeface="([^"]+)"([^/>]*)/>')
AX_BLOCK_RE = re.compile(r"<c:(catAx|valAx|serAx|dateAx)>.*?</c:\1>", re.S)
AXID_RE = re.compile(r'<c:axId val="(\d+)"\s*/>')
SER_RE = re.compile(r"<c:ser>.*?</c:ser>", re.S)
DPT_RE = re.compile(r"<c:dPt>.*?</c:dPt>", re.S)


def inject_ea(text):
    """차트 런의 <a:latin> 뒤에 같은 서체의 <a:ea>를 붙인다(cs는 넣지 않는다 — 중복 위험)."""
    if "<a:ea typeface=" in text:
        return text, 0
    count = 0

    def repl(m):
        nonlocal count
        count += 1
        return f'<a:latin typeface="{m.group(1)}"{m.group(2)}/><a:ea typeface="{m.group(1)}"{m.group(2)}/>'

    return LATIN_RE.sub(repl, text), count


def fix_axes(text):
    """선언되지 않은 <c:axId> 참조를 제거한다."""
    declared = set()
    for m in AX_BLOCK_RE.finditer(text):
        declared.update(AXID_RE.findall(m.group(0)))
    if not declared:
        return text, 0
    blocks = []

    def stash(m):
        blocks.append(m.group(0))
        return f"@@AX{len(blocks) - 1}@@"

    body = AX_BLOCK_RE.sub(stash, text)
    removed = 0

    def drop(m):
        nonlocal removed
        if m.group(1) in declared:
            return m.group(0)
        removed += 1
        return ""

    body = AXID_RE.sub(drop, body)
    for i, b in enumerate(blocks):
        body = body.replace(f"@@AX{i}@@", b)
    return body, removed


def fix_ser_order(text):
    """<c:ser> 안의 <c:dPt>를 <c:dLbls> 앞으로 옮긴다."""
    moved = 0

    def fix(m):
        nonlocal moved
        ser = m.group(0)
        d = ser.find("<c:dLbls>")
        if d == -1:
            return ser
        pts = [p.group(0) for p in DPT_RE.finditer(ser) if p.start() > d]
        if not pts:
            return ser
        for b in pts:
            ser = ser.replace(b, "", 1)
        moved += len(pts)
        at = ser.find("<c:dLbls>")
        return ser[:at] + "".join(pts) + ser[at:]

    return SER_RE.sub(fix, text), moved


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out", help="출력 경로 (생략 시 제자리 교체)")
    args = ap.parse_args()
    src = Path(args.pptx)
    if not src.exists():
        print(f"파일 없음: {src}", file=sys.stderr)
        return 2
    dst = Path(args.out) if args.out else src

    ea = ax = dpt = 0
    tmp = Path(tempfile.mkdtemp()) / "out.pptx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/charts/chart") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                text, a = inject_ea(text)
                text, b = fix_axes(text)
                text, c = fix_ser_order(text)
                ea, ax, dpt = ea + a, ax + b, dpt + c
                if a or b or c:
                    data = text.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(str(tmp), str(dst))

    if ea:
        print(f"차트 한글 폰트(<a:ea>) {ea}곳 주입")
    if ax:
        print(f"미선언 축 참조(<c:axId>) {ax}개 제거 — PowerPoint 손상 경고 원인")
    if dpt:
        print(f"<c:dPt> {dpt}개를 <c:dLbls> 앞으로 이동 (ISO 순서)")
    if not (ea or ax or dpt):
        print("차트 파트에 고칠 것이 없다")
    print(f"→ {dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
