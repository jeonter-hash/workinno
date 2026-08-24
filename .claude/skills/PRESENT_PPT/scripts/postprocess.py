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



COVER_LEAN = 0.2126        # 표지 띠 사선 기울기 tan(12°) — ksa_mono.js COVER.lean 과 같아야 한다
PIC_RE = re.compile(r"<p:pic>.*?</p:pic>", re.S)


def shear_cover(text):
    """표지 띠 사진을 평행사변형으로 바꾼다.

    pptxgenjs는 그림을 사각형으로만 넣는다. altText가 KSA_COVER_BAND 로 시작하는
    그림의 도형을 parallelogram 으로 바꿔 사선 분할을 만든다. 그림 도형이므로
    PowerPoint에서 [그림 바꾸기]를 해도 사선이 유지된다.
    """
    n = 0

    def one(m):
        nonlocal n
        pic = m.group(0)
        if 'descr="KSA_COVER_BAND' not in pic:
            return pic
        ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', pic)
        if not ext:
            return pic
        cx, cy = int(ext.group(1)), int(ext.group(2))
        adj = int(round(cy * COVER_LEAN / min(cx, cy) * 100000))
        adj = max(0, min(adj, int(100000 * cx / min(cx, cy))))
        new = ('<a:prstGeom prst="parallelogram"><a:avLst>'
               f'<a:gd name="adj" fmla="val {adj}"/></a:avLst></a:prstGeom>')
        out, k = re.subn(r'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>', new, pic, count=1)
        n += k
        return out

    return PIC_RE.sub(one, text), n



FONT_OK = "맑은 고딕"
FACE_RE = re.compile(r"<a:(latin|ea|cs) typeface=\"([^\"]+)\"")


def fix_chart_font(text):
    """차트 파트의 다른 서체를 맑은 고딕으로 바꾼다.

    pptxgenjs는 옵션으로 덮지 못하는 자리(축 제목·보조축 등)에 Arial 18pt를 박아 넣는다.
    옵션을 하나씩 채워 넣기보다 여기서 일괄로 고치는 편이 확실하다.
    """
    n = 0

    def one(m):
        nonlocal n
        tag, face = m.group(1), m.group(2)
        if face.startswith("+") or face.startswith(FONT_OK):
            return m.group(0)
        n += 1
        return f'<a:{tag} typeface="{FONT_OK}"'

    return FACE_RE.sub(one, text), n


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

    ea = ax = dpt = shear = face = 0
    tmp = Path(tempfile.mkdtemp()) / "out.pptx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", item.filename):
                text = data.decode("utf-8")
                text, d = shear_cover(text)
                shear += d
                if d:
                    data = text.encode("utf-8")
            if item.filename.startswith("ppt/charts/chart") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                text, a = inject_ea(text)
                text, b = fix_axes(text)
                text, c = fix_ser_order(text)
                text, d = fix_chart_font(text)
                ea, ax, dpt, face = ea + a, ax + b, dpt + c, face + d
                if a or b or c or d:
                    data = text.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(str(tmp), str(dst))

    if face:
        print(f"차트 서체 {face}곳을 맑은 고딕으로 교체")
    if shear:
        print(f"표지 띠 사진 {shear}장을 평행사변형으로 변경")
    if ea:
        print(f"차트 한글 폰트(<a:ea>) {ea}곳 주입")
    if ax:
        print(f"미선언 축 참조(<c:axId>) {ax}개 제거 — PowerPoint 손상 경고 원인")
    if dpt:
        print(f"<c:dPt> {dpt}개를 <c:dLbls> 앞으로 이동 (ISO 순서)")
    if not (ea or ax or dpt or shear or face):
        print("차트 파트에 고칠 것이 없다")
    print(f"→ {dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
