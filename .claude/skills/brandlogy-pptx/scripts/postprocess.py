#!/usr/bin/env python3
"""postprocess.py — Hero Gradient + 한글 폰트 후처리.

두 가지를 고친다.
  1) 그라디언트 — pptxgenjs는 그라디언트 채우기를 지원하지 않는다. 헬퍼가 센티넬
     색(0A0B0C)으로 채워 둔 도형을 찾아 벡터 gradFill(135deg, #1456f0 → #3b82f6 →
     #60a5fa)로 바꾼다. 이미지로 굽지 않으므로 확대·인쇄 품질이 유지된다.
  2) 차트 한글 폰트 — 슬라이드 런에는 pptxgenjs가 latin/ea/cs를 모두 써 주지만
     차트 파트(ppt/charts/*.xml)에는 <a:latin>만 쓴다. 한글 축·데이터 라벨이
     테마 폰트로 떨어지지 않도록 같은 서체의 <a:ea>를 주입한다.
  3) 차트 구조 결함(pptxgenjs 4.x) — PowerPoint가 파일을 거부하는 원인:
     · 선언되지 않은 세 번째 <c:axId>를 plot 안에 쓴다 → 없는 축을 가리켜 "복구" 대화상자
     · <c:ser> 안에서 <c:dPt>를 <c:dLbls> 뒤에 쓴다 → ISO 순서 위반
     둘 다 바로잡는다.

    python3 postprocess.py deck.pptx [-o out.pptx]

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
FONT = "맑은 고딕"
LATIN_RE = re.compile(r'<a:latin typeface="([^"]+)"([^/>]*)/>')


def slide_no(name: str) -> int:
    m = re.search(r"slide(\d+)\.xml$", name)
    return int(m.group(1)) if m else 0


def inject_ea(text: str) -> tuple:
    """차트 파트의 <a:latin> 뒤에 같은 서체의 <a:ea>를 붙인다.

    <a:cs>는 넣지 않는다 — 원본에 이미 cs가 있으면 중복이 되어 스키마를 깨뜨린다.
    한글 렌더링에 필요한 것은 ea뿐이다.
    """
    count = 0

    def repl(m):
        nonlocal count
        face, attrs = m.group(1), m.group(2)
        count += 1
        return f'<a:latin typeface="{face}"{attrs}/><a:ea typeface="{face}"{attrs}/>'

    if "<a:ea typeface=" in text:      # 이미 붙어 있으면 건드리지 않는다
        return text, 0
    return LATIN_RE.sub(repl, text), count


AX_BLOCK_RE = re.compile(r"<c:(catAx|valAx|serAx|dateAx)>.*?</c:\1>", re.S)
AXID_RE = re.compile(r'<c:axId val="(\d+)"\s*/>')
SER_RE = re.compile(r"<c:ser>.*?</c:ser>", re.S)
DPT_RE = re.compile(r"<c:dPt>.*?</c:dPt>", re.S)


def fix_chart_axes(text: str) -> tuple:
    """선언되지 않은 <c:axId> 참조를 제거한다 (pptxgenjs 4.x가 세 번째 축 id를 만든다).

    PowerPoint는 존재하지 않는 축을 가리키는 차트를 만나면 파일 전체를 손상으로 보고
    복구를 요구한다. LibreOffice는 조용히 무시하므로 렌더만으로는 드러나지 않는다.
    """
    declared = set()
    for m in AX_BLOCK_RE.finditer(text):
        ids = AXID_RE.findall(m.group(0))
        declared.update(ids)
    if not declared:
        return text, 0

    # 축 블록은 건드리지 않도록 잠시 빼 둔다
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


def fix_ser_order(text: str) -> tuple:
    """<c:ser> 안의 <c:dPt>를 <c:dLbls> 앞으로 옮긴다 (ISO 자식 순서)."""
    moved = 0

    def fix(m):
        nonlocal moved
        ser = m.group(0)
        dlbls = ser.find("<c:dLbls>")
        if dlbls == -1:
            return ser
        pts = [p for p in DPT_RE.finditer(ser) if p.start() > dlbls]
        if not pts:
            return ser
        blocks = [p.group(0) for p in pts]
        for b in blocks:
            ser = ser.replace(b, "", 1)
        moved += len(blocks)
        at = ser.find("<c:dLbls>")
        return ser[:at] + "".join(blocks) + ser[at:]

    return SER_RE.sub(fix, text), moved


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
    ea_total = ax_total = dpt_total = 0
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
            elif item.filename.startswith("ppt/charts/chart") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                text, n = inject_ea(text)
                ea_total += n
                text, a = fix_chart_axes(text)
                ax_total += a
                text, d = fix_ser_order(text)
                dpt_total += d
                if n or a or d:
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
        print(f"센티넬({SENTINEL}) 도형 없음 — 교체할 그라디언트가 없다")
    else:
        detail = ", ".join(f"slide{k}:{v}" for k, v in sorted(per_slide.items()))
        print(f"Hero Gradient {total}개 적용 ({detail})")
    if ea_total:
        print(f"차트 한글 폰트(<a:ea>) {ea_total}곳 주입")
    if ax_total:
        print(f"미선언 축 참조(<c:axId>) {ax_total}개 제거 — PowerPoint 손상 경고 원인")
    if dpt_total:
        print(f"<c:dPt> {dpt_total}개를 <c:dLbls> 앞으로 이동 (ISO 순서)")
    print(f"→ {dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
