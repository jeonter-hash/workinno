#!/usr/bin/env python3
"""measure_head.py — 렌더된 장표에서 헤드 글자 띠의 실제 경계를 잰다.

    soffice --headless --convert-to pdf deck.pptx
    pdftoppm -r 105 -png deck.pdf p
    python3 measure_head.py p-1.png

맑은 고딕은 상자 위에서 글자까지 여백이 생기고 줄 간격도 pt×행간÷72보다 넓다.
그래서 헤드 좌표는 계산이 아니라 이 스크립트의 실측값으로 잡는다.
타이틀↔메시지 간격과 메시지↔본문 간격이 같은지 확인하는 데 쓴다.
"""
import sys
from PIL import Image


def bands(path, top=0.75, bottom=2.45, thresh=140):
    im = Image.open(path).convert("L")
    W, H = im.size
    inch = 7.5 / H
    px = im.load()
    x0, x1 = int(W * 0.04), int(W * 0.92)
    rows = []
    for y in range(int(top / inch), int(bottom / inch)):
        dark = sum(1 for x in range(x0, x1, 2) if px[x, y] < thresh)
        rows.append((y, dark))
    out, cur = [], None
    for y, d in rows:
        if d >= 2:
            cur = [y, y] if cur is None else [cur[0], y]
        elif cur:
            if cur[1] - cur[0] >= 2:
                out.append(tuple(cur))
            cur = None
    if cur and cur[1] - cur[0] >= 2:
        out.append(tuple(cur))
    merged = []
    for b in out:                      # 3px 이내로 붙은 띠는 한 줄로 본다
        if merged and b[0] - merged[-1][1] <= 3:
            merged[-1] = (merged[-1][0], b[1])
        else:
            merged.append(b)
    return [(a * inch, b * inch) for a, b in merged], inch


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    for path in sys.argv[1:]:
        bs, inch = bands(path)
        print(f"{path}  (1px = {inch:.4f}\")")
        prev = None
        for a, b in bs:
            gap = f"   간격 {a - prev:.3f}\"" if prev is not None else ""
            print(f'  {a:6.3f}" ~ {b:6.3f}"  (높이 {b - a:.3f}"){gap}')
            prev = b
        if len(bs) >= 3:
            g = [bs[i + 1][0] - bs[i][1] for i in range(len(bs) - 1)]
            print(f"  → 간격 편차 {max(g) - min(g):.3f}\" (1px 이내면 같은 것으로 본다)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
