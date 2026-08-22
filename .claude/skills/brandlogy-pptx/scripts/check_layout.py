#!/usr/bin/env python3
"""check_layout.py — 디자인 시스템 자동 점검 (Iteration Checklist 기계화).

    python3 check_layout.py deck.pptx [--special 1,5] [--strict]

검사 항목
  1  A4 가로 (10.8333" × 7.5") 슬라이드 크기
  2  맑은 고딕 외 폰트 사용
  3  본문 하드 경계 침범 (2.39"–6.85"), 클리어런스 버퍼(6.85"–7.05") 침범, 슬라이드 밖 이탈
  4  5존 앵커 고정 (챕터 0.40 / 헤드라인 1.00 / 부제 1.63 / 본문 2.39 / 푸터 7.05)
  5  본문 밀도 — 하단 30% 밴드 공백, 본문 박스 점유율
  6  Hero Gradient 개수(장표 1 / 덱 3), 센티넬 잔존
  7  Brand Glow 장표당 1개
  8  최소 폰트 9pt
  9  로고 존재·위치·비율·로고 뒤 도형(디펙트)
 10  이모지
 11  팔레트 밖 색상 (경고)
 12  차트/도식 없는 본문 장표 (경고 — Visualization-First)
 13  차트 구조 결함 — 선언되지 않은 <c:axId>, <c:dPt>/<c:dLbls> 순서 (PowerPoint가 파일을
     손상으로 보고 복구를 요구하는 원인). postprocess.py를 돌리면 고쳐진다

FAIL이 하나라도 있으면 종료 코드 1. 표지·섹션 디바이더처럼 프레임을 의도적으로
깨는 장표는 --special 로 번호를 넘기면 존 앵커·밀도 검사를 건너뛴다(전면 도형이
있으면 자동 인식도 한다).
"""
import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

EMU = 914400.0
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

SLIDE_W, SLIDE_H = 10.8333, 7.5
ANCHORS = {"챕터": 0.40, "헤드라인": 1.00, "부제": 1.63, "본문": 2.39, "푸터": 7.05}
BODY_TOP, BODY_BOTTOM, CLEAR_BOTTOM = 2.39, 6.85, 7.05
TOL = 0.02
LOGO_H, LOGO_Y, LOGO_RIGHT, LOGO_MAX_W = 0.24, 0.44, SLIDE_W - 0.5, 1.9
FONT_OK = ("맑은 고딕", "Malgun Gothic")
# 저장소에 로고 원본이 있으면 그 비율과 대조한다
ASSETS = __import__("pathlib").Path(__file__).resolve().parent.parent / "assets"

PALETTE = {
    "1456F0", "3B82F6", "60A5FA", "BFDBFE", "2563EB", "1D4ED8", "17437D", "3DAEFF",
    "EA5EC1", "222222", "18181B", "181E25", "45515E", "8E8E93", "5F5F5F", "333333",
    "FFFFFF", "F0F0F0", "F2F3F5", "E5E7EB", "FAFAFA", "E8FFEA", "16A34A", "2C1E74",
    "000000", "242424",
}
EMOJI = re.compile("[\U0001F300-\U0001FAFF☀-➿️]")


class Box:
    def __init__(self, kind, x, y, w, h, el):
        self.kind, self.x, self.y, self.w, self.h, self.el = kind, x, y, w, h, el

    @property
    def bottom(self):
        return self.y + self.h

    def __repr__(self):
        return f"{self.kind}(x={self.x:.2f} y={self.y:.2f} w={self.w:.2f} h={self.h:.2f})"


def boxes(tree):
    out = []
    spTree = tree.find(f"{P}cSld/{P}spTree")
    if spTree is None:
        return out
    for el in list(spTree):
        tag = el.tag.split("}")[1]
        if tag not in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
            continue
        xfrm = el.find(f".//{A}xfrm")
        if xfrm is None:
            xfrm = el.find(f"{P}xfrm")
        if xfrm is None:
            continue
        off, ext = xfrm.find(f"{A}off"), xfrm.find(f"{A}ext")
        if off is None or ext is None:
            continue
        out.append(Box(tag,
                       int(off.get("x", 0)) / EMU, int(off.get("y", 0)) / EMU,
                       int(ext.get("cx", 0)) / EMU, int(ext.get("cy", 0)) / EMU, el))
    return out


def texts(tree):
    return [(t.text or "") for t in tree.iter(f"{A}t")]


def image_ratio(p):
    """PNG·JPEG 헤더에서 가로/세로 비율 (읽을 수 없으면 None)"""
    try:
        b = open(p, "rb").read()
    except OSError:
        return None
    if b[1:4] == b"PNG":
        w, h = int.from_bytes(b[16:20], "big"), int.from_bytes(b[20:24], "big")
        return w / h if h else None
    if b[:2] == b"\xff\xd8":                       # JPEG: SOF 마커
        i = 2
        while i < len(b) - 9:
            if b[i] != 0xFF:
                i += 1
                continue
            m = b[i + 1]
            if 0xC0 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
                h = int.from_bytes(b[i + 5:i + 7], "big")
                w = int.from_bytes(b[i + 7:i + 9], "big")
                return w / h if h else None
            if m in (0xD8, 0xD9) or 0xD0 <= m <= 0xD7:
                i += 2
                continue
            i += 2 + int.from_bytes(b[i + 2:i + 4], "big")
    return None


def logo_asset():
    for name in ("ksa_logo.png", "ksa_logo.jpg"):
        p = ASSETS / name
        if p.exists():
            return p
    return None


def check(path, special, strict):
    fails, warns = [], []
    asset = logo_asset()
    src_ratio = image_ratio(asset) if asset else None
    with zipfile.ZipFile(path) as z:
        pres = ET.fromstring(z.read("ppt/presentation.xml"))
        sz = pres.find(f"{P}sldSz")
        w_in, h_in = int(sz.get("cx")) / EMU, int(sz.get("cy")) / EMU
        if abs(w_in - SLIDE_W) > 0.02 or abs(h_in - SLIDE_H) > 0.02:
            fails.append(f"[deck] 슬라이드 크기 {w_in:.3f}\"×{h_in:.3f}\" — A4 가로 10.833×7.5 이어야 한다")

        names = sorted((n for n in z.namelist()
                        if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                       key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        grad_total = 0
        for name in names:
            n = int(re.search(r"slide(\d+)", name).group(1))
            raw = z.read(name).decode("utf-8")
            tree = ET.fromstring(raw)
            bs = boxes(tree)
            tag = f"[slide{n}]"
            is_special = n in special or any(b.w > SLIDE_W - 0.2 and b.h > SLIDE_H - 0.4 for b in bs)

            # 2 · 8 폰트
            for f in tree.iter():
                if f.tag in (f"{A}latin", f"{A}ea", f"{A}cs"):
                    face = f.get("typeface", "")
                    if face and not face.startswith(FONT_OK) and not face.startswith("+"):
                        fails.append(f"{tag} 맑은 고딕 외 폰트: {face}")
            for rpr in tree.iter(f"{A}rPr"):
                sz_ = rpr.get("sz")
                if sz_ and int(sz_) < 900:
                    fails.append(f"{tag} 폰트 {int(sz_) / 100:.1f}pt — 최소 9pt")

            # 3 슬라이드 밖 이탈 (판형을 바꿀 때 좌표를 놓치면 여기서 잡힌다)
            for b in bs:
                if b.x < -0.01 or b.y < -0.01 or b.x + b.w > SLIDE_W + 0.01 or b.bottom > SLIDE_H + 0.01:
                    fails.append(f"{tag} 슬라이드 밖으로 나감({SLIDE_W}×{SLIDE_H}\"): {b} → 우측끝 {b.x + b.w:.3f}\" 하단 {b.bottom:.3f}\"")

            # 3b 본문 하드 경계
            for b in bs:
                if is_special:
                    break
                if b.y >= BODY_TOP - 0.4 and b.bottom > BODY_BOTTOM + TOL and b.y < CLEAR_BOTTOM - TOL:
                    fails.append(f"{tag} 본문 하단 6.85\" 침범: {b}")
                elif BODY_BOTTOM + TOL < b.y < CLEAR_BOTTOM - TOL:
                    fails.append(f"{tag} 클리어런스 버퍼(6.85–7.05\") 침범: {b}")
                if 2.03 + TOL < b.y < BODY_TOP - TOL:
                    fails.append(f"{tag} 부제 존과 본문 사이(2.03–2.39\") 침범: {b}")

            # 4 존 앵커
            if not is_special:
                ys = [b.y for b in bs]
                missing = [k for k, v in ANCHORS.items() if not any(abs(y - v) <= TOL for y in ys)]
                if missing:
                    fails.append(f"{tag} 존 앵커 없음: {', '.join(missing)} — 5존 좌표가 흔들렸다")

            # 5 밀도
            if not is_special:
                body_area = (SLIDE_W - 1.0) * (BODY_BOTTOM - BODY_TOP)
                covered = 0.0
                bottom_band = False
                for b in bs:
                    top, bot = max(b.y, BODY_TOP), min(b.bottom, BODY_BOTTOM)
                    if bot > top:
                        covered += (bot - top) * min(b.w, SLIDE_W - 1.0)
                        if bot > BODY_BOTTOM - 0.30 * (BODY_BOTTOM - BODY_TOP):
                            bottom_band = True
                ratio = covered / body_area
                if not bottom_band:
                    fails.append(f"{tag} 본문 하단 30%(5.51–6.85\")가 비었다 — 근거 카드·콜아웃·차트로 채울 것")
                elif ratio < 0.45:
                    warns.append(f"{tag} 본문 점유율 {ratio:.0%} — 밀도 부족. 근거·차트를 보강할 것")

            # 6 그라디언트
            g = raw.count("<a:gradFill")
            grad_total += g
            if g > 1:
                fails.append(f"{tag} Hero Gradient {g}개 — 장표당 1개")
            if "0A0B0C" in raw:
                fails.append(f"{tag} 그라디언트 센티넬(0A0B0C) 잔존 — postprocess.py 를 실행하지 않았다")

            # 7 Brand Glow
            glow = len(re.findall(r'<a:outerShdw[^>]*>\s*<a:srgbClr val="2C1E74"', raw))
            if glow > 1:
                fails.append(f"{tag} Brand Glow {glow}개 — 장표당 1개")

            # 9 로고
            pics = [b for b in bs if b.kind == "pic"]
            logo = [b for b in pics
                    if abs(b.y - LOGO_Y) <= 0.10 and b.x > SLIDE_W / 2 and 0.10 <= b.h <= LOGO_H + 0.06]
            if not logo and not is_special:
                warns.append(f"{tag} 우상단 로고 없음 — 한국표준협회(KSA) 누끼 PNG를 넣을 것")
            for b in logo:
                if abs((b.x + b.w) - LOGO_RIGHT) > 0.03:
                    fails.append(f"{tag} 로고 우측 끝 {b.x + b.w:.3f}\" — 오른쪽 여백 0.5\"(={LOGO_RIGHT:.3f}\")로 맞출 것")
                if b.w > LOGO_MAX_W + TOL:
                    fails.append(f"{tag} 로고 폭 {b.w:.3f}\" — 상한 {LOGO_MAX_W}\"")
                if src_ratio and abs(b.w / b.h - src_ratio) > 0.06:
                    fails.append(f"{tag} 로고 종횡비 변형: {b.w:.3f}×{b.h:.3f}\" (원본 비율 {src_ratio:.2f})")
                for o in bs:
                    if o is b or o.kind not in ("sp", "cxnSp"):
                        continue
                    if (o.x < b.x + b.w and o.x + o.w > b.x
                            and o.y < b.y + b.h and o.bottom > b.y):
                        fails.append(f"{tag} 로고 뒤/위 도형 발견 — 배경 박스·밑줄·프레임은 디펙트: {o}")

            # 10 이모지 · 11 색상 · 12 시각화
            for t in texts(tree):
                if EMOJI.search(t):
                    fails.append(f"{tag} 이모지 사용: {t[:30]}")
            for c in set(re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', raw)):
                if c.upper() not in PALETTE:
                    warns.append(f"{tag} 팔레트 밖 색상 #{c.upper()}")
            if not is_special and not any(b.kind == "graphicFrame" for b in bs):
                warns.append(f"{tag} 차트·표 없음 — 데이터·비교·프로세스를 다루면 시각화할 것(Visualization-First)")

        # 차트 파트 — 폰트와 구조 (슬라이드 XML에는 안 나온다)
        for name in (n for n in z.namelist() if re.fullmatch(r"ppt/charts/chart\d+\.xml", n)):
            raw = z.read(name).decode("utf-8")
            short = name.split("/")[-1]
            for face in set(re.findall(r'<a:(?:latin|ea|cs) typeface="([^"]+)"', raw)):
                if not face.startswith(FONT_OK) and not face.startswith("+"):
                    fails.append(f"[{short}] 맑은 고딕 외 폰트: {face}")
            declared = set()
            for m in re.finditer(r"<c:(catAx|valAx|serAx|dateAx)>.*?</c:\1>", raw, re.S):
                declared.update(re.findall(r'<c:axId val="(\d+)"\s*/>', m.group(0)))
            body = re.sub(r"<c:(catAx|valAx|serAx|dateAx)>.*?</c:\1>", "", raw, flags=re.S)
            ghost = {i for i in re.findall(r'<c:axId val="(\d+)"\s*/>', body)} - declared
            if ghost:
                fails.append(f"[{short}] 선언되지 않은 축 참조 <c:axId val={','.join(sorted(ghost))}> — "
                             f"PowerPoint가 파일 손상으로 보고 복구를 요구한다. postprocess.py 를 실행할 것")
            for m in re.finditer(r"<c:ser>.*?</c:ser>", raw, re.S):
                ser = m.group(0)
                d = ser.find("<c:dLbls>")
                if d != -1 and any(p.start() > d for p in re.finditer(r"<c:dPt>", ser)):
                    fails.append(f"[{short}] <c:dPt>가 <c:dLbls> 뒤에 있다 (ISO 자식 순서 위반) — postprocess.py 를 실행할 것")
                    break

        if grad_total > 3:
            fails.append(f"[deck] Hero Gradient {grad_total}개 — 덱 전체 3개 상한 초과")

    dedup = lambda xs: list(dict.fromkeys(xs))
    fails, warns = dedup(fails), dedup(warns)
    for f in fails:
        print("FAIL " + f)
    for w in warns:
        print("WARN " + w)
    print(f"\n{len(fails)} fail / {len(warns)} warn")
    if fails:
        return 1
    return 1 if (strict and warns) else 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--special", default="", help="검사에서 제외할 장표 번호(표지·디바이더·클로징), 쉼표 구분")
    ap.add_argument("--strict", action="store_true", help="WARN도 실패로 취급")
    a = ap.parse_args()
    special = {int(x) for x in a.special.split(",") if x.strip()}
    return check(a.pptx, special, a.strict)


if __name__ == "__main__":
    raise SystemExit(main())
