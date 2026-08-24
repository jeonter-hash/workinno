#!/usr/bin/env python3
"""check_layout.py — KSA 모노톤 A4 덱 자동 점검.

    python3 check_layout.py deck.pptx [--mode present|report] [--special 1,12] [--strict]

검사 항목
   1  A4 가로 (10.8333" × 7.5") 슬라이드 크기
   2  맑은 고딕 외 폰트 (슬라이드·표·차트 파트 전부)
   3  최소 글자 크기 (발표용 9.5pt / 보고서용 9pt)
   4  슬라이드 밖 이탈
   5  본문 하드 경계 침범 (2.39"–6.85") — **네이티브 표는 <a:tr h> 합산으로 실제 높이 계산**
   6  5존 앵커 고정 (챕터 0.40 / 헤드라인 1.00 / 부제 1.63 / 본문 2.39 / 푸터 7.05)
   7  본문 밀도 — 하단 30% 공백, 본문 점유율
   8  의도하지 않은 검은 윤곽선 — pptxgenjs가 line:{width:0}을 1pt #333333으로 그린다
   9  흰 배경 위 너무 흐린 글자 (#9E9E9E보다 밝은 색)
  10  무채색 팔레트 이탈 (유채색 사용)
  11  로고 위치·비율·로고 뒤 도형
  12  이모지
  13  차트 구조 결함 — 미선언 <c:axId>, <c:dPt>/<c:dLbls> 순서 (PowerPoint 복구 대화상자 원인)
  14  각주 과다 · 결론 밴드 남용 — 잔글씨 3줄 이상, 결론 밴드가 본문 장표 60% 초과
  15  문안 — 이중 피동·번역투·공문 축약형 등 걷어낼 표현

FAIL이 하나라도 있으면 종료 코드 1. 표지·디바이더·클로징은 --special 로 제외한다.
"""
import argparse
import pathlib
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

EMU = 914400.0
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

SLIDE_W, SLIDE_H = 10.8333, 7.5
# 헤드 존 좌표는 모드마다 다르다 (assets/ksa_mono.js MODE 와 같아야 한다)
ZONE = {
    "present": {"챕터": 0.40, "타이틀": 0.94, "메시지": 1.43, "본문": 2.11, "푸터": 7.05},
    "report":  {"챕터": 0.40, "타이틀": 0.94, "메시지": 1.37, "본문": 2.00, "푸터": 7.05},
}
BODY_BOTTOM, CLEAR_BOTTOM = 6.85, 7.05
# 헤드 메시지 종결 — 명사형 개조식. 서술체·의문형은 헤드에 쓰지 않는다
MSG_END = re.compile(r"(함|임|음|됨|짐|필요|시급|불가피|전망)\s*$")
MSG_LINE = {"present": 54, "report": 59}      # 한 줄에 들어가는 한글 글자 수
TITLE_MAX = 30
TOL = 0.02
LOGO_H, LOGO_Y, LOGO_RIGHT, LOGO_MAX_W = 0.24, 0.44, SLIDE_W - 0.5, 1.9
FONT_OK = ("맑은 고딕", "Malgun Gothic")
MIN_PT = {"present": 9.5, "report": 8.5}   # 보고서 dense 모드까지 허용, 8.5pt가 절대 하한
ASSETS = pathlib.Path(__file__).resolve().parent.parent / "assets"

# 무채색 팔레트 — 이 밖의 색은 유채색이거나 정의되지 않은 회색
PALETTE = {"111111", "3A3A3A", "6B6B6B", "9E9E9E", "C7C7C7", "E4E4E4",
           "F4F4F4", "FAFAFA", "FFFFFF", "000000", "333333"}
ACCENT = {"1F3864", "4A6491", "C4303C", "F2E4E6", "E8A0A6"}   # --palette accent 에서만 허용
TEXT_FLOOR = 0x9E          # 흰 배경 위 글자는 이보다 밝으면 안 된다
# 겹침을 볼 도형 — 체브론·화살표는 설명글을 덮기 쉬우므로 포함한다
OVERLAP_PRST = ("rect", "roundRect", "chevron", "homePlate", "rightArrow", "downArrow", "pentagon")
GHOST_LINE = re.compile(r'<a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/>')
EMOJI = re.compile("[\U0001F300-\U0001FAFF\U00002600-\U000027BF\uFE0F]")


class Box:
    def __init__(self, kind, x, y, w, h, el=None, text="", filled=False, z=0):
        self.kind, self.x, self.y, self.w, self.h, self.el = kind, x, y, w, h, el
        self.text, self.filled, self.z = text, filled, z

    @property
    def bottom(self):
        return self.y + self.h

    def __repr__(self):
        return f"{self.kind}(x={self.x:.2f} y={self.y:.2f} w={self.w:.2f} h={self.h:.2f})"


def table_height(frame_xml: str):
    """네이티브 표의 실제 높이.

    pptxgenjs는 graphicFrame의 ext cy를 1.0"로 고정 기재하므로 그 값은 쓸 수 없다.
    행 높이(<a:tr h>)를 합산해야 표가 실제로 차지하는 높이가 나온다.
    """
    rows = [int(v) for v in re.findall(r'<a:tr h="(\d+)"', frame_xml)]
    return sum(rows) / EMU if rows else None


def boxes(raw: str, tree):
    """상위 도형의 좌표 목록. 네이티브 표는 실제 높이로 보정한다."""
    out = []
    spTree = tree.find(f"{P}cSld/{P}spTree")
    if spTree is None:
        return out
    frames = re.findall(r"<p:graphicFrame>.*?</p:graphicFrame>", raw, re.S)
    fi = 0
    for el in list(spTree):
        tag = el.tag.split("}")[1]
        if tag not in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
            continue
        xfrm = el.find(f".//{A}xfrm") or el.find(f"{P}xfrm")
        if xfrm is None:
            continue
        off, ext = xfrm.find(f"{A}off"), xfrm.find(f"{A}ext")
        if off is None or ext is None:
            continue
        x, y = int(off.get("x", 0)) / EMU, int(off.get("y", 0)) / EMU
        w, h = int(ext.get("cx", 0)) / EMU, int(ext.get("cy", 0)) / EMU
        if tag == "graphicFrame":
            frame = frames[fi] if fi < len(frames) else ""
            fi += 1
            real = table_height(frame)
            if real:
                h = real
                tag = "표"
        text = "".join((t.text or "") for t in el.iter(f"{A}t")).strip()
        fill = el.find(f"{P}spPr/{A}solidFill/{A}srgbClr")
        filled = fill is not None and fill.get("val", "").upper() != "FFFFFF"
        geom = el.find(f".//{A}prstGeom")
        prst = geom.get("prst", "") if geom is not None else ""
        b = Box(tag, x, y, w, h, el, text, filled, len(out)); b.prst = prst
        b.fillc = fill.get("val", "").upper() if fill is not None else ""
        out.append(b)
    return out


def overlaps(a, b):
    """두 상자가 부분적으로 겹치는가 (한쪽이 다른 쪽을 온전히 담고 있으면 겹침이 아니다)."""
    ix = min(a.x + a.w, b.x + b.w) - max(a.x, b.x)
    iy = min(a.y + a.h, b.y + b.h) - max(a.y, b.y)
    if ix <= 0.02 or iy <= 0.02:
        return 0.0
    contains = lambda p, q: (p.x <= q.x + 0.01 and p.y <= q.y + 0.01
                             and p.x + p.w >= q.x + q.w - 0.01 and p.y + p.h >= q.y + q.h - 0.01)
    if contains(a, b) or contains(b, a):
        return 0.0
    return ix * iy


def light(rgb: str) -> float:
    return int(rgb[0:2], 16) * 0.299 + int(rgb[2:4], 16) * 0.587 + int(rgb[4:6], 16) * 0.114


def is_gray(rgb: str) -> bool:
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return max(r, g, b) - min(r, g, b) <= 8


def image_ratio(p):
    try:
        b = open(p, "rb").read()
    except OSError:
        return None
    if b[1:4] == b"PNG":
        w, h = int.from_bytes(b[16:20], "big"), int.from_bytes(b[20:24], "big")
        return w / h if h else None
    if b[:2] == b"\xff\xd8":
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


# ── 잔글씨(각주·출처)와 결론 밴드 판별 ────────────────────────
FOOT_MAX_PT = 950          # 9.5pt 이하 + 회색이면 각주·출처로 본다
FOOT_GRAY = {"6B6B6B", "9E9E9E"}
FOOT_RE = re.compile(r"^\s*(?:주\s*\d|주\)|출처|자료|단위|※|\*)")


def small_gray_texts(tree):
    """표 아래 잔글씨(각주·출처) 줄 목록."""
    out = []
    for sp in tree.iter(f"{P}sp"):
        szs = [int(r.get("sz")) for r in sp.iter(f"{A}rPr") if r.get("sz")]
        cols = {c.get("val", "").upper() for c in sp.iter(f"{A}srgbClr")}
        t = "".join((x.text or "") for x in sp.iter(f"{A}t")).strip()
        if t and szs and max(szs) <= FOOT_MAX_PT and cols and cols <= FOOT_GRAY and FOOT_RE.match(t):
            out.append(t)
    return out


def has_callout(bs):
    """결론 밴드(가로로 긴 채움 띠 + 글자)가 있는가."""
    for b in bs:
        if (b.kind == "sp" and b.text and getattr(b, "prst", "") in ("rect", "roundRect")
                and getattr(b, "fillc", "") in CALLOUT_FILL and b.w > 3.0 and b.h <= 0.95):
            return True
    return False


CALLOUT_FILL = {"F4F4F4", "111111", "1F3864", "C4303C"}

# ── 문안 규칙 — 컨설턴트 문투에서 걷어낼 표현 ──────────────────
# 본문에만 적용한다. 헤드는 한자어 명사구·'~ 시'·'~ 필요함'을 표준으로 쓴다(references/headline.md).
BAD_PHRASE = [
    (re.compile(r"되어지|하여지|불려지|보여지"), "이중 피동 — '~된다'로 쓸 것"),
    (re.compile(r"것으로 (?:판단|사료|보여)됨"), "상투구 — 판단을 그대로 쓸 것"),
    (re.compile(r"대폭|획기적|전방위적|비약적"), "수치 없는 평가어 — 숫자로 대체할 것"),
]


def head_texts(bs, z):
    """헤드 존(타이틀·메시지)의 글상자를 (타이틀, 메시지 목록)으로 나눈다."""
    title, msg = [], []
    for b in bs:
        if not b.text:
            continue
        if abs(b.y - z["타이틀"]) <= TOL:
            title.append(b)
        elif abs(b.y - z["메시지"]) <= TOL:
            msg.append(b)
    return title, msg


def logo_asset():
    for name in ("ksa_logo.png", "ksa_logo.jpg"):
        p = ASSETS / name
        if p.exists():
            return p
    return None


def check(path, mode, special, strict, palette='mono'):
    fails, warns = [], []
    body_slides, callout_slides = [], []
    zone = ZONE[mode]
    body_top = zone["본문"]
    msg_bottom = zone["본문"] - 0.10        # 메시지 글자 하단과 본문 사이의 완충 구간
    min_pt = MIN_PT[mode]
    allowed_chroma = ACCENT if palette == 'accent' else set()
    asset = logo_asset()
    src_ratio = image_ratio(asset) if asset else None

    with zipfile.ZipFile(path) as z:
        pres = ET.fromstring(z.read("ppt/presentation.xml"))
        sz = pres.find(f"{P}sldSz")
        w_in, h_in = int(sz.get("cx")) / EMU, int(sz.get("cy")) / EMU
        if abs(w_in - SLIDE_W) > 0.02 or abs(h_in - SLIDE_H) > 0.02:
            fails.append(f'[deck] 슬라이드 크기 {w_in:.3f}"×{h_in:.3f}" — A4 가로 10.833×7.5 이어야 한다')

        names = sorted((n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                       key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        for name in names:
            n = int(re.search(r"slide(\d+)", name).group(1))
            raw = z.read(name).decode("utf-8")
            tree = ET.fromstring(raw)
            bs = boxes(raw, tree)
            tag = f"[slide{n}]"
            is_special = n in special or any(b.w > SLIDE_W - 0.2 and b.h > SLIDE_H - 0.4 for b in bs)

            # 2·3 폰트
            for f in tree.iter():
                if f.tag in (f"{A}latin", f"{A}ea", f"{A}cs"):
                    face = f.get("typeface", "")
                    if face and not face.startswith(FONT_OK) and not face.startswith("+"):
                        fails.append(f"{tag} 맑은 고딕 외 폰트: {face}")
            for rpr in tree.iter(f"{A}rPr"):
                v = rpr.get("sz")
                if v and int(v) / 100 < min_pt - 0.01:
                    fails.append(f"{tag} 글자 {int(v)/100:.1f}pt — {mode} 최소 {min_pt}pt")

            # 4 슬라이드 밖
            for b in bs:
                if b.x < -0.01 or b.y < -0.01 or b.x + b.w > SLIDE_W + 0.01 or b.bottom > SLIDE_H + 0.01:
                    fails.append(f'{tag} 슬라이드 밖으로 나감: {b} → 우측끝 {b.x + b.w:.3f}" 하단 {b.bottom:.3f}"')

            # 5 본문 하드 경계
            if not is_special:
                for b in bs:
                    if b.y >= body_top - 0.4 and b.bottom > BODY_BOTTOM + TOL and b.y < CLEAR_BOTTOM - TOL:
                        extra = " (네이티브 표 실제 높이 기준)" if b.kind == "표" else ""
                        fails.append(f'{tag} 본문 하단 6.85" 침범: {b}{extra}')
                    elif BODY_BOTTOM + TOL < b.y < CLEAR_BOTTOM - TOL:
                        fails.append(f'{tag} 클리어런스 버퍼(6.85–7.05") 침범: {b}')
                    if msg_bottom + TOL < b.y < body_top - TOL:
                        fails.append(f'{tag} 헤드와 본문 사이({msg_bottom:.2f}–{body_top:.2f}") 침범: {b}')

            # 6 존 앵커
            if not is_special:
                ys = [b.y for b in bs]
                missing = [k for k, v in zone.items() if not any(abs(y - v) <= TOL for y in ys)]
                # 챕터·메시지는 목차·간지에서 비울 수 있다 — 경고로만 본다
                hard = [k for k in missing if k not in ("챕터", "메시지")]
                if hard:
                    fails.append(f"{tag} 존 앵커 없음: {', '.join(hard)} — 5존 좌표가 흔들렸다")
                for k in ("챕터", "메시지"):
                    if k in missing:
                        warns.append(f"{tag} {k} 없음 (목차·간지면 정상)")

            # 6b 헤드 문법 — 타이틀은 명사구 1줄, 메시지는 명사형 개조식 2줄까지
            if not is_special:
                titles, msgs = head_texts(bs, zone)
                for b in titles:
                    t = b.text
                    if len(t) > TITLE_MAX:
                        fails.append(f"{tag} 헤드 타이틀 {len(t)}자 — 상한 {TITLE_MAX}자: '{t[:34]}'")
                    if t.endswith(("니다", "습니다", "함니다")) or t.endswith("?"):
                        fails.append(f"{tag} 헤드 타이틀이 문장이다 — 명사구로 쓸 것: '{t[:34]}'")
                    if ":" in t or "：" in t:
                        warns.append(f"{tag} 헤드 타이틀에 대분류가 붙어 있다 — 챕터에 있으므로 뺄 것: '{t[:34]}'")
                for b in msgs:
                    m = b.text
                    if not MSG_END.search(m):
                        fails.append(f"{tag} 헤드 메시지가 명사형으로 끝나지 않는다 — '{m[-14:]}'")
                    est = -(-len(m) // MSG_LINE[mode])
                    if est > 2:
                        fails.append(f"{tag} 헤드 메시지 {len(m)}자 — {mode} 2줄({MSG_LINE[mode]*2}자)을 넘는다")
                    elif len(m) < 40:
                        warns.append(f"{tag} 헤드 메시지가 {len(m)}자로 짧다 — 근거가 빠졌는지 볼 것: '{m}'")

            # 7 밀도
            if not is_special:
                area = (SLIDE_W - 1.0) * (BODY_BOTTOM - body_top)
                covered, bottom_band = 0.0, False
                for b in bs:
                    top, bot = max(b.y, body_top), min(b.bottom, BODY_BOTTOM)
                    if bot > top:
                        covered += (bot - top) * min(b.w, SLIDE_W - 1.0)
                        if bot > BODY_BOTTOM - 0.30 * (BODY_BOTTOM - body_top):
                            bottom_band = True
                if not bottom_band:
                    fails.append(f'{tag} 본문 하단 30%가 비었다 — 근거·콜아웃·표로 채울 것')
                # 본문 안의 빈 띠 — 요소 사이가 크게 비면 도형·표를 늘려 채운다
                spans = sorted((max(b.y, body_top), min(b.bottom, BODY_BOTTOM))
                               for b in bs if b.bottom > body_top and b.y < BODY_BOTTOM)
                cur, gaps = body_top, []
                for a0, a1 in spans:
                    if a0 - cur > 0.001:
                        gaps.append((cur, a0))
                    cur = max(cur, a1)
                if BODY_BOTTOM - cur > 0.001:
                    gaps.append((cur, BODY_BOTTOM))
                if gaps:
                    g0, g1 = max(gaps, key=lambda g: g[1] - g[0])
                    if g1 - g0 > 0.45:
                        warns.append(f'{tag} 본문에 빈 띠 {g1-g0:.2f}" ({g0:.2f}–{g1:.2f}") — 도형·표를 늘려 채울 것')
                elif covered / area < 0.45:
                    warns.append(f"{tag} 본문 점유율 {covered/area:.0%} — 밀도 부족")

            # 7b 요소 겹침 — 나중에 그린 채움 도형이 앞선 글자를 덮는 경우
            if not is_special:
                for i, a in enumerate(bs):
                    if not a.text or a.y < body_top - 0.1:
                        continue
                    for b in bs[i + 1:]:
                        # 체브론·화살표처럼 서로 물리도록 설계된 도형은 제외 (사각형 계열만 본다)
                        if not b.filled or b.text or getattr(b, "prst", "") not in OVERLAP_PRST:
                            continue
                        ratio = overlaps(a, b) / max(a.w * a.h, 0.001)
                        if ratio > 0.30:
                            fails.append(f"{tag} 글자가 도형에 가림({ratio:.0%}): '{a.text[:18]}' {a} ↔ {b}")
                # 네이티브 표는 글상자가 아니라 graphicFrame이라 위 규칙에 걸리지 않는다 — 따로 본다
                for t in (x for x in bs if x.kind == "표"):
                    for b in bs:
                        if b is t or not b.filled or getattr(b, "prst", "") not in OVERLAP_PRST:
                            continue
                        if overlaps(t, b) > 0.05:
                            fails.append(f"{tag} 표가 도형과 겹침: {t} ↔ {b} — 표 반환값 bottom으로 다음 y를 잡을 것")

            # 8 의도하지 않은 검은 윤곽선
            ghosts = len(GHOST_LINE.findall(raw))
            if ghosts:
                fails.append(f"{tag} 검은 윤곽선 {ghosts}개 — line:{{width:0}}은 1pt #333333으로 그려진다. "
                             f"line:{{type:'none'}}을 쓸 것")

            # 9·10 글자 색
            for sp in tree.iter(f"{A}rPr"):
                fill = sp.find(f"{A}solidFill/{A}srgbClr")
                if fill is None:
                    continue
                c = fill.get("val", "").upper()
                if not c:
                    continue
                if not is_gray(c):
                    if c not in allowed_chroma:
                        warns.append(f"{tag} 팔레트 밖 글자색 #{c}")
                elif light(c) > TEXT_FLOOR * 0.98 and c != "FFFFFF":
                    warns.append(f"{tag} 흐린 글자색 #{c} — 흰 배경이면 #6B6B6B 이상으로 (어두운 배경 위면 무시)")
            for c in set(re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', raw)):
                cu = c.upper()
                if not is_gray(cu) and cu not in allowed_chroma:
                    warns.append(f"{tag} 팔레트 밖 색 #{cu}")

            # 11 로고
            pics = [b for b in bs if b.kind == "pic"]
            logo = [b for b in pics if abs(b.y - LOGO_Y) <= 0.10 and b.x > SLIDE_W / 2 and 0.10 <= b.h <= LOGO_H + 0.06]
            # 이 체계는 로고를 쓰지 않는다(표지는 사진). logoAt()으로 넣었을 때만 위치·비율을 본다.
            for b in logo:
                if abs((b.x + b.w) - LOGO_RIGHT) > 0.03:
                    fails.append(f'{tag} 로고 우측 끝 {b.x + b.w:.3f}" — 오른쪽 여백 0.5"(={LOGO_RIGHT:.3f}")로 맞출 것')
                if b.w > LOGO_MAX_W + TOL:
                    fails.append(f'{tag} 로고 폭 {b.w:.3f}" — 상한 {LOGO_MAX_W}"')
                if src_ratio and abs(b.w / b.h - src_ratio) > 0.06:
                    fails.append(f'{tag} 로고 종횡비 변형: {b.w:.3f}×{b.h:.3f}" (원본 {src_ratio:.2f})')
                for o in bs:
                    if o is b or o.kind not in ("sp", "cxnSp"):
                        continue
                    if o.x < b.x + b.w and o.x + o.w > b.x and o.y < b.y + b.h and o.bottom > b.y:
                        fails.append(f"{tag} 로고 뒤/위 도형 — 배경 박스·밑줄·프레임은 디펙트: {o}")

            # 12 이모지
            for t in tree.iter(f"{A}t"):
                if t.text and EMOJI.search(t.text):
                    fails.append(f"{tag} 이모지 사용: {t.text[:24]}")

            # 14 각주 과다 — 단위·출처는 한 줄로 합치고, 각주는 오독 위험이 있을 때만
            if not is_special:
                foot = small_gray_texts(tree)
                if len(foot) >= 3:
                    warns.append(f"{tag} 표 아래 잔글씨 {len(foot)}줄 — 단위·출처는 한 줄로 합치고, "
                                 f"각주는 수치가 오독될 때만 남길 것: {' / '.join(f[:20] for f in foot[:3])}")
                if has_callout(bs):
                    callout_slides.append(n)
                body_slides.append(n)

            # 15 문안 — 걷어낼 표현
            # 문단 단위로 이어 붙인다 — 런 경계에 공백을 넣으면 '정합 시간'이 '~ 시'로 오탐된다.
            # 헤드(타이틀·메시지)는 한자어 명사구와 '~ 시'를 표준으로 쓰므로 검사에서 뺀다.
            head_str = {b.text for b in bs if b.text and
                        (abs(b.y - zone["타이틀"]) <= TOL or abs(b.y - zone["메시지"]) <= TOL)}
            paras = [x for x in ("".join((t.text or "") for t in para.iter(f"{A}t"))
                                 for para in tree.iter(f"{A}p")) if x not in head_str]
            for pat, why in BAD_PHRASE:
                hit = next((pat.search(x) for x in paras if pat.search(x)), None)
                if hit:
                    warns.append(f"{tag} 문안: '{hit.group(0).strip()}' — {why}")

        # 13 차트 파트
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
                    fails.append(f"[{short}] <c:dPt>가 <c:dLbls> 뒤에 있다 (ISO 순서 위반) — postprocess.py 를 실행할 것")
                    break

    # 14b 결론 밴드 남용 — 장(章)의 마지막 장표에만 다는 것이 원칙
    if len(body_slides) >= 4 and len(callout_slides) > 0.6 * len(body_slides):
        warns.append(f"[deck] 결론 밴드가 본문 {len(body_slides)}장 중 {len(callout_slides)}장에 붙어 있다 — "
                     f"장(章)의 마지막 장표에만 달 것 (현재: {', '.join(str(x) for x in callout_slides)})")

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
    ap.add_argument("--mode", choices=("present", "report"), default="report")
    ap.add_argument("--special", default="", help="검사 제외 장표 번호(표지·디바이더·클로징), 쉼표 구분")
    ap.add_argument("--palette", choices=("mono", "accent"), default="mono",
                    help="accent면 KSA 강조 2색(네이비·레드)을 허용한다")
    ap.add_argument("--strict", action="store_true", help="WARN도 실패로 취급")
    a = ap.parse_args()
    special = {int(x) for x in a.special.split(",") if x.strip()}
    return check(a.pptx, a.mode, special, a.strict, a.palette)


if __name__ == "__main__":
    raise SystemExit(main())
