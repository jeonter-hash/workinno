#!/usr/bin/env python3
"""audit_skills.py — present-ppt·report-ppt의 문서가 코드와 어긋났는지 본다.

    python3 tools/audit_skills.py

이 저장소에서 스킬을 여러 번 고치는 동안, 코드를 바꾸고 문서를 안 고쳐
**서로 반대되는 지시**가 들어간 적이 두 번 있었다. 파일 동일성만 보는
check_skill_sync 로는 잡히지 않는다. 여기서는 코드에서 사실을 뽑아
문서의 주장과 대조한다. 스킬을 고칠 때마다 커밋 전에 돌린다.

  1  공용 파일 동일성 — 두 스킬이 같아야 하는 파일
  2  export ↔ 문서 — 내보내는데 문서에 없는 이름, 문서가 부르는데 없는 이름
  3  점검기 규칙 번호 — docstring 목록과 구현 주석의 번호가 맞는가
  4  좌표 — ksa_mono.js MODE / check_layout.py ZONE / 문서의 좌표가 같은가
  5  수치 주장 — '예문 N개' '예시 N장' 'N종'이 실제와 맞는가
  6  폐기 문구 — 규칙에서 뺀 표현이 문서에 남았는가
  7  명령어 — --special 값이 파일마다 같은가, cp 대상 파일이 실재하는가
  8  특정 기관명 — 다른 기관의 과업 내용이 예시에 박혔는가

FAIL이 하나라도 있으면 종료 코드 1.
"""
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent / ".claude" / "skills"
SKILLS = ("present-ppt", "report-ppt")
DOCS = ("SKILL.md", "INSTALL.md", "VERSION.md",
        "references/components.md", "references/design-system.md",
        "references/diagrams.md", "references/headline.md")
SHARED = ("assets/ksa_mono.js", "assets/photo_placeholder.png",
          "scripts/postprocess.py", "scripts/check_layout.py", "scripts/measure_head.py",
          "references/diagrams.md", "references/headline.md")
# design-system.md 는 본문 상단이 모드마다 달라 두 스킬이 서로 다르다 — 공용 대상이 아니다.

# 규칙에서 뺐는데 문서에 남으면 지시가 충돌한다
RETIRED = {
    r"지양 필요": "헤드 문법에서 허용으로 바뀐 표현",
    r"뜻 없는 외래어": "금지 목록에서 뺀 항목",
    r"공문 축약형": "금지 목록에서 뺀 항목",
    # '클로징 장표는 두지 않는다' 같은 부정문은 정상이므로 뺀다
    r"클로징(?!\s*장표는 (두지|없))(?![^\n]*없)": "클로징 장표는 없앴다",
    r"정중한 개조식": "명사형 개조식으로 정리했다",
    r"logoAt|ksa_logo": "로고를 스킬에서 제거했다",
    r"B\.sub\(": "sub()는 제거했다",
}
# 다른 기관의 과업 내용이 예시에 박히면 안 된다
ORG_WORDS = ("한국국제문화교류진흥원", "국제문화교류진흥법", "문화ODA", "기획예산처",
             "서울메트로", "도시철도공사", "시정 주요분야")
# 상수·내부용이라 문서에 없어도 되는 export
EXPORT_OK = {"P", "K", "W", "H", "M", "CW", "BB", "Z", "FONT", "TEXT_MIN",
             "TITLE_Y", "col", "cx", "cw", "split", "textW", "lines", "needH",
             "zones", "guard", "seriesRamp"}

fails, notes = [], []


def read(sk, rel):
    p = ROOT / sk / rel
    return p.read_text() if p.exists() else ""


def alldocs(sk):
    return "\n".join(read(sk, d) for d in DOCS)


# ── 1 공용 파일 동일성 ────────────────────────────────────────
def check_shared():
    import hashlib
    for rel in SHARED:
        ds = {}
        for sk in SKILLS:
            p = ROOT / sk / rel
            ds[sk] = hashlib.sha256(p.read_bytes()).hexdigest()[:12] if p.exists() else None
        if None in ds.values():
            fails.append(f"[공용] 없음 {rel} — {', '.join(k for k, v in ds.items() if v is None)}")
        elif len(set(ds.values())) > 1:
            fails.append(f"[공용] 불일치 {rel} — " + " / ".join(f"{k} {v}" for k, v in ds.items()))


# ── 2 export ↔ 문서 ──────────────────────────────────────────
def check_exports():
    for sk in SKILLS:
        src = read(sk, "assets/ksa_mono.js")
        m = re.search(r"module\.exports\s*=\s*\{(.*?)\};", src, re.S)
        if not m:
            fails.append(f"[{sk}] module.exports 를 못 찾음")
            continue
        exports = {x.strip() for x in m.group(1).replace("\n", " ").split(",") if x.strip()}
        docs = alldocs(sk)
        missing = sorted(e for e in exports if e not in EXPORT_OK and e not in docs)
        if missing:
            fails.append(f"[{sk}] 내보내는데 문서에 없음: {', '.join(missing)}")
        # B.name( 과 표 안의 `name(s, …)` `name(deck, …)` 둘 다 본다
        called = {x for x in re.findall(r"B\.([a-zA-Z_]\w*)\s*\(", docs)}
        called |= {x for x in re.findall(r"`([a-zA-Z_]\w*)\((?:s|deck)[,)]", docs)}
        ghost = sorted(called - exports)
        if ghost:
            fails.append(f"[{sk}] 문서가 부르는데 없는 이름: {', '.join(ghost)}")


# ── 3 점검기 규칙 번호 ───────────────────────────────────────
def check_rule_numbers():
    for sk in SKILLS:
        src = read(sk, "scripts/check_layout.py")
        doc = re.search(r"검사 항목\n(.*?)\n\nFAIL", src, re.S)
        if not doc:
            fails.append(f"[{sk}] 점검기 docstring 의 '검사 항목' 목록을 못 찾음")
            continue
        listed = re.findall(r"^\s{1,4}(\d+b?)\s{1,2}", doc.group(1), re.M)
        impl = []
        for grp in re.findall(r"^\s+# ((?:\d+b?)(?:·\d+b?)*)[ ]", src, re.M):
            impl += [x for x in grp.split("·") if x not in impl]
        # 문서에만 있는 번호는 정상이다(주석 없이 구현된 규칙이 있다)
        only_impl = [x for x in impl if x not in listed]
        if only_impl:
            fails.append(f"[{sk}] 구현에만 있는 규칙 번호: {', '.join(only_impl)}")
        nums = [int(re.sub(r"b$", "", x)) for x in listed]
        gaps = [n for n in range(1, max(nums) + 1) if n not in nums] if nums else []
        if gaps:
            fails.append(f"[{sk}] 규칙 번호가 건너뜀: {', '.join(map(str, gaps))}")


# ── 4 좌표 ───────────────────────────────────────────────────
def check_coords():
    for sk in SKILLS:
        src = read(sk, "assets/ksa_mono.js")
        mode = "present" if sk == "present-ppt" else "report"
        m = re.search(rf"\n  {mode}:\s*\{{(.*?)\}},\n", src, re.S)
        if not m:
            fails.append(f"[{sk}] MODE.{mode} 를 못 찾음")
            continue
        vals = dict(re.findall(r"(\w+):\s*([\d.]+)", m.group(1)))
        bt, msg_y = vals.get("bt"), vals.get("msgY")
        chk = read(sk, "scripts/check_layout.py")
        z = re.search(rf'"{mode}":\s*\{{([^}}]*)\}}', chk)
        if z:
            zv = dict(re.findall(r'"([^"]+)":\s*([\d.]+)', z.group(1)))
            for key, got, want in (("본문", zv.get("본문"), bt), ("메시지", zv.get("메시지"), msg_y)):
                if got and want and abs(float(got) - float(want)) > 0.005:
                    fails.append(f"[{sk}] 점검기 ZONE {key} {got}\" ≠ MODE {want}\"")
        # 문서가 다른 본문 상단을 적고 있는가
        pats = (r'본문 상단[^\d\n]{0,8}(\d\.\d\d)"',      # "본문 상단 2.00""
                r'본문[^\n]{0,4}?\*?\*?(\d\.\d\d)"[–-]6\.85"')  # "본문 2.00"–6.85""
        other = "report" if mode == "present" else "present"
        om = re.search(rf"\n  {other}:\s*\{{(.*?)\}},\n", src, re.S)
        other_bt = dict(re.findall(r"(\w+):\s*([\d.]+)", om.group(1))).get("bt") if om else None
        for d in ("SKILL.md", "INSTALL.md", "references/design-system.md"):
            # 두 모드를 나란히 적은 비교 표 줄은 뺀다
            text = "\n".join(l for l in read(sk, d).split("\n")
                             if not (other_bt and other_bt in l and bt in l))
            for pat in pats:
                for found in set(re.findall(pat, text)):
                    if abs(float(found) - float(bt)) > 0.005:
                        fails.append(f"[{sk}] {d} 의 본문 상단 {found}\" ≠ 코드 {bt}\"")


# ── 5 수치 주장 ──────────────────────────────────────────────
def check_claims():
    for sk in SKILLS:
        docs = alldocs(sk)
        # 헤드 예문 개수
        head = read(sk, "references/headline.md")
        sec = re.search(r"## 3\. 장표 유형별 예문(.*?)\n## 4\.", head, re.S)
        real = len(re.findall(r"^- ", sec.group(1), re.M)) if sec else 0
        for claimed in set(re.findall(r"예문 (\d+)개", docs)):
            if int(claimed) != real:
                fails.append(f"[{sk}] '예문 {claimed}개' — 실제 {real}개")
        # 예시 덱 장수
        deck = read(sk, "assets/example_deck.js")
        n = len(re.findall(r"=\s*S\(", deck)) + len(re.findall(r"B\.(cover|toc)\(deck", deck))
        for claimed in set(re.findall(r"예시 (\d+)장|(\d+)장 예시", docs)):
            c = int(claimed[0] or claimed[1])
            if c != n:
                fails.append(f"[{sk}] '예시 {c}장' — example_deck.js 는 {n}장")
        # 도식·차트 종수
        src = read(sk, "assets/ksa_mono.js")
        dia = sum(1 for f in ("venn", "hubSpoke", "cycle", "pyramid", "steps", "harvey", "causeEffect")
                  if f"function {f}(" in src)
        cht = sum(1 for f in ("barChart", "lineChart", "pieChart", "comboChart")
                  if f"function {f}(" in src)
        for claimed in set(re.findall(r"원인→결과 (\d+)종", docs)):
            if int(claimed) != dia:
                fails.append(f"[{sk}] '관계 도식 {claimed}종' — 실제 {dia}종")
        for claimed in set(re.findall(r"혼합 (\d+)종", docs)):
            if int(claimed) != cht:
                fails.append(f"[{sk}] '차트 {claimed}종' — 실제 {cht}종")


# ── 6 폐기 문구 ──────────────────────────────────────────────
def check_retired():
    for sk in SKILLS:
        for d in DOCS:
            text = read(sk, d)
            for word, why in RETIRED.items():
                m = re.search(word, text)
                if m:
                    # headline.md 는 '이렇게 쓰지 않는다' 대비표라 예외
                    if d == "references/headline.md" and "지양" in word or "축약형" in word:
                        continue
                    fails.append(f"[{sk}] {d} 에 폐기 문구 '{m.group(0)}' — {why}")


# ── 7 명령어 ─────────────────────────────────────────────────
def check_commands():
    for sk in SKILLS:
        vals = set()
        for d in ("SKILL.md", "INSTALL.md"):
            vals |= set(re.findall(r"--special ([\d,]+)", read(sk, d)))
        if len(vals) > 1:
            fails.append(f"[{sk}] --special 값이 파일마다 다름: {' / '.join(sorted(vals))}")
        for d in ("SKILL.md", "INSTALL.md"):
            for grp in re.findall(r"\$SK/assets/\{([^}]*)\}", read(sk, d)):
                for f in (x.strip() for x in grp.split(",")):
                    if f and not (ROOT / sk / "assets" / f).exists():
                        fails.append(f"[{sk}] {d} 가 없는 파일을 복사한다: assets/{f}")
            for f in re.findall(r"\$SK/scripts/(\w+\.py)", read(sk, d)):
                if not (ROOT / sk / "scripts" / f).exists():
                    fails.append(f"[{sk}] {d} 가 없는 스크립트를 부른다: scripts/{f}")


# ── 8 특정 기관명 ────────────────────────────────────────────
def check_orgs():
    for sk in SKILLS:
        for d in DOCS + ("assets/example_deck.js",):
            text = read(sk, d)
            for w in ORG_WORDS:
                if w in text:
                    fails.append(f"[{sk}] {d} 에 특정 기관·사업명 '{w}' — 예시는 일반 문장으로")


def main():
    for fn in (check_shared, check_exports, check_rule_numbers, check_coords,
               check_claims, check_retired, check_commands, check_orgs):
        fn()
    for n in notes:
        print("참고 " + n)
    for f in fails:
        print("FAIL " + f)
    print(f"\n{len(fails)} fail")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
