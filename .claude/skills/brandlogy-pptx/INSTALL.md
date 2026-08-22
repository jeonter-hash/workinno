# brandlogy-pptx 설치와 사용법

한국표준협회(KSA) 브랜드 PPT 디자인 시스템 스킬. **A4 가로 · 맑은 고딕 · 화이트 캔버스 · 브랜드 블루(#1456f0)** 규칙으로 슬라이드를 만들고, 그 규칙을 코드로 강제한다.

---

## 1. 설치

폴더 하나를 통째로 `skills` 아래에 넣으면 끝난다. 어디에 넣느냐만 고르면 된다.

### A. 개인 전용 (모든 프로젝트에서 쓰기)

```bash
mkdir -p ~/.claude/skills
unzip brandlogy-pptx.zip -d ~/.claude/skills/
# 결과: ~/.claude/skills/brandlogy-pptx/SKILL.md
```

### B. 특정 저장소/프로젝트 전용 (팀과 git으로 공유)

```bash
mkdir -p <프로젝트>/.claude/skills
unzip brandlogy-pptx.zip -d <프로젝트>/.claude/skills/
git add .claude/skills/brandlogy-pptx && git commit -m "brandlogy-pptx 스킬 추가"
```

### C. claude.ai / Cowork 웹에서 쓰기

설정의 스킬(기능) 관리 화면에서 이 **zip 파일을 그대로 업로드**한다. 메뉴 이름은 버전에 따라 다를 수 있다.

### 설치 확인

새 세션을 열고 `/` 를 눌렀을 때 목록에 `brandlogy-pptx`가 보이면 된다. 안 보이면 폴더 구조를 확인한다 — `skills/brandlogy-pptx/SKILL.md` 경로가 정확해야 하고, zip 안에 폴더가 한 겹 더 들어가 있으면 인식하지 않는다.

---

## 2. 필요한 것

| 항목 | 비고 |
|---|---|
| Node.js + `pptxgenjs` | 슬라이드 생성. Claude의 pptx 스킬 환경에는 이미 설치돼 있다. 없으면 `npm install pptxgenjs` |
| Python 3 | 후처리·점검 스크립트. **표준 라이브러리만** 쓰므로 추가 설치 불필요 |
| 맑은 고딕 | 완성 파일을 여는 PC에 필요(Windows 기본 탑재). 없는 환경(macOS·리눅스)에는 PDF로 전달한다 |
| KSA 로고 | `assets/ksa_logo.jpg`로 **이미 포함**돼 있다 |

---

## 3. 사용법 — 그냥 말로 시킨다

설치했다면 별도 명령 없이 아래처럼 요청하면 스킬이 자동으로 걸린다.

- "2026년 사업계획 브랜드 덱 만들어줘"
- "이 데이터로 A4 가로 슬라이드 6장 만들어줘"
- "KPI 카드 4개랑 채널별 추이 차트로 장표 하나"
- "표지 + 섹션 디바이더 포함해서 덱 구성해줘"

안 걸리면 `/brandlogy-pptx` 로 직접 부른다.

**자료를 함께 주면 좋다** — 숫자·출처가 있어야 차트를 만든다. 근거 없는 수치는 스킬이 지어내지 않는다.

---

## 4. 직접 만들 때 (개발자용)

```bash
SK=~/.claude/skills/brandlogy-pptx          # 설치 위치에 맞게
mkdir build && cp $SK/assets/{brandlogy.js,example_deck.js,ksa_logo.jpg} build/ && cd build

node example_deck.js                         # 예시 4장 생성 (로고 자동 포함)
python3 $SK/scripts/postprocess.py example_deck.pptx    # 그라디언트 + 차트 한글 폰트
python3 $SK/scripts/check_layout.py example_deck.pptx   # 디자인 규칙 자동 점검
```

`example_deck.js`를 복사해 내용만 갈아끼우는 방식으로 쓴다. **좌표는 직접 쓰지 않는다** — 헬퍼가 갖고 있다.

```js
const B = require('./brandlogy.js');
const deck = B.createDeck({ title: '2026 사업계획' });     // 로고 자동

const s = deck.slide({ chapter: '01 시장 진단', source: '출처: 통계청(2025)' });
B.headline(s, '성장은 브랜드 검색 한 곳에서 나왔다');       // 32pt, 한글 22자 내외 한 줄
B.subtitle(s, '브랜드 검색 세션만 연 20% 증가, 전환율은 2.4배');

B.kpiRow(s, [                                              // 상단 KPI 4장
  { value: '+42.6%', label: '브랜드 검색 유입' },
  { value: '2.4x',   label: '전환율' },
  { value: '-8.1%',  label: '비브랜드 검색' },
  { value: '71%',    label: '상위 3채널 기여도', gradient: true },
]);

const d = B.dataCard(s, { x: B.Z.body.x, y: B.BAND.A.detail.y,   // 하단 차트 컨테이너
  w: B.Z.body.w, h: B.BAND.A.detail.h, title: '채널별 세션 추이', source: '출처: GA4' });
s.addChart('bar', data, B.chartOpts({ ...d.area, barDir: 'col' }));

await deck.save('deck.pptx');
```

주요 함수: `headline` `subtitle` `kpiRow` `dataCard` `chartOpts` `bullets` `h2` `soWhat` `pill` `caption` `card` `cover` `divider` `split(n)`

---

## 5. 지켜지는 규칙 (자동 강제)

**만들 때 막는 것** — 어기면 파일 생성이 중단된다.

- 본문이 2.39"–6.85" 밖으로 나가면 → 예외
- Hero Gradient 장표당 2개 이상, 덱 전체 4개 이상 → 예외
- Brand Glow 그림자 장표당 2개 이상 → 예외
- 이모지 → 예외

**만든 뒤 잡는 것** (`check_layout.py`)

A4 판형 · 5존 좌표 이탈 · 맑은 고딕 외 폰트(차트 내부까지) · 9pt 미만 · 본문 경계/푸터 여백 침범 · 하단 30% 공백 · 로고 위치·비율·**로고 뒤 도형** · 팔레트 밖 색상 · 차트 없는 데이터 장표(경고)

---

## 6. 알아둘 제약

- **동봉 로고는 흰 배경 JPEG(누끼 아님)** 이다. 흰 장표에서는 문제없지만 어두운 배경·그라디언트 장표에 올리면 흰 사각형이 보인다. 그런 장표에 로고가 필요하면 투명 PNG나 흰 반전본을 `logoWhite`로 넘긴다.
- **헤드라인은 한 줄**이다. A4 가로 32pt에서 한글 22자 안팎이 한계이고, 두 줄이 되면 부제 존을 침범한다. 길면 폰트를 줄이지 말고 문장을 줄인다.
- **맑은 고딕은 굵기가 두 단계**(Regular/Bold)뿐이다. 원 디자인의 500 vs 600 대비는 사라지므로 크기와 색으로 구분한다.
- 판형은 A4 가로만 낸다. 16:9가 필요하면 별도 요청해야 한다(좌표계가 다르다).

---

## 7. 파일 구성

| 파일 | 내용 |
|---|---|
| `SKILL.md` | 스킬 본문 — 강제 규칙, 워크플로, 패턴 맵, QA |
| `references/design-system.md` | 규범 사양 — 색·타이포·컴포넌트·Do/Don't·체크리스트 |
| `references/layout-geometry.md` | 인치 확정 좌표 — 5존, 12열 그리드, 패턴 밴드, 차트 기본값 |
| `assets/brandlogy.js` | 생성 헬퍼(좌표·부품·가드레일) |
| `assets/example_deck.js` | 예시 코드 4장 |
| `assets/ksa_logo.jpg` | 한국표준협회 워드마크(325×42) |
| `scripts/postprocess.py` | Hero Gradient 벡터화 + 차트 한글 폰트 |
| `scripts/check_layout.py` | 디자인 규칙 자동 점검 |
| `INSTALL.md` | 이 문서 |
