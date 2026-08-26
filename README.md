# workinno

업무 혁신용 스킬의 문서와 점검 도구.

스킬 원본은 **코워크(claude.ai)** 에 있다. 이 기기에는 계정 동기화본이
`~/.claude/skills/synced/` 로 내려오며, 스킬 3종은 그 한 벌만 쓴다.
이 저장소는 사본을 두지 않는다 — 두 벌이 되면 어느 쪽이 실제로 적용되는지
알 수 없기 때문이다. 고칠 일이 있으면 코워크에서 고치고 다시 올린다.

## 스킬

### report-contract — 보고서 논리·문안 계약

`~/.claude/skills/synced/report-contract/`

컨설팅 보고서의 논리 구조와 문안 언어를 계약으로 고정한다. 조판 스킬 위에 얹는 상위 계약이며, 두 조판 스킬이 §0에서 이 스킬을 먼저 읽도록 지시한다.

| 파일 | 내용 |
|---|---|
| `SKILL.md` | 3계층 재귀 구조 계약(프레임→적용→결론), 미결정 안건 5슬롯, 문안 언어 계약, 자기검증 게이트 G1~G9 |
| `references/defect-catalog.md` | 결함 카탈로그 D1~D16 |
| `references/ban-dictionary.md` | 금지 표현 사전 — 메타 표현·학술 용어·구어체·수사 표기와 대체어 |
| `references/build-traps.md` | 대규모 덱 제작 함정 |
| `scripts/meta_check.py` | 산출물 자동 점검 — M1~M8·S1~S3·H1. **전 규칙 0건이 제출 조건** |

### present-ppt — 발표용 PPT (KSA 모노톤 + 강조 2색)

`~/.claude/skills/synced/present-ppt/`

A4 가로 · 맑은 고딕 · 흰 캔버스. 헤드를 타이틀과 메시지로 나누고 한 장에 메시지 하나와 도식·차트 하나를 담는 스크린 발표 덱. 타이틀 24pt · 메시지 13pt · 본문 11.5pt · 본문 2.11–6.85".
구조색 네이비(#1F3864)와 강조색(#C4303C)을 **장표당 한 곳**만 쓴다.

### report-ppt — 보고서용 PPT (KSA 모노톤)

`~/.claude/skills/synced/report-ppt/`

같은 프레임·같은 도식에 밀도만 높인 읽는 문서형 장표. 타이틀 20pt · 메시지 12pt · 본문 9.5pt · 표 행 0.335" · 본문 2.00–6.85", 색은 무채색만.
내용이 많으면 `dense` 옵션으로 본문 9pt까지 낮춘다. 8.5pt가 절대 하한.

두 조판 스킬은 같은 모듈(`assets/ksa_mono.js`)과 같은 점검기를 공유하며, 표지·목차·슬라이드 마스터도 동일하다.

| 파일 | 내용 |
|---|---|
| `SKILL.md` | 밀도, 헤드 문법, 생성 워크플로, 표준 골격, 도식·차트 선택, 문안 규칙, QA |
| `INSTALL.md` | 설치·사용 설명서 |
| `references/design-system.md` | 규범 — A4 가로·팔레트·5존·모드별 밀도·표·선·차트 |
| `references/components.md` | 컴포넌트 API — 헤드·표지·목차·표·도식·차트·배치 |
| `references/diagrams.md` | 의미에서 도식·차트·표를 고르는 기준 |
| `references/headline.md` | 헤드 문법 — 타이틀·메시지 규격과 유형별 예문 |
| `assets/ksa_mono.js` | 공용 모듈. 마스터(자동 번호·문서 제목), 사진 표지·목차, 본문 벽·표 폭·헤드 길이를 예외로 차단 |
| `assets/example_deck.js` | 예시 덱(발표용 11장 / 보고서용 13장) |
| `assets/photo_placeholder.png` | 표지 사진 자리표시 |
| `scripts/postprocess.py` | **필수 후처리** — 차트 축 참조·한글 폰트·표지 사선·피라미드 기울기 보정 |
| `scripts/check_layout.py` | 자동 점검기 — 판형·폰트·좌표·겹침·헤드 문법·문안·표 편중 등 15규칙 |
| `scripts/measure_head.py` | 렌더에서 헤드 글자 띠 경계를 실측 |

## 도구

| 파일 | 내용 |
|---|---|
| `tools/audit_skills.py` | 스킬 문서가 코드와 어긋났는지 점검 — 공용 파일·export·규칙 번호·좌표·수치 주장·폐기 문구·명령어·기관명. 인자 없이 돌리면 동기화본을 본다 |

```
python3 tools/audit_skills.py            # ~/.claude/skills/synced/
python3 tools/audit_skills.py <폴더>      # 다른 사본
```
