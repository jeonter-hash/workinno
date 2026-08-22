# workinno

업무 혁신용 Claude Code 스킬 저장소.

## 스킬

### present-pptx — 발표자료 PPT (KSA 하우스 스타일)

`.claude/skills/present-pptx/`

제안발표·심사PT·사업설명회용 발표 덱을 표준 패턴 라이브러리(31장)를 복제해 채우는 방식으로 만드는 스킬. 원본 템플릿(산업·일자리전환 지원센터 제안발표 덱)의 폰트·서식·도식·도형·디자인·도해 구성을 세밀 분석해 준용한다.

| 파일 | 내용 |
|---|---|
| `SKILL.md` | 디자인 토큰(HY헤드라인M·맑은 고딕, 페리윙클 #5D6DBE·비비드블루 #1040C3·틸 #58A7AE·레드 #E35740 색 체계), 7존 장표 문법, 헤드라인·요약 밴드 작법, 표준 골격, 제작 워크플로, QA 규칙 |
| `references/pattern-catalog.md` | 31개 패턴 상세 카탈로그 — 패턴별 용도·구성·부품(도형 좌표/색/폰트 수치)·슬롯 가감법·주의사항 |
| `assets/patterns.pptx` | 표준 패턴 라이브러리 (원본 발표 덱 31장, A4 가로) |

파일 조작(unzip → add_slide.py 복제 → XML 텍스트 교체 → zip → validate → 렌더 QA)은 pptx 스킬의 표준 절차를 따른다.

### brandlogy-pptx — 브랜드 PPT (A4 가로 디자인 시스템)

`.claude/skills/brandlogy-pptx/`

순백 캔버스 · 맑은 고딕 전용 · 브랜드 블루(#1456f0) 체계의 **A4 가로(10.8333" × 7.5")** 브랜드 덱을 만드는 스킬. 로고는 한국표준협회(KSA). 5존 고정 좌표(헤더 0.4" / 헤드라인 1.0" / 부제 1.63" / 본문 2.39–6.85" / 푸터 7.05")와 Hero Gradient 상한, 로고 무결성, 시각화 우선 규칙을 코드로 강제한다.

| 파일 | 내용 |
|---|---|
| `SKILL.md` | 3대 강제 규칙(존 고정·하드 경계·시각화 우선), 생성 워크플로, 패턴 선택 맵, 토큰 요약, QA 절차, 템플릿 경유 절차 |
| `references/design-system.md` | 규범 사양 원문 — 색·타이포·컴포넌트·레이아웃·Do/Don't·체크리스트 |
| `references/layout-geometry.md` | 인치 확정 좌표표 — 5존, 12열 그리드(1열 0.63611"), 패턴 A~F 밴드, 그림자·그라디언트 OOXML, 차트 기본값 |
| `assets/brandlogy.js` | pptxgenjs 헬퍼 — 존·그리드·밴드 상수, KPI/차트/콜아웃 부품, 경계 침범·그라디언트 초과 시 예외 |
| `assets/example_deck.js` | 4장 예시(표지·패턴 A·패턴 B·섹션 디바이더) — 복사해 내용만 교체 |
| `scripts/postprocess.py` | Hero Gradient 센티넬 → 벡터 `a:gradFill` 후처리(장표 1 / 덱 3 상한 검사) + 차트 파트 한글 폰트 주입 |
| `scripts/check_layout.py` | 체크리스트 자동 점검 — A4 판형·존 앵커·하드 경계·맑은 고딕·밀도·로고 위치/비율·로고 뒤 도형·이모지·팔레트 |

파일 조작(unzip → add_slide.py → XML 교체 → zip → validate → 렌더 QA)은 pptx 스킬의 표준 절차를 따른다.
