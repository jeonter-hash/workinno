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

### PRESENT_PPT — 발표용 PPT (KSA 모노톤)

`.claude/skills/PRESENT_PPT/`

A4 가로 · 맑은 고딕 · **흑백 무채색**으로 한 장에 메시지 하나와 도식 하나를 담는 스크린 발표 덱. 헤드라인 24pt · 본문 11.5pt.

### REPORT_PPT — 보고서용 PPT (KSA 모노톤)

`.claude/skills/REPORT_PPT/`

같은 프레임·같은 도식에 밀도만 높인 읽는 문서형 장표. 헤드라인 20pt · 본문 9.5pt · 표 행 0.32", PowerPoint 네이티브 표 기본.

두 스킬은 같은 모듈(`assets/ksa_mono.js`)을 공유한다.

| 파일 | 내용 |
|---|---|
| `SKILL.md` | 밀도, 생성 워크플로, 표준 골격 12장, 도식 선택, QA, 자주 하는 실수 |
| `INSTALL.md` | 설치·사용 설명서 |
| `references/design-system.md` | 규범 — 판형(A4 가로)·무채색 팔레트·5존·모드별 밀도·표·선·문안 |
| `references/components.md` | 컴포넌트 API — 배치·표·도식(체브론·트리·매트릭스·간트·워터폴·계층) |
| `assets/ksa_mono.js` | 공용 모듈. 본문 벽 침범·표 폭 초과·네이티브 표 자람을 예외로 차단 |
| `assets/example_deck.js` | 12장 예시(모드별) |
| `assets/ksa_logo.jpg` | 한국표준협회 워드마크(325×42) |
| `scripts/postprocess.py` | 차트 결함 보정 — 미선언 축 참조 제거(**미실행 시 PowerPoint 복구 경고**), dPt 순서, 한글 폰트 |
| `scripts/check_layout.py` | 자동 점검 — 판형·폰트·최소 크기·경계·5존·밀도·검은 윤곽선·흐린 글자·유채색·로고·이모지·차트 구조. **네이티브 표는 `<a:tr h>` 합산으로 실제 높이 계산** |

파일 조작은 pptx 스킬의 표준 절차를 따른다.
