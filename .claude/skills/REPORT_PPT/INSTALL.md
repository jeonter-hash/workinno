# REPORT_PPT 설치와 사용법

한국표준협회(KSA) 모노톤 **보고서용** PPT 스킬. A4 가로 · 맑은 고딕 · 흑백 무채색.

---

## 1. 설치

폴더를 통째로 `skills` 아래에 넣으면 끝난다.

```bash
# 개인 전용 (모든 프로젝트)
unzip REPORT_PPT.zip -d ~/.claude/skills/

# 또는 특정 저장소 전용 (팀과 git 공유)
unzip REPORT_PPT.zip -d <프로젝트>/.claude/skills/
```

claude.ai·Cowork 웹이면 스킬 관리 화면에서 이 zip을 그대로 올린다.
새 세션에서 `/`를 눌렀을 때 목록에 `REPORT_PPT`가 보이면 설치된 것이다.

## 2. 필요한 것

| 항목 | 비고 |
|---|---|
| Node.js + `pptxgenjs` | 슬라이드 생성. pptx 스킬 환경에는 이미 설치돼 있다 |
| Python 3 | 후처리·점검 스크립트. **표준 라이브러리만** 쓴다 |
| 맑은 고딕 | 완성 파일을 여는 PC에 필요(Windows 기본). 없는 환경엔 PDF로 전달 |
| KSA 로고 | `assets/ksa_logo.jpg`로 **포함**돼 있다 |

## 3. 사용법 — 말로 시킨다

- "보고서용 자료 만들어줘"
- "이 데이터로 장표 만들어줘"

안 걸리면 `/REPORT_PPT`로 직접 부른다. 숫자·출처가 있어야 도식을 만든다 — 근거 없는 수치는 지어내지 않는다.

## 4. 직접 만들 때

```bash
SK=~/.claude/skills/REPORT_PPT
mkdir build && cp $SK/assets/{ksa_mono.js,example_deck.js,ksa_logo.jpg} build/ && cd build
node example_deck.js deck.pptx                        # 12장 예시 생성
python3 $SK/scripts/postprocess.py deck.pptx          # 필수
python3 $SK/scripts/check_layout.py deck.pptx --mode report --special 1
```

`example_deck.js`를 복사해 내용만 갈아끼운다. **좌표는 직접 쓰지 않는다** — 컴포넌트가 갖고 있다.

## 5. 자동으로 막히는 것

생성 단계에서 예외를 던져 **파일이 만들어지지 않는다**:

- 본문이 2.39"–6.85" 밖으로 나감
- 표 열 폭 합계가 우측 여백 초과 / 행의 칸 수 불일치
- 네이티브 표가 글 길이 때문에 자라서 벽을 넘음

만든 뒤 `check_layout.py`가 잡는 것: A4 판형 · 맑은 고딕 외 폰트(표·차트 내부까지) · 9pt 미만 · 슬라이드 밖 이탈 · 본문 벽 침범(네이티브 표는 실제 행 높이 합산) · 5존 앵커 이탈 · 하단 30% 공백 · 의도하지 않은 검은 윤곽선 · 흐린 글자색 · 유채색 · 로고 위치/비율/로고 뒤 도형 · 이모지 · 차트 구조 결함.

## 6. 알아둘 것

- `postprocess.py`를 빼면 **PowerPoint가 "복구하시겠습니까" 대화상자**를 띄운다(pptxgenjs가 차트에 없는 축을 참조).
- 동봉 로고는 흰 배경 JPEG다. 어두운 배경 장표에 넣으려면 흰 반전본이 필요하다.
- 렌더 QA는 반드시 한다 — 겹침·넘침은 좌표 검사로 안 걸린다.
- 화면에 띄우는 자료라면 **PRESENT_PPT**(발표용)를 쓴다.

## 7. 파일 구성

| 파일 | 내용 |
|---|---|
| `SKILL.md` | 스킬 본문 — 밀도, 워크플로, 도식 선택, QA |
| `references/design-system.md` | 규범 — 판형·색·5존·밀도·표·선·문안 |
| `references/components.md` | 컴포넌트 API |
| `assets/ksa_mono.js` | 공용 모듈(두 스킬 동일) |
| `assets/example_deck.js` | 12장 예시 |
| `assets/ksa_logo.jpg` | 한국표준협회 워드마크 |
| `scripts/postprocess.py` | 차트 결함 보정(필수) |
| `scripts/check_layout.py` | 자동 점검 |
