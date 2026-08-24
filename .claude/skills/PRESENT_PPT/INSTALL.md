# PRESENT_PPT 설치와 사용법

한국표준협회(KSA) 모노톤 **발표용** PPT 스킬. A4 가로 · 맑은 고딕 · 흑백 무채색.

---

## 1. 설치

폴더를 통째로 `skills` 아래에 넣으면 끝난다.

```bash
# 개인 전용 (모든 프로젝트)
unzip PRESENT_PPT.zip -d ~/.claude/skills/

# 또는 특정 저장소 전용 (팀과 git 공유)
unzip PRESENT_PPT.zip -d <프로젝트>/.claude/skills/
```

claude.ai·Cowork 웹이면 스킬 관리 화면에서 이 zip을 그대로 올린다.
새 세션에서 `/`를 눌렀을 때 목록에 `PRESENT_PPT`가 보이면 설치된 것이다.

## 2. 필요한 것

| 항목 | 비고 |
|---|---|
| Node.js + `pptxgenjs` | 슬라이드 생성. pptx 스킬 환경에는 이미 설치돼 있다 |
| Python 3 | 후처리·점검 스크립트. **표준 라이브러리만** 쓴다 |
| 맑은 고딕 | 완성 파일을 여는 PC에 필요(Windows 기본). 없는 환경엔 PDF로 전달 |
| 색 | 무채색 + KSA 강조 2색(네이비·레드)이 기본. 무채색만 쓰려면 `palette:'mono'` |

## 3. 사용법 — 말로 시킨다

- "발표용 자료 만들어줘"
- "이 데이터로 장표 만들어줘"

안 걸리면 `/PRESENT_PPT`로 직접 부른다. 숫자·출처가 있어야 도식을 만든다 — 근거 없는 수치는 지어내지 않는다.

## 4. 직접 만들 때

```bash
SK=~/.claude/skills/PRESENT_PPT
mkdir build && cp $SK/assets/{ksa_mono.js,example_deck.js,photo_placeholder.png} build/ && cd build
node example_deck.js deck.pptx                        # 발표용 11장 예시 생성
python3 $SK/scripts/postprocess.py deck.pptx          # 필수
python3 $SK/scripts/check_layout.py deck.pptx --mode present --special 1,12
```

`example_deck.js`를 복사해 내용만 갈아끼운다. **좌표는 직접 쓰지 않는다** — 컴포넌트가 갖고 있다.

## 5. 자동으로 막히는 것

생성 단계에서 예외를 던져 **파일이 만들어지지 않는다**:

- 본문이 2.11"–6.85" 밖으로 나감
- 표 열 폭 합계가 우측 여백 초과 / 행의 칸 수 불일치
- 네이티브 표가 글 길이 때문에 자라서 벽을 넘음
- 불릿이 지정한 상자 높이를 넘침

만든 뒤 `check_layout.py`가 잡는 것: A4 판형 · 맑은 고딕 외 폰트(표·차트 내부까지) · 최소 글자 크기 · 슬라이드 밖 이탈 · 본문 벽 침범(네이티브 표는 실제 행 높이 합산) · 5존 앵커 이탈(모드별) · 하단 30% 공백 · 본문 빈 띠 · 요소 겹침 · 의도하지 않은 검은 윤곽선 · 흐린 글자색 · 팔레트 밖 색 · 이모지 · 차트 구조 결함 · 헤드 문법(타이틀 길이·메시지 종결·줄 수) · 각주 과다 · 결론 밴드 남용 · 문안 · 표 편중.

## 6. 알아둘 것

- `postprocess.py`를 빼면 **PowerPoint가 "복구하시겠습니까" 대화상자**를 띄운다(pptxgenjs가 차트에 없는 축을 참조).
- 렌더 QA는 반드시 한다 — 겹침·넘침은 좌표 검사로 안 걸린다.
- 화면에 띄우는 자료라면 **REPORT_PPT**(보고서용)를 쓴다.

## 7. 파일 구성

| 파일 | 내용 |
|---|---|
| `SKILL.md` | 스킬 본문 — 밀도, 워크플로, 도식 선택, QA |
| `references/design-system.md` | 규범 — 판형·색·5존·밀도·표·선·문안 |
| `references/components.md` | 컴포넌트 API |
| `assets/ksa_mono.js` | 공용 모듈(두 스킬 동일) |
| `assets/example_deck.js` | 11장 예시 |
| `scripts/postprocess.py` | 차트 결함 보정(필수) |
| `scripts/check_layout.py` | 자동 점검 |

## 표지 사진 바꾸기

표지는 사진 5장(상단 띠 4장 + 우하단 대형 1장)으로 짜여 있다. 사진을 주지 않으면 회색 자리표시가 들어간 채로 나온다.

**PowerPoint에서 바꾸기** — 자리표시를 오른쪽 클릭 → **[그림 바꾸기] → [파일에서]**.
사진마다 독립된 도형이라 사선 모양과 위치가 그대로 유지된다. 잘린 범위는 [그림 서식 → 자르기]에서 조정한다.

**만들 때 넣기** — `photos`에 경로 5개를 준다.

```js
B.cover(deck, { org:'○○공단', title:'…', date:'2026.01', by:'한국표준협회',
  photos:['공연.jpg','전경.jpg','교실.jpg','축제.jpg','대표.jpg'] });
```

가로·세로 비율이 달라도 중앙을 기준으로 잘라 채운다. 띠 사진은 세로로 긴 구도, 대형 사진은 가로로 넓은 구도가 잘 맞는다.

**사진 출처** — 고객사가 제공한 사진을 우선한다. 웹에서 가져올 때는 공공누리 1유형 또는 CC0/CC-BY(출처 표기)만 쓰고, 출처가 불분명한 이미지는 제출용 문서에 넣지 않는다.
