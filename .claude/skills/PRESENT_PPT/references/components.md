# 컴포넌트 API — `assets/ksa_mono.js`

좌표를 직접 타이핑하지 않는다. 아래 함수가 5존·그리드·밴드를 갖고 있고, 규칙을 어기면 **예외를 던져 생성을 중단**한다.

```js
const B = require('./ksa_mono.js');
const deck = B.createDeck({
  mode:'report',            // 'present' | 'report'
  palette:'mono',           // 'mono' | 'accent' (발표용 기본은 accent)
  dense:false,              // 보고서용에서만 — 내용이 많을 때 한 단계 낮춘 밀도
  title:'AX 전환 로드맵',
  docTitle:'AX 전환 로드맵', // 마스터 우측 상단에 반복 표기
});
const s = deck.slide({ chapter:'1. 요약', source:'출처: ...' });                  // 5존 프레임 자동
B.head(s, '헤드라인');   B.sub(s, '부제');
await deck.save('out.pptx');
```

## 상수

| 이름 | 뜻 |
|---|---|
| `K` | 무채색 팔레트 (`K.ink` `K.g1`~`K.g7` `K.w`) |
| `M` `CW` `W` `H` | 여백 0.5 · 콘텐츠 폭 9.8333 · 슬라이드 10.8333 × 7.5 |
| `BT` `BB` | 본문 상한 2.20 · 하한 6.85 (벽) |
| `Z` | 존 좌표 (`Z.body.x` 등) |
| `deck.T` | 현재 모드의 크기표 (`T.body` `T.rowH` `T.pad` …) |

## 배치

| 호출 | 설명 |
|---|---|
| `cx(i)` `cw(n)` | 12열 그리드 x좌표 / n열 폭 |
| `split(n)` | 균등 n분할 → `[{x,w}, …]` (n = 2·3·4·6) |
| `textW(str, pt)` | 글자 실폭(인치) — **밑줄 길이 산출용** |
| `lines(str, pt, w)` `needH(str, pt, w)` | 줄 수 / 필요한 높이 추정 |
| `guard(y, h, what)` | 본문 벽 검사 (컴포넌트가 자동 호출) |

## 조각

| 호출 | 설명 |
|---|---|
| `box(s, {x,y,w,h,fill,line,round,lw})` | 사각형. `line:'none'` 또는 `lw:0`이면 **테두리 없음** |
| `txt(s, text, {x,y,w,h,sz,b,c,align,valign,lh})` | 글상자 |
| `hr(s, {x,y,w,color,width})` | 수평선 |
| `underline(s, text, pt, {x,y})` | **글자 폭에 맞춘** 밑줄 |
| `sectionTitle(s, t, {x,y,w})` | 소제목 + 먹 구분선 |
| `logoAt(s, {})` | 로고를 우상단에 원본 비율로 (표지용) |
| `source(s, t, {x,y,w})` | 출처 — 도식·표 **바로 아래**에 붙인다. 푸터에 두지 않는다. **단위와 한 줄로 합친다**: `단위: 억 원 / 출처: …` |
| `footnote(s, t, {x,y,w})` | 각주 — **오독 위험이 있을 때만.** 한 장표에 3줄 이상이면 점검기가 경고한다 |

## 도형 안 텍스트

| 호출 | 설명 |
|---|---|
| `panel(s, {x,y,w,h,title,items,sz})` | **제목 박스.** 제목은 가운데 정렬 + 밑줄, 본문은 박스 높이에 맞춰 줄간격을 벌려 채운다. `items[i]`는 문자열 또는 `{t,sub}` |

```js
B.panel(s, { x:B.cx(0), y:2.46, w:B.cw(6), h:1.86, title:'추진 배경', items:[
  { t:'문화ODA가 법정 사무로 신설되어 ’26.11.20부터 시행', sub:'국제문화교류진흥법 제5조의2' },
  '기획예산처 재설치로 문화외교 추진체계와 재정 협의 창구가 함께 이동',
]});
```

제목을 왼쪽에 붙이고 글을 위로 몰면 카드 아래가 빈다. `panel()`이 이를 막는다. 글이 넘치면 필요한 높이를 알려주며 생성을 멈춘다.

## 표지·목차

| 호출 | 설명 |
|---|---|
| `cover(deck, {org, title, date, by, photos, panel})` | **사진 표지.** 사선 4분할 띠 + 먹 패널 + 대형 사진. 로고 없음. `photos` 생략 시 자리표시. 사진은 각각 독립 도형이라 PowerPoint에서 바꿔 끼운다 |
| `coverPlain(deck, {title, subtitle, org, team, date})` | 격자 도형 표지(사진을 쓰지 않을 때) |
| `toc(deck, {items})` | 목차. `items[i]`가 `{n,t,p}`면 장, `{t,p}`면 세부 항목. 점선 리더 + 쪽번호 |

## 표

| 호출 | 설명 |
|---|---|
| `tableNative(s, {x,y,colW,head,rows,align,boldCol,headH,rowH})` | **기본** — PowerPoint 네이티브 표. 머리글 가운데 정렬·먹 배경, 셀 윤곽선 없음, 교차 음영. 글이 길면 행이 자라며 그 높이를 미리 계산해 벽을 지킨다. 반환 `{h, bottom, grown}` |
| `table(s, {…동일…})` | 도형 표. 도식과 x좌표를 맞춰야 할 때만 |

`colW` 합계가 우측 여백을 넘거나 행의 칸 수가 열 수와 다르면 예외를 던진다.

## 도식

| 호출 | 설명 |
|---|---|
| `kpiRow(s, items, {y,h})` | KPI 스트립. `items[i] = {v, l, dark}` — `dark:true`면 먹 카드 |
| `chevrons(s, steps, {x,y,w,h})` | 체브론 프로세스. `steps[i] = {t, d, tone:'mid'\|'light'}` |
| `waterfall(s, {x,y,w,h,max,steps})` | 워터폴. `steps[i] = {t, v, base}` — 반환값은 x축 라벨 하단 y |
| `tree(s, {root, mids})` | 드라이버 트리 (루트 → 중간 → 잎, 직각 연결선) |
| `matrix(s, {x,y,w,h,quadrants,points,xLabel,yLabel})` | 2×2 매트릭스 + 버블 |
| `gantt(s, {x,y,w,cols,rows,labW,noteW,rowH})` | 간트. `rows[i] = {t, a, b, ms, tone}` (a·b는 0~1) |
| `layers(s, {x,y,w,h,items,taper})` | 3단 계층(교육체계 등) |
| `bullets(s, items, {x,y,w,h,sz})` | 불릿 목록 |
| `callout(s, t, {x,y,w,h,dark})` | 결론 밴드 |
| `pill(s, t, {x,y,w,h,dark})` | 태그 |
| `footnote(s, t, {x,y,w})` | 각주 — 오독 위험이 있을 때만 |

## 자주 쓰는 밴드

```js
B.kpiRow(s, items, { y:BT, h:1.25 });                    // 상단 KPI
B.tableNative(s, { x:M, y:BT+1.50, … });                 // 그 아래 표
B.callout(s, '결론 …', { x:M, y:BB-0.58, w:CW, h:0.58 }); // 하단 결론
```

**결론 밴드는 장(章)의 마지막 장표에만.** 본문 장표의 60%를 넘게 달면 점검기가 경고한다. 표를 그대로 읽는 문장은 결론이 아니다 — 수치가 말하지 않는 함의를 적는다.

**하단 기준선까지 채운다.** 도형·표를 늘려 6.85"에 닿게 하고, 점검기의 `빈 띠` 경고(0.45" 초과)가 나오면 늘린다.
간트·워터폴은 실제 하단 y를 반환하므로 그 값으로 다음 요소를 배치한다(`const gB = B.gantt(...)`).
