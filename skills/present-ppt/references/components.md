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
const s = deck.slide({ chapter:'1. 요약' });                       // 5존 프레임 자동
B.head(s, { title:'주제를 가리키는 명사구', message:'판단을 담은 한 문장 …함' });
await deck.save('out.pptx');
```

## 헤드

헤드 좌표는 모드마다 다르다. `deck.bt`가 본문 상단이며, 본문 요소는 이 값을 기준으로 배치한다.

| | 발표용 | 보고서용 |
|---|---|---|
| 타이틀 | 0.94" · 24pt | 0.94" · 20pt |
| 메시지 | 1.43" · 13pt | 1.37" · 12pt |
| 본문(`deck.bt`) | 2.11" | 2.00" |

타이틀 30자·메시지 2줄을 넘기면 `head()`가 몇 자를 줄여야 하는지 알려주며 멈춘다.

## 상수

| 이름 | 뜻 |
|---|---|
| `K` | 무채색 팔레트 (`K.ink` `K.g1`~`K.g7` `K.w`) |
| `M` `CW` `W` `H` | 여백 0.5 · 콘텐츠 폭 9.8333 · 슬라이드 10.8333 × 7.5 |
| `deck.bt` `BB` | 본문 상단(모드별 2.11) · 하단 6.85 (벽). `BT` 상수는 없다 |
| `Z` | 모드와 무관한 존 (`Z.chapter` `Z.rule` `Z.foot`) |
| `zones(T)` | 모드별 헤드·본문 존 (`zones(deck.T).body` 등) |
| `deck.T` | 현재 모드의 크기표 (`T.body` `T.rowH` `T.pad` …) |
| `MODE` `PALETTE` | 모드·팔레트 정의. 값을 읽을 때만 쓰고 고치지 않는다 |
| `TITLE_MAX` `SERIES_MAX` | 타이틀 30자 · 차트 계열 5개 상한 |

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
| `head(s, {title, message})` | **헤드.** 타이틀(명사구 30자·1줄) + 메시지(명사형 개조식 2줄까지). 문법과 예문은 `references/headline.md` |
| `sectionTitle(s, t, {x,y,w})` | 소제목 + 먹 구분선 — **제목 글자는 가운데 정렬이 기본**(`align:'left'`로 변경 가능) |
| `source(s, t, {x,y,w})` | 출처 — 도식·표 **바로 아래**에 붙인다. 푸터에 두지 않는다. **단위와 한 줄로 합친다**: `단위: 억 원 / 출처: …` |
| `footnote(s, t, {x,y,w})` | 각주 — **오독 위험이 있을 때만.** 한 장표에 3줄 이상이면 점검기가 경고한다 |

## 도형 안 텍스트

| 호출 | 설명 |
|---|---|
| `panel(s, {x,y,w,h,title,items,sz})` | **제목 박스.** 제목은 가운데 정렬 + 밑줄, 본문은 박스 높이에 맞춰 줄간격을 벌려 채운다. `items[i]`는 문자열 또는 `{t,sub}` |

```js
B.panel(s, { x:B.cx(0), y:2.46, w:B.cw(6), h:1.86, title:'추진 배경', items:[
  { t:'관련 법령 개정으로 신규 사무가 신설되어 차년도부터 시행', sub:'시행 시점은 관계 부처 고시에 따름' },
  '재정당국 개편으로 예산 협의 창구와 심의 일정이 함께 이동',
]});
```

제목을 왼쪽에 붙이고 글을 위로 몰면 카드 아래가 빈다. `panel()`이 이를 막는다. 글이 넘치면 필요한 높이를 알려주며 생성을 멈춘다.

## 표지·목차

| 호출 | 설명 |
|---|---|
| `cover(deck, {org, title, date, by, photos, panel})` | **사진 표지.** 사선 4분할 띠 + 먹 패널 + 대형 사진. `photos` 생략 시 자리표시. 사진은 각각 독립 도형이라 PowerPoint에서 바꿔 끼운다 |
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

## 관계 도식

모두 무채색이며 그라데이션·3D·베벨을 쓰지 않는다. 겹침과 명도, 채운 몫으로만 뜻을 만든다.
모두 `{ bottom }`을 돌려주므로 그 값으로 다음 요소의 y를 잡는다.

| 호출 | 설명 |
|---|---|
| `venn(s, {x,y,w,h,items,center})` | 벤. 영역 2~3개. 원을 반투명으로 겹쳐 교집합이 저절로 진해진다. 이름은 원 바깥에 붙는다 |
| `hubSpoke(s, {x,y,w,h,hub,items})` | 허브&스포크. 위성 3~8개 |
| `cycle(s, {x,y,w,h,items,hi,center})` | 순환. 단계 3~6개. 이웃끼리 회전 화살표로 잇는다. `hi`로 한 단계를 반전 |
| `pyramid(s, {x,y,w,h,items,notes,noteW})` | 피라미드. 2~5층, 위→아래 순서. 위로 갈수록 진해진다 |
| `steps(s, {x,y,w,h,items})` | 계단. 3~6단. `items[i] = {t, v, d}` (이름·수치·설명) |
| `harvey(s, {x,y,w,cols,rows,firstW})` | 하비볼 평가표. `rows[i] = {t, v:[0..4]}`. 범례가 자동으로 붙는다 |
| `harveyBall(s, {x,y,d,v})` | 하비볼 하나. `x,y`는 **중심**, `v`는 0~4 |
| `HARVEY_STEPS` | 하비볼 5단계의 뜻 배열 (범례를 직접 쓸 때) |
| `causeEffect(s, {x,y,w,h,rows})` | 원인→결과. 1~4행. `rows[i] = {cause, effect, note}` |
| `onTone(hex)` | 그 배경에 얹을 글자색(흰/먹)을 돌려준다 |

**피라미드와 하비볼은 `postprocess.py`가 도형의 조절점을 보정해야 제 모양이 된다.**
pptxgenjs가 사다리꼴 기울기와 부채꼴 각도를 넘기지 못하므로 도형 이름에 값을 실어 보낸다.

## 차트

무채색이므로 **계열을 명도로만 구분하고 5개까지**다. 넘으면 예외를 던진다.

| 호출 | 설명 |
|---|---|
| `barChart(s, {x,y,w,h,cats,values \| series,horizontal,stacked,max})` | 막대. `horizontal`이면 가로(순위가 위에서 아래로 정렬됨), `stacked`면 누적 |
| `lineChart(s, {x,y,w,h,cats,series,max,marker})` | 선. 추세용이라 값 축을 표시하고 데이터 레이블은 끈다 |
| `pieChart(s, {x,y,w,h,labels,values,hole,percent})` | 원. `hole`을 주면 도넛. 레이블은 조각 바깥, 범례는 오른쪽 |
| `comboChart(s, {x,y,w,h,cats,bars,lines,rightAxis,unit,unit2})` | 혼합. 막대는 왼쪽 축, 선은 오른쪽 축 |
| `chartNote(s, t, {x,y,w})` | 차트 아래 단위·출처 한 줄 |

모두 `{ bottom }`을 돌려주므로 그 값으로 다음 요소의 y를 잡는다.
**차트를 만들면 `postprocess.py`를 반드시 돌린다.**

## 자주 쓰는 밴드

```js
B.kpiRow(s, items, { y:BT, h:1.25 });                    // 상단 KPI
B.tableNative(s, { x:M, y:BT+1.50, … });                 // 그 아래 표
B.callout(s, '결론 …', { x:M, y:BB-0.58, w:CW, h:0.58 }); // 하단 결론
```

**결론 밴드는 장(章)의 마지막 장표에만.** 본문 장표의 60%를 넘게 달면 점검기가 경고한다. 표를 그대로 읽는 문장은 결론이 아니다 — 수치가 말하지 않는 함의를 적는다.

**하단 기준선까지 채운다.** 도형·표를 늘려 6.85"에 닿게 하고, 점검기의 `빈 띠` 경고(0.45" 초과)가 나오면 늘린다.
간트·워터폴은 실제 하단 y를 반환하므로 그 값으로 다음 요소를 배치한다(`const gB = B.gantt(...)`).
