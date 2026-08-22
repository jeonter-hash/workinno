# 좌표표 — A4 가로 인치 확정값

`design-system.md`의 규칙을 **A4 가로(10.8333" × 7.5", 9906000 × 6858000 EMU, 27.52 × 19.05cm)** 슬라이드 위 실제 좌표로 환산한 표. 코드는 이 값을 쓰고, 값의 의미·금지사항은 규범 사양을 따른다. 헬퍼(`assets/brandlogy.js`)가 이 값을 그대로 상수로 갖고 있으므로 **직접 좌표를 타이핑하지 말고 헬퍼를 쓴다.**

세로 좌표(5존)는 A4 가로와 16:9의 높이가 7.5"로 같으므로 그대로다. **바뀐 것은 가로 좌표뿐이다** — 콘텐츠 폭 12.333" → 9.8333".

## 1. 단위 환산

| 환산 | 값 | 근거 |
|---|---|---|
| CSS px → inch | `px / 144` | 사양 §5 "Base unit: 4px (0.028\")" = 4/144 |
| CSS px → pt | `px / 2` | (px/144) × 72 |
| inch → pt | `× 72` | pptxgenjs의 `margin`·`shadow.blur`·`shadow.offset`은 pt 단위 |
| inch → EMU | `× 914400` | 슬라이드 크기 검증용 |

폰트 크기는 사양에 pt로 직접 명시돼 있으므로 환산하지 않는다.

## 2. 5개 존 (전 장표 고정)

| 존 | x | y | w | h | 비고 |
|---|---|---|---|---|---|
| Header strip | 0.5 | 0.40 | 9.8333 | 0.30 | 챕터명 좌측, 로고 우측 |
| — 챕터명 | 0.5 | 0.40 | 6.0 | 0.30 | 맑은 고딕 Bold 12pt #8e8e93, valign middle |
| — 로고 | 자동 | 0.44 | 자동 | 0.24 | 높이 0.24"에 맞추고 **폭은 원본 비율로 자동 계산**, 오른쪽 끝 = 10.3333" (0.5" 여백). 폭 1.9" 초과 시 폭 기준으로 축소 |
| Headline | 0.5 | 1.00 | 9.8333 | 0.75 | Bold 32pt #222222, valign top, 1.20 |
| Subtitle | 0.5 | 1.63 | 9.8333 | 0.40 | Regular 16pt #45515e, valign top, 1.45 |
| **Body box** | 0.5 | **2.39** | 9.8333 | **4.46** | 하단 경계 6.85 — 벽 |
| Clearance | — | 6.85 | — | 0.20 | **비워둔다** |
| Footer strip | 0.5 | 7.05 | 9.8333 | 0.25 | 페이지번호 좌(10pt), 출처 우(9–10pt), 둘 다 #8e8e93 |

헤드라인 존(1.00–1.75)과 부제 존(1.63–2.03)이 0.12" 겹치는 것은 사양대로다 — 헤드라인이 **한 줄**이라는 전제. A4 가로에서 32pt 한 줄은 한글 22자 안팎이 한계다. 넘으면 폰트를 줄이지 말고 **문장을 줄인다**(두 줄이 되면 1.07"를 먹어 부제 존을 침범한다).

## 3. 12열 그리드

- 콘텐츠 폭 9.8333", 거터 0.2" → **1열 = 0.63611"**, 열 간격(step) = 0.83611"
- `x(i) = 0.5 + i × 0.83611` (i = 0-based 열 인덱스)
- `span(n) = n × 0.63611 + (n−1) × 0.2`

> 사양 §5의 "column width ≈ 0.95\"", "5.5\" 두 열"은 16:9 시절 표기이고 A4 가로 폭과 맞지 않는다. **좌표는 아래 표로 확정한다.** 사양의 반올림 표기와 이 표가 어긋나면 이 표가 이긴다.

자주 쓰는 분할 (오른쪽 끝은 항상 10.3333"):

| 분할 | 각 열 w | x 좌표 |
|---|---|---|
| 2-up (6+6) | 4.8166 | 0.5 / 5.5166 |
| 3-up (4+4+4) | 3.1444 | 0.5 / 3.8444 / 7.1889 |
| 4-up (3×4) | 2.3083 | 0.5 / 3.0083 / 5.5166 / 8.0250 |
| 주장 5 : 근거 7 | 3.9805 / 5.6528 | 0.5 / 4.6805 |
| 근거 7 : 주장 5 | 5.6528 / 3.9805 | 0.5 / 6.3528 |

## 4. 본문 패턴 밴드 (세로는 16:9와 동일, 합이 정확히 4.46")

| 패턴 | 밴드 | y | h |
|---|---|---|---|
| **A** KPI Strip + Detail | KPI 행 | 2.39 | 1.60 |
| | (갭 0.24) | | |
| | 차트/2열 | 4.23 | 2.62 |
| **B** Two-Column Compare | 좌·우 열 | 2.39 | 3.66 |
| | (갭 0.20) | | |
| | So What 콜아웃 | 6.25 | 0.60 |
| | *콜아웃 생략 시* | 2.39 | 4.46 |
| **C** Diagram-Centered | 도식 영역 | 2.39 | 3.56 |
| | 캡션·요약 스트립 | 6.15 | 0.70 |
| **D** Process Flow | (상단 여백 0.30) | | |
| | 스테이지 행 | 2.69 | 1.90 |
| | 결과·인용 밴드 | 4.79 | 2.06 |
| **E** Quote + Evidence | 좌 인용(6열) | 2.39 | 4.46 |
| | 우 카드 3장(6열) | 2.39 / 3.94 / 5.49 | 각 1.36 (갭 0.19) |
| **F** Stacked Insight Layers | KPI 밴드 | 2.39 | 1.30 |
| | 차트·도식 밴드 | 3.89 | 1.86 |
| | 근거 카드 3-up | 5.95 | 0.90 |

## 5. 컴포넌트 치수

| 컴포넌트 | 값 |
|---|---|
| Standard Content Card | 흰 배경, radius 0.090" (13px), 그림자 Standard, 내부 패딩 0.111–0.167" (16–24px) |
| Featured Card | radius 0.139–0.167" (20–24px), 그림자 Brand Glow, 장표당 1개 |
| Data Card (차트 컨테이너) | 흰 배경, radius 0.090", 테두리 1px #f2f3f5 — 그림자는 없거나 Standard 하나만 |
| 카드 간 갭 | 0.111–0.167" (16–24px). 그리드 거터(0.2")를 쓰면 자동 충족 |
| 버튼(8px radius) | 0.056" |
| 필(9999px) | `rectRadius = h / 2` |
| So What 콜아웃 | 배경 #f2f3f5, radius 0.090", 패딩 0.111", Bold 14pt #222222 |
| KPI 카드 | 최소 높이 1.30". 4-up(2.31" 폭)에서 숫자는 **36pt**가 상한 — 더 키우면 잘린다. 3-up(3.14")은 40pt까지 |

## 6. 그림자 환산 (pptxgenjs `shadow`, 단위 pt)

| 토큰 | pptxgenjs |
|---|---|
| Standard | `{ type:'outer', color:'000000', opacity:0.08, blur:3, offset:2, angle:90 }` |
| Soft Glow | `{ type:'outer', color:'000000', opacity:0.08, blur:11.3, offset:0, angle:90 }` |
| Brand Glow | `{ type:'outer', color:'2C1E74', opacity:0.16, blur:7.5, offset:0, angle:90 }` |
| Brand Glow Offset | `{ type:'outer', color:'2C1E74', opacity:0.11, blur:8.8, offset:3.4, angle:17 }` |
| Elevated | `{ type:'outer', color:'242424', opacity:0.08, blur:8, offset:6, angle:90 }` |

`offset`은 음수가 되면 파일이 깨진다. 위로 드리우려면 `angle:270` + 양수 offset.

## 7. Hero Gradient (OOXML)

pptxgenjs는 그라디언트 채우기를 지원하지 않는다. **센티넬 색 `0A0B0C`로 도형을 채운 뒤 `scripts/postprocess.py`로 후처리**해 벡터 `a:gradFill`로 바꾼다(이미지로 굽지 않는다 — 텍스트·확대 품질 유지).

```xml
<a:gradFill rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="1456F0"/></a:gs>
    <a:gs pos="50000"><a:srgbClr val="3B82F6"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="60A5FA"/></a:gs>
  </a:gsLst>
  <a:lin ang="2700000" scaled="0"/>
</a:gradFill>
```

`ang="2700000"` = 45°(시계방향, 3시 기준) = 좌상 → 우하. CSS `135deg`와 같은 방향이다. 각도·정지점·색을 바꾸지 않는다.

## 8. 한글 폰트 고정 (동아시아 런)

pptxgenjs는 `<a:latin>`만 쓰고 `<a:ea>`(동아시아)를 비워 둔다. 그대로 두면 한글이 테마 기본 폰트로 떨어질 수 있으므로 `scripts/postprocess.py`가 모든 런에 `<a:ea typeface="맑은 고딕"/>`·`<a:cs>`를 주입하고 테마(major/minor)의 동아시아 폰트도 맑은 고딕으로 바꾼다. **후처리를 건너뛰면 한글 서체가 보장되지 않는다.**

## 9. 차트 기본값 (pptxgenjs)

```js
{
  chartColors: ['1456F0', '60A5FA', 'BFDBFE', '17437D'],
  showLegend: false,                 // 시리즈 1개면 범례 없음
  catAxisLabelFontFace: '맑은 고딕', catAxisLabelFontSize: 10, catAxisLabelColor: '45515E',
  valAxisLabelFontFace: '맑은 고딕', valAxisLabelFontSize: 10, valAxisLabelColor: '45515E',
  dataLabelFontFace: '맑은 고딕', dataLabelFontSize: 11, dataLabelFontBold: true, dataLabelColor: '222222',
  showValue: true, dataLabelPosition: 'outEnd',   // 누적형은 반드시 'ctr'|'inEnd'|'inBase'
  valGridLine: { color: 'E5E7EB', size: 1 },
  catGridLine: { style: 'none' },
  valAxisLineShow: false, catAxisLineShow: false,
  chartArea: { fill: { color: 'FFFFFF' } },
}
```

시리즈 색은 **평면 브랜드 블루**만 — 데이터에 그라디언트 금지. 비교/부정 시리즈만 `EA5EC1` 또는 `8E8E93`.
