# Brandlogy 좌표표 — 인치 확정값

`design-system.md`의 규칙을 13.333" × 7.5" 슬라이드 위 실제 좌표로 환산한 표. 코드는 이 값을 쓰고, 값의 의미·금지사항은 규범 사양을 따른다. 헬퍼(`assets/brandlogy.js`)가 이 값을 그대로 상수로 갖고 있으므로 **직접 좌표를 타이핑하지 말고 헬퍼를 쓴다.**

## 1. 단위 환산

| 환산 | 값 | 근거 |
|---|---|---|
| CSS px → inch | `px / 144` | 사양 §5 "Base unit: 4px (0.028\")" = 4/144. 1920×1080 기준 = 144dpi |
| CSS px → pt | `px / 2` | (px/144) × 72 |
| inch → pt | `× 72` | pptxgenjs의 `margin`·`shadow.blur`·`shadow.offset`은 pt 단위 |

폰트 크기는 사양에 pt로 직접 명시돼 있으므로 환산하지 않는다(헤드라인 32–40pt 등).

## 2. 5개 존 (전 장표 고정)

| 존 | x | y | w | h | 비고 |
|---|---|---|---|---|---|
| Header strip | 0.5 | 0.40 | 12.333 | 0.30 | 챕터명 좌측 정렬, 로고 우측 |
| — 챕터명 | 0.5 | 0.40 | 8.0 | 0.30 | Pretendard 600 12pt #8e8e93, valign middle |
| — 로고 | 11.613 | 0.44 | 1.22 | 0.24 | `13.333 − 0.5 − 1.22 = 11.613`. 종횡비 고정 |
| Headline | 0.5 | 1.00 | 12.333 | 0.75 | 700 32–40pt #222222, valign top, lineSpacingMultiple 1.20 |
| Subtitle | 0.5 | 1.63 | 12.333 | 0.40 | 500 16pt #45515e, valign top, 1.45 |
| **Body box** | 0.5 | **2.39** | 12.333 | **4.46** | 하단 경계 6.85 — 벽 |
| Clearance | — | 6.85 | — | 0.20 | **비워둔다** |
| Footer strip | 0.5 | 7.05 | 12.333 | 0.25 | 페이지번호 좌(500 10pt), 출처 우(400 9–10pt), 둘 다 #8e8e93 |

헤드라인 존(1.00–1.75)과 부제 존(1.63–2.03)이 0.12" 겹치는 것은 사양대로다 — 헤드라인 텍스트가 존을 다 채우지 않는다는 전제. 헤드라인 박스는 **위 정렬(valign top)** 로 두고, 2줄이 필요하면 32pt로 낮춰 1.75" 안에 들어오게 한다. 3줄이 되면 헤드라인이 긴 것이니 문장을 줄인다.

## 3. 12열 그리드

- 콘텐츠 폭 12.333", 거터 0.2" → **1열 = 0.8444"**, 열 간격(step) = 1.0444"
- `x(i) = 0.5 + i × 1.0444` (i = 0-based 열 인덱스)
- `span(n) = n × 0.8444 + (n−1) × 0.2`

> 사양 §5는 "column width ≈ 0.95\""라고 적었지만, 12.333" 폭에 12열·0.2" 거터를 넣으면 산술적으로 0.844"다(0.95"면 총 13.6"로 슬라이드를 넘는다). **좌표는 0.8444"로 확정한다.** 마찬가지로 Pattern B의 "5.5\" 두 열"은 12.333"를 다 쓰지 않으므로 아래 6+6열(6.066")로 확정한다. 사양의 반올림 표기와 이 표가 어긋나면 이 표가 이긴다.

자주 쓰는 분할:

| 분할 | 각 열 w | x 좌표 |
|---|---|---|
| 2-up (6+6) | 6.066 | 0.5 / 6.766 |
| 3-up (4+4+4) | 3.978 | 0.5 / 4.678 / 8.855 |
| 4-up (3×4) | 2.933 | 0.5 / 3.633 / 6.766 / 9.900 |
| 주장 5 : 근거 7 | 5.022 / 7.111 | 0.5 / 5.722 |
| 근거 7 : 주장 5 | 7.111 / 5.022 | 0.5 / 7.811 |

## 4. 본문 패턴 밴드 (모두 합이 정확히 4.46")

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
| 배지 필(30–32px) | 0.208–0.222" |
| So What 콜아웃 | 배경 #f2f3f5, radius 0.090", 패딩 0.111", 600 14pt #222222 |
| KPI 카드 최소 높이 | 1.30" (숫자 36–48pt + 라벨 11–12pt가 들어가는 최소치) |

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

pptxgenjs는 그라디언트 채우기를 지원하지 않는다. **센티넬 색 `0A0B0C`로 도형을 채운 뒤 `scripts/apply_gradient.py`로 후처리**해 벡터 `a:gradFill`로 바꾼다(이미지로 굽지 않는다 — 텍스트·확대 품질 유지).

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

## 8. 차트 기본값 (pptxgenjs)

```js
{
  chartColors: ['1456F0', '60A5FA', 'BFDBFE', '17437D'],
  showLegend: false,                 // 시리즈 1개면 범례 없음
  catAxisLabelFontFace: 'Pretendard', catAxisLabelFontSize: 10, catAxisLabelColor: '45515E',
  valAxisLabelFontFace: 'Pretendard', valAxisLabelFontSize: 10, valAxisLabelColor: '45515E',
  dataLabelFontFace: 'Pretendard', dataLabelFontSize: 11, dataLabelFontBold: true, dataLabelColor: '222222',
  showValue: true, dataLabelPosition: 'outEnd',   // 누적형은 반드시 'ctr'|'inEnd'|'inBase'
  valGridLine: { color: 'E5E7EB', size: 1 },
  catGridLine: { style: 'none' },
  valAxisLineShow: false, catAxisLineShow: false,
  chartArea: { fill: { color: 'FFFFFF' } },
}
```

시리즈 색은 **평면 브랜드 블루**만 — 데이터에 그라디언트 금지. 비교/부정 시리즈만 `EA5EC1` 또는 `8E8E93`.
