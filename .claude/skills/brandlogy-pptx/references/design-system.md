# PPT Design System (A4 가로) — 규범 사양

이 문서는 사용자가 제공한 디자인 시스템 원문에 **판형(A4 가로) · 서체(맑은 고딕) · 로고(한국표준협회)** 세 축의 변경을 반영한 것이다. **값(수치·색·규칙)에 대한 최종 권한은 이 문서에 있다.** 색 체계·그림자·Hero Gradient·레이아웃 문법은 원안 그대로다.
좌표를 인치 단위로 확정한 계산표는 `layout-geometry.md`, 실행 지침은 `../SKILL.md`를 본다.

---

## 0. Production Constraints (Read First)

**Output**
- A4 가로(landscape) only — PowerPoint A4 가로 = 10.8333" × 7.5" (9906000 × 6858000 EMU, 27.52 × 19.05cm), reference resolution 1560 × 1080 px @144dpi. No other page size or aspect ratio is a valid output.
- 표준 템플릿(.pptx)이 주어지면 그것을 경유하고, 없으면 `assets/brandlogy.js` 헬퍼로 생성한다.

**Brand Assets (mandatory, no substitution)**
- **Logo**: 한국표준협회(KSA) 로고만 사용한다. MiniMax·Brandlogy를 비롯한 어떤 제3자 로고도 대체 사용 금지. Place at top-right of every slide unless explicitly told otherwise. Use the logo file exactly as provided — see Logo Integrity Rule below. 로고 원본은 `assets/ksa_logo.png`(누끼 PNG)에 두고, 없으면 사용자에게 요청한다 — 직접 그리거나 재현하지 않는다.
- **Typography**: **맑은 고딕(Malgun Gothic) only** — no exceptions. Do NOT use Pretendard, DM Sans, Outfit, Poppins, Roboto, Noto, HY계열, or any other family under any circumstance. 맑은 고딕은 Semilight / Regular / Bold 세 단계뿐이므로, 아래 표의 6단 웨이트는 §3의 매핑표대로 접어서 쓰고 접혀 사라진 대비는 **크기와 색**으로 만든다. 맑은 고딕은 Windows 기본 탑재 서체라 별도 임베드가 필요 없다(대신 macOS·리눅스에는 없으므로 그 환경에 배포할 때는 PDF로 낸다).

### Slide Skeleton — locked positions across the deck

Every slide in the deck must place these five zones at identical coordinates. The reader's eye should never have to relearn the layout when flipping pages — only the body contents change, never the frame.

| Zone | Position (slide 10.8333" × 7.5") | Contents | Style |
|---|---|---|---|
| Header strip | 0.4"–0.7" from top, full width within 0.5" side margins | Chapter name (left), 한국표준협회 로고 (right) | Chapter: 맑은 고딕 Bold, 12pt, #8e8e93. Logo: original transparent PNG provided by user, 높이 0.24"에 맞추고 폭은 **원본 비율로 자동 계산**(늘리지 않는다), 오른쪽 끝을 슬라이드 우측 0.5" 여백에 정렬 |
| Headline zone | 1.0"–1.75" from top, 0.5" left margin | Slide headline (one-sentence copy, 대제목) | 맑은 고딕 Bold, 32pt, #222222, line-height 1.20 |
| Subtitle zone | 1.63"–2.03" from top, 0.5" left margin | Subtitle (부제목) | 맑은 고딕 Regular, 16pt, #45515e, line-height 1.45 |
| Body box | 2.39"–6.85" from top, 0.5" side margins | All body components — see §5 | Mixed |
| Footer strip | 7.05"–7.3" from top, full width within 0.5" side margins | Page number (left), source/footnote (right) | Page: 맑은 고딕 Regular, 10pt, #8e8e93. Source: 맑은 고딕 Regular, 9–10pt, #8e8e93 |

**Vertical rhythm**: The gaps in the upper half of the slide are intentional and not uniform — Header strip → Headline = 0.3" (deliberate breathing space below chapter line). Headline zone → Subtitle zone = 0.1" zone-to-zone, but because the headline text typically renders shorter than its zone, the visual gap between the headline text bottom and the subtitle text is closer to 0.13" — title and subtitle read as one tightly-coupled unit. Subtitle bottom → Body top ≈ 0.36" — a clear visual break that lets the body box read as its own region while still feeling anchored to the title block above. This pattern (loose top, tight middle, medium bottom) is what makes the title block feel like the slide's "anchor" rather than a floating header.

**Lock rule**: These five zones do not move between slides. Chapter name stays at the same baseline, headline starts at the same Y, subtitle starts at the same Y, body box starts and ends at the same Y, footer strip is identical. Override is permitted only when structurally unavoidable — i.e., a section divider that intentionally breaks the frame, a full-bleed cover, or a closing slide. Routine "this body is taller than usual" is NOT an unavoidable case; restructure the content instead.

**Hard boundary**: Body content lives strictly inside 2.39"–6.85" (the Body box). It does NOT bleed upward into the Subtitle zone (above 2.39") and does NOT bleed downward into the Footer strip (below 6.85"). A 0.2" clearance buffer (6.85"–7.05") sits between the body box and the footer strip — keep it empty so the page number and source line never get crowded. Anything taller than 4.46" of body height must be split, scaled down, or moved to a second slide — not allowed to invade adjacent zones.

**Logo Integrity Rule**: The 한국표준협회 로고 must be placed exactly as provided — original file, original proportions, original colors, original transparency (alpha channel preserved). The user supplies a transparent PNG (누끼); insert that file as-is. Do NOT add an underline, strikethrough, drop shadow, glow, border, frame, recolor, gradient, opacity change, background fill, opaque box behind the logo, or any other visual treatment. Do NOT crop, stretch, skew, rotate, or duplicate the logo. The only permitted operations are uniform scaling (preserving aspect ratio) to fit the 0.24" target height (resulting width follows the source proportions — 폭을 임의 값으로 강제하지 않는다), and uniform color inversion to a white variant when placed on dark backgrounds (section dividers, closing slides). A black or white rectangle behind the logo is a defect, not the design — if a generated output shows any line, mark, decoration, or solid box on/behind the logo that is not in the original file, treat it as a defect and fix it before exporting.

### Body Density Rule

The lower body box must NOT be left half-empty. Plan body content to fill the available area at a comfortable reading density — charts, diagrams, KPI tiles, comparison tables, dual-column layouts, supporting captions. Empty bottom space breaks the McKinsey/BCG sharpness target. Whitespace is a tool for breathing rhythm between elements, not a default for the bottom 30% of every slide.

Density never overrides the hard boundary. "Filling the bottom" means filling **inside** the body box (2.39"–6.85") — it does NOT mean spilling content into the 0.2" footer clearance, the footer strip, or the subtitle/headline zones above. If pursuing density tempts you to push a card down to 7.0" or up to 2.35", you have too much content on the slide; split it.

If a slide genuinely has thin content, use one of these density tactics — never decorative padding, never zone invasion:
- Pull supporting evidence (quote, data point, mini-chart, source) into a side panel, still inside the body box
- Add a "So What" callout box at the bottom of the body box (above 6.85") summarizing the takeaway
- Insert a diagram that visually reinforces the headline
- Split the body into a 2-column claim / evidence layout
- Use Pattern F (Stacked Insight Layers, see §5) — three horizontal bands within the body box

Do NOT pad with decorative shapes or stock illustrations to fake density.

### Visualization-First Rule (Style 1 — strong)

Style 1 is data-first. Whenever a slide carries data, comparison, process, structure, or relationship — **visualize it, do not narrate it in prose**. This is a strong default, not a suggestion.

**Trigger conditions** (if any of these apply, the slide MUST include a visualization):
- Two or more numbers being compared (chart or KPI tile row, never inline prose)
- A trend over time (line chart or timeline, even with only 2–3 points)
- Composition / share / distribution (bar, donut, or 100% stacked bar)
- A process or sequence (horizontal arrow flow, numbered stages)
- A comparison across categories (grouped/stacked bar or table — chart preferred over table)
- A structural relationship between concepts (diagram, matrix, 2×2)
- A geographic or hierarchical breakdown (map, tree, or org chart)

**Visualization options to reach for, in priority order:**
1. **Charts** — bar (horizontal/vertical/grouped/stacked), line, area, scatter, donut. Default. Use Style 1 chart palette (§4).
2. **KPI tiles with sparklines** — when a single number deserves emphasis but context still matters
3. **Diagrams** — flow, sequence, 2×2 matrix, layered architecture, Venn (rare), funnel, hierarchy
4. **Annotated images / screenshots** — only when the visual artifact itself is the evidence
5. **Tables** — last resort, only when individual cell values matter and ranking/comparison is secondary

**Constraints** (visualization never breaks the layout):
- Visualizations live strictly inside the body box (2.39"–6.85"). Never bleed into headline/subtitle/footer zones.
- A single slide should carry 1–2 visualizations max, not 4+. Cramming charts breaks the pacing more than missing them.
- Every chart and diagram must have: a title (맑은 고딕 Bold 14pt), axis labels (맑은 고딕 Regular 10pt #45515e), and a source line (맑은 고딕 Regular 9pt #8e8e93) directly below.
- If a visualization would force font sizes below 9pt or compress data labels into illegibility, the slide has too much data — split it, don't shrink the chart.
- Pure-prose body slides are reserved for: section openers, hero takeaways, single-quote callouts, definitions. Everything else gets a visualization.

When in doubt, ask: "Could this be a chart instead of bullets?" If yes, make it a chart.

---

## 1. Visual Theme & Atmosphere

The aesthetic bridges Apple-grade product-marketing clarity with a playful, rounded, gallery-like feel. Pure white (#ffffff) is the structural background; color enters via charts, KPI cards, gradients, and accent elements. 맑은 고딕(Regular/Bold) carries a confident-but-approachable tone — not aggressive, not airy.

**Key Characteristics**
- White-dominant canvas with colorful accent elements (charts, KPI cards, gradients) carrying visual interest
- 맑은 고딕 across the entire system — 서체를 늘리지 않고 굵기·크기·색으로만 위계를 만든다
- Pill buttons (9999px / fully rounded) for nav, tabs, toggles
- Generous rounded cards (16–24px radius) for content blocks
- Brand blue spectrum: #1456f0 → #3b82f6 → #60a5fa
- Brand pink (#ea5ec1) reserved for decorative accents
- Near-black text (#222222, #18181b) on white
- Purple-tinted shadows (rgba(44, 30, 116, 0.16)) for featured cards — subtle brand glow
- Dark sections (#181e25) for divider/closing slides if needed

---

## 2. Color Palette & Roles

### Brand Primary
- **Brand Blue (#1456f0)**: primary brand identity color
- **Sky Blue (#3daeff)**: lighter brand variant for accents
- **Brand Pink (#ea5ec1)**: secondary accent — decorative only, never on body text

### Blue Scale
- `#bfdbfe` — light blue background
- `#60a5fa` — primary-light, active states, chart fills
- `#3b82f6` — primary-500, standard blue actions, primary chart series
- `#2563eb` — primary-600, emphasis
- `#1d4ed8` — primary-700, deep emphasis
- `#17437d` — brand-deep

### Text
- `#222222` — primary text (body, headline)
- `#18181b` — heading-dark, dark button text
- `#181e25` — dark surface text, footer-bg
- `#45515e` — secondary text (subtitle, captions)
- `#8e8e93` — tertiary/muted text (chapter name, page number, source)
- `#5f5f5f` — helper text

### Surface
- `#ffffff` — primary background (every slide)
- `#f0f0f0` — secondary container background
- `hsla(0, 0%, 100%, 0.4)` — frosted glass overlay
- `#f2f3f5` — subtle dividers
- `#e5e7eb` — component borders

### Semantic
- `#e8ffea` — success background (pair with `#16a34a` for success text)

### Shadow Library
| Token | Value | Use |
|---|---|---|
| Standard | rgba(0, 0, 0, 0.08) 0px 4px 6px | Default cards |
| Soft Glow | rgba(0, 0, 0, 0.08) 0px 0px 22.576px | Ambient shadow |
| Brand Glow | rgba(44, 30, 116, 0.16) 0px 0px 15px | Featured cards |
| Brand Glow Offset | rgba(44, 30, 116, 0.11) 6.5px 2px 17.5px | Hero product cards |
| Elevated | rgba(36, 36, 36, 0.08) 0px 12px 16px -4px | Lifted/hover-equivalent emphasis |

### Hero Gradient (Premium Accent — use sparingly)

A single elegant blue gradient is permitted to elevate hero moments. The gradient is built from the existing blue scale — no new colors are introduced.

**Token — Hero Gradient**: `linear-gradient(135deg, #1456f0 0%, #3b82f6 50%, #60a5fa 100%)`

**Fixed parameters — do NOT vary these:**
- Angle: 135° (top-left dark → bottom-right light). Consistent across the deck.
- Stops: 0% / 50% / 100% with the three brand blues in order. No additional color stops.
- Colors: Only #1456f0, #3b82f6, #60a5fa from the existing scale. No purple, cyan, teal, or pink mixed in.

**Permitted locations (max 3 across the entire deck)**
1. **Cover slide hero card** — the single featured card on the cover slide (Pattern E in §5)
2. **Section divider background** — the dark divider slide may use Hero Gradient instead of solid #181e25 for a more premium feel
3. **One Featured KPI card per slide** — at most one KPI tile per slide may use the gradient as background, with white text on it. Pair with Brand Glow shadow.

**Forbidden locations**
- Chart bars / lines / data points — gradients on data create false visual hierarchy (longer bars look "more saturated"). Use flat brand blue (#1456f0 or #3b82f6) for all chart series.
- Headline or body text — text-on-gradient or gradient-text is AI-slop visual. Text stays solid #222222.
- Header strip and footer strip — these zones are flat, always.
- Standard content cards (non-featured) — gradients on every card destroy the white-canvas brand identity.
- Body card backgrounds in bulk — only one gradient element per slide. Multiple gradients break the design.

**Premium-look rules**
- Always pair Hero Gradient with the **Brand Glow shadow** (rgba(44, 30, 116, 0.16) 0px 0px 15px) for the soft halo effect that prevents flat-poster look.
- Inside text on gradient must be **white (#ffffff)** at 맑은 고딕 Regular/Bold, never #222222 or any blue.
- Gradient cards use **20–24px radius** (the larger end of the radius scale) — sharp corners on gradient look cheap.
- Never overlay another gradient, image, or pattern on top of Hero Gradient. The gradient itself IS the visual interest.

If the gradient would compete with charts on the same slide, the chart wins — move the gradient element to a different slide. Charts and gradients on the same slide create visual chaos.

---

## 3. Typography Rules (맑은 고딕 전용)

### Family
**맑은 고딕(Malgun Gothic) only** (한글 + Latin). 다른 어떤 패밀리도 허용하지 않는다 — 헤드라인도, 데이터도, 대체 서체도. 맑은 고딕은 Windows 기본 탑재 서체이므로 .pptx 임베드가 필요 없다. 맑은 고딕이 없는 환경(macOS·리눅스)에 배포할 때는 **서체를 바꾸지 말고 PDF로 낸다.**

### Weight Map — 6단 웨이트를 3단으로 접는다
맑은 고딕에는 Semilight / Regular / Bold 세 단계뿐이다. 원래 시스템의 웨이트 역할을 아래처럼 접고, **접혀서 사라진 대비(500 vs 600)는 크기와 색으로 만든다.**

| 원 웨이트 | 맑은 고딕 | 역할 |
|---|---|---|
| 800 / 700 | **Bold** | 슬라이드 헤드라인, 섹션 제목, KPI 숫자, 본문 강조 |
| 600 | **Bold** | 카드 제목, 버튼 텍스트, 챕터명, 본문 H2/H3 — 700과 굵기가 같으므로 **크기(14–20pt)와 색으로 구분** |
| 500 | Regular | 부제, 소제목, 라벨, 강조 본문 — 400과 굵기가 같으므로 **색(#45515e vs #222222)으로 구분** |
| 400 | Regular | 본문, 캡션, 출처, 각주 |
| 300 | Semilight | 특수 용도(옅은 보조 텍스트). Semilight가 없는 PC를 고려해 남발하지 않는다 |

### Hierarchy (A4 가로 슬라이드, 1560 × 1080 @144dpi 기준)

| Role | Weight | Size (pt) | Line Height | Color |
|---|---|---|---|---|
| Slide Headline (대제목) | Bold | 32pt | 1.20 | #222222 |
| Subtitle (부제목) | Regular | 16pt | 1.45 | #45515e |
| Body H2 (본문 중제목) | Bold | 18pt | 1.40 | #222222 |
| Body H3 (본문 소제목) | Bold | 14–16pt | 1.45 | #222222 |
| Body | Regular | 12–13pt | 1.50 | #222222 |
| Body Emphasized | Regular | 12–13pt | 1.50 | #222222 (또는 #1456f0) |
| Body Bold | Bold | 12–13pt | 1.50 | #222222 |
| KPI Number | Bold | 36pt (4-up) / 40pt (3-up) | 1.10 | #1456f0 (또는 문맥색) |
| KPI Label | Regular | 11–12pt | 1.30 | #45515e |
| Chapter Name | Bold | 12pt | 1.30 | #8e8e93 |
| Page Number | Regular | 10pt | 1.30 | #8e8e93 |
| Caption / Source | Regular | 9–10pt | 1.40 | #8e8e93 |
| Tag / Badge | Bold | 10–11pt | 1.20 | varies |

### Principles
- **굵기 2단(Regular/Bold) + 크기 + 색**이 위계를 만든다. 서체를 늘려 해결하지 않는다.
- A4 가로 본문 폭은 9.8333"다. **헤드라인은 32pt 한 줄, 한글 22자 내외**가 한계 — 넘으면 크기를 줄이지 말고 문장을 줄인다(두 줄이 되면 부제 존을 침범한다).
- 기본 행간 1.50(본문) / 1.45(부제) / 1.20–1.30(헤드라인·라벨) / 1.10(큰 숫자).
- 인용 외에는 이탤릭을 쓰지 않는다 — 굵기·색 대비로 대신한다.
- 한글·영문 혼용: 맑은 고딕이 둘 다 처리한다. 문장 중간에 서체를 바꾸지 않는다.
- 자간: 본문 0, 32pt 헤드라인 −0.65pt(≈ −0.02em).

## 4. Component Stylings

### Buttons / Pills

**Pill Primary Dark**
- BG #181e25, text #ffffff, padding 11px 20px, radius 8px, 맑은 고딕 Bold 13–14pt
- Use: primary CTA on cover/closing slides

**Pill Nav / Tab**
- BG rgba(0, 0, 0, 0.05), text #18181b, radius 9999px, 맑은 고딕 Regular 11–12pt
- Use: section tabs, filter indicators

**Pill White**
- BG #ffffff, text rgba(24, 30, 37, 0.8), radius 9999px, 맑은 고딕 Regular
- Use: secondary nav, inactive tabs

**Secondary Light**
- BG #f0f0f0, text #333333, padding 11px 20px, radius 8px, 맑은 고딕 Regular
- Use: secondary actions, divider tags

### Content Cards (body zone)

**Standard Content Card**
- BG #ffffff, radius 13–16px, shadow Standard (rgba(0,0,0,0.08) 0px 4px 6px)
- Internal padding 16–24px
- Use: KPI tiles, point-by-point breakdowns, capability cards

**Featured Card**
- BG vibrant gradient (blue/purple/pink/orange family) or white, radius 20–24px
- Shadow Brand Glow (rgba(44,30,116,0.16) 0px 0px 15px)
- Use: hero takeaway, section opener, headline product card

**Data Card (chart container)**
- BG #ffffff, radius 13px, border 1px solid #f2f3f5, no shadow OR Standard shadow
- Title row at top (맑은 고딕 Bold, 14pt), source line at bottom (맑은 고딕 Regular, 9pt, #8e8e93)

### Charts
- Primary series: #1456f0 or #3b82f6
- Secondary series: #60a5fa, #bfdbfe, #17437d
- Negative/comparison series: #ea5ec1 or neutral #8e8e93
- Gridlines: #e5e7eb, 1px
- Axis labels: 맑은 고딕 Regular, 10pt, #45515e
- Data labels on bars/points: 맑은 고딕 Bold, 11pt, #222222
- Always cite source under chart in 9–10pt #8e8e93

### Tables
- Header row: BG #f2f3f5, 맑은 고딕 Bold 12pt, #222222
- Body rows: 맑은 고딕 Regular 12pt, #222222, alternating BG #ffffff / #fafafa optional
- Row dividers: 1px #e5e7eb
- Cell padding: 8px 12px
- No vertical dividers — rely on column spacing

### Links / Inline Emphasis
- Primary inline: #1456f0, no underline, 맑은 고딕 Regular
- Source attribution: #8e8e93, 맑은 고딕 Regular, 9–10pt

---

## 5. Layout Principles

### Slide Grid (A4 가로, 10.8333" × 7.5")
- Outer margins: 0.5" left/right, 0.4" top, 0.3" bottom
- Content width: 9.8333"
- Content height: 6.8"
- Internal column system: 12-column grid, 0.2" gutter (column width = 0.63611" — `layout-geometry.md` §3 확정값)

### Vertical Zones (locked — identical coordinates on every slide)
| Zone | Y-range (from top) | Contents |
|---|---|---|
| Header strip | 0.4" – 0.7" | Chapter name (left), 한국표준협회 로고 (right) |
| Headline | 1.0" – 1.75" | Slide headline (대제목) |
| Subtitle | 1.63" – 2.03" | Subtitle (부제목, one-sentence lead, 16pt) |
| Body box | 2.39" – 6.85" | All body components, charts, diagrams |
| Clearance buffer | 6.85" – 7.05" | Empty — no content, no padding |
| Footer strip | 7.05" – 7.3" | Page number (left), source line (right) |

**Lock rule**: These zones do not shift between slides. The header → headline → subtitle → body → footer rhythm is fixed across the entire deck. Body zone (2.39"–6.85", a 4.46" tall box) is where 95% of design work happens, and it must be filled densely — but only inside the box (see §0 Body Density Rule and Hard boundary).

**Vertical rhythm (gaps between zones)**: Header → Headline = 0.3" (loose, breathing space below chapter line). Headline zone → Subtitle zone = 0.1" zone-to-zone, but visual gap between rendered headline text and subtitle text is closer to 0.13" because the headline text doesn't fill its full zone — title and subtitle read as a single tightly-coupled unit. Subtitle → Body = 0.36" (medium, lets the body box read as its own region while staying anchored to the title block). The non-uniform rhythm — loose top, tight middle, medium bottom — is what makes the title block feel like an anchor rather than a floating header.

**Clearance buffer**: The 0.2" gap between body box bottom (6.85") and footer strip top (7.05") is intentional. It must remain empty so the page number and source line never visually collide with body cards. Treat 6.85" as a wall.

**Override exception**: Section dividers, full-bleed covers, and closing slides may break the frame intentionally — but only if the break is the design point. "I had too much content" is not a valid override.

### Spacing Scale (within body zone)
- Base unit: 4px (0.028" / 0.07cm)
- Steps: 4, 8, 12, 16, 20, 24, 32, 40, 48, 64, 80px
- Card-to-card gap: 16–24px
- Section-internal padding: 16–24px
- Headline-to-subtitle gap: 16px
- Subtitle-to-body gap: 32px

### Border Radius Scale
- 4px — small tags, micro badges
- 8px — buttons, small cards, input-like elements
- 11–13px — medium cards, data tiles
- 16–20px — large content cards
- 22–24px — hero product cards, major containers
- 30–32px — badge pills
- 9999px — full pill (buttons, tabs)

### Body Composition Patterns (use these to maintain density)

**Pattern A — KPI Strip + Detail** (most common)
- Top half of body: 3–4 KPI cards in a row (each ~3" wide × 1.6" tall)
- Bottom half: supporting chart or 2-column claim/evidence

**Pattern B — Two-Column Compare**
- Left column (4.8166" wide): claim + supporting bullets
- Right column (4.8166" wide): chart, diagram, or visual evidence
- (주장 5열 3.9805" : 근거 7열 5.6528" 비대칭 분할도 가능 — `layout-geometry.md` §3)
- Optional bottom-spanning "So What" callout box

**Pattern C — Diagram-Centered**
- Centered diagram occupies ~70% of body
- 3–4 caption boxes around the diagram explain components
- Bottom strip: source + summary takeaway

**Pattern D — Process Flow**
- Horizontal arrow flow with 4–6 stages across body
- Each stage: numbered circle, stage label, 1–2 line description
- Below the flow: outcomes summary or pull-quote

**Pattern E — Quote + Evidence**
- Large pull-quote (맑은 고딕 Regular, 24–28pt) on left half
- Stack of 2–3 supporting data cards on right half

**Pattern F — Stacked Insight Layers** (use when content is thin to keep density)
- Top band: KPI summary (1 row)
- Middle band: one chart or diagram
- Bottom band: 3-up evidence cards (claim + 1-line proof + source)
- Eliminates empty bottom space without padding

---

## 6. Depth & Elevation

| Level | Treatment | Use |
|---|---|---|
| 0 — Flat | No shadow | Background, in-flow text |
| 1 — Subtle | rgba(0,0,0,0.08) 0px 4px 6px | Standard content cards |
| 2 — Ambient | rgba(0,0,0,0.08) 0px 0px 22.576px | Soft surrounding glow |
| 3 — Brand Glow | rgba(44,30,116,0.16) 0px 0px 15px | Featured/takeaway cards |
| 4 — Elevated | rgba(36,36,36,0.08) 0px 12px 16px -4px | Hero, hover-equivalent emphasis |

Use Brand Glow sparingly — at most one element per slide. Standard shadow handles most cards. Flat is the default for blocks of text directly on the white slide background.

---

## 7. Do's and Don'ts

### Do
- Anchor chapter name, headline, and subtitle at the same coordinates on every single slide
- Fill the lower body box with structured, dense content (charts, KPI cards, 2-column layouts, evidence stacks)
- 맑은 고딕의 Bold/Regular와 크기·색으로 위계를 만든다 — 서체를 늘리지 않는다
- Apply pill radius (9999px) for tabs/toggles, 8px for action buttons, 16–24px for content cards
- Reserve the brand purple-tinted shadow for the single featured element on a slide
- 본문은 맑은 고딕 Regular; Bold는 강조와 KPI 숫자에만
- Cite every data source in 9–10pt #8e8e93 at the bottom of the relevant element
- Use 12-column internal grid logic for body layouts

### Don't
- Don't leave the bottom 20–30% of the body zone visually empty — restructure or add evidence/callout
- Don't use any font other than 맑은 고딕 — Pretendard·DM Sans·Outfit·Poppins·Roboto·Noto·HY계열 모두 금지
- Don't use the MiniMax·Brandlogy logo or any logo other than 한국표준협회(KSA)
- Don't deviate from the locked zone coordinates (header strip, headline, subtitle, body box, footer strip) across slides — same Y for chapter, headline, subtitle, body top, body bottom, logo, and source line on every page. Override only when structurally unavoidable (section divider, full-bleed cover, closing slide).
- Don't let body content invade the headline/subtitle zones above 2.39" or the clearance buffer / footer strip below 6.85" — if it doesn't fit, split the slide
- Don't apply brand pink (#ea5ec1) to body text or buttons — decorative accents only
- Don't use sharp corners on content cards — minimum radius is 8px, body cards 13–24px
- Don't darken shadows past 0.16 opacity — light-and-airy is the brand register
- Don't apply Hero Gradient to chart bars, lines, or any data series — gradients create false visual hierarchy on data. Charts use flat brand blue.
- Don't apply Hero Gradient to text or use gradient-text effects — solid #222222 for ink, white for text on gradient surfaces.
- Don't use more than one Hero Gradient element per slide, and don't exceed 3 gradient elements across the entire deck.
- Don't vary the gradient angle, stops, or colors — Hero Gradient is fixed at `linear-gradient(135deg, #1456f0 0%, #3b82f6 50%, #60a5fa 100%)`.
- Don't pad slides with decorative shapes or stock illustrations to fake density — use real evidence
- Don't introduce a second display family alongside 맑은 고딕
- Don't use weight 800–900 for body headings (reserve for closing slide / section divider only)
- Don't use emojis anywhere on slides

---

## 8. Aspect Ratio & Export Notes

- **A4 가로만.** 16:9·4:3·1:1·9:16·A4 세로·letter 등 다른 판형 요청은 거절하고 A4 가로로 제안한다.
- Export resolution target: 1560 × 1080 px minimum for image preview; native PowerPoint vector preserved in the .pptx.
- **Font**: 맑은 고딕은 Windows 기본 서체라 임베드가 필요 없다. 맑은 고딕이 없는 환경(macOS·리눅스)에 배포할 때는 .pptx 대신 PDF로 낸다 — 다른 서체로 대체하지 않는다.
- All chart text and data labels must be **live text** (not rasterized images) to keep edit-ability.

---

## 9. Agent Prompt Guide

### Quick Reference Strip
- 판형: A4 가로(10.8333" × 7.5")만
- Bg: #ffffff (every slide), #181e25 (closing/divider only)
- Headline: #222222, 맑은 고딕 Bold, 32pt
- Subtitle: #45515e, 맑은 고딕 Regular, 16pt
- Body: #222222, 맑은 고딕 Regular, 12–14pt
- Source/caption: #8e8e93, 맑은 고딕 Regular, 9–10pt
- Brand blue: #1456f0 / #3b82f6 / #60a5fa
- Hero Gradient (premium accent, max 3 elements per deck): `linear-gradient(135deg, #1456f0 0%, #3b82f6 50%, #60a5fa 100%)` — for cover hero card / section divider bg / 1 featured KPI per slide. Pair with Brand Glow shadow. Never on charts, text, headers, or footers.
- Brand pink (accents only): #ea5ec1
- Borders: #e5e7eb, #f2f3f5
- Logo: 한국표준협회(KSA), top-right — 오른쪽 끝이 우측 0.5" 여백, y≈0.44", 높이 0.24"에 원본 비율 유지, alpha 보존 (배경 채움 / 밑줄 / 그림자 / 색 보정 / 크롭 금지)
- Page number: bottom-left, 맑은 고딕 Regular 10pt #8e8e93
- Font: 맑은 고딕 only

### Example Component Prompts

**Cover Slide (Hero Gradient option)**
> "Build a A4 가로 표지 장표 on #ffffff. Slide headline at 1.0"–1.75" from top, 맑은 고딕 Bold, 32pt, #222222, line-height 1.20. Subtitle at 1.63"–2.03", 맑은 고딕 Regular, 16pt, #45515e, line-height 1.45. Body zone (2.39"–6.85"): single hero featured card with Hero Gradient background (linear-gradient(135deg, #1456f0 0%, #3b82f6 50%, #60a5fa 100%)), 24px radius, Brand Glow shadow rgba(44,30,116,0.16) 0px 0px 15px, containing the deck's central KPI in 맑은 고딕 Bold 44pt #ffffff (white text on gradient — never blue) with a 12pt 맑은 고딕 Regular rgba(255,255,255,0.85) label below. 한국표준협회 로고 at top-right (insert provided PNG file as-is, 높이 0.24"·원본 비율 유지, alpha preserved — no background fill, no decorations, no recolor), page number at bottom-left. Body content stays strictly above 6.85" — clearance buffer 6.85"–7.05" remains empty."

**Content Slide — KPI Strip + Chart**
> "Build a A4 가로 본문 장표. Chapter name top-left at y=0.4" baseline, 맑은 고딕 Bold 12pt #8e8e93. 한국표준협회 로고 top-right at y≈0.44", insert provided PNG file as-is (높이 0.24"·원본 비율 유지, transparent alpha preserved, no background fill / underline / box / recolor). Headline at 1.0"–1.75", 맑은 고딕 Bold 32pt #222222. Subtitle at 1.63"–2.03", 맑은 고딕 Regular 16pt #45515e. Body zone split: top half (2.39"–3.99") is a row of 4 KPI cards (white BG, 13px radius, Standard shadow, internal padding 20px), each with KPI number 맑은 고딕 Bold 36pt #1456f0 and label 맑은 고딕 Regular 11pt #45515e. Bottom half (4.23"–6.85") is a horizontal bar chart, primary series #3b82f6, axis labels 맑은 고딕 Regular 10pt #45515e, source line under chart 맑은 고딕 Regular 9pt #8e8e93. Page number at bottom-left (y=7.05"), source/footnote at bottom-right (y=7.05")."

**Two-Column Compare**
> "Build a A4 가로 장표(앵커는 표준). Body zone: two columns, 4.8166" wide each, 0.2" gutter (x = 0.5" / 5.5166"). Left column header 맑은 고딕 Bold 18pt #222222, body bullets 맑은 고딕 Regular 13pt #222222 line-height 1.50. Right column same structure but with a vertical bar chart (primary #1456f0, comparison #ea5ec1). Add a 'So What' callout box spanning full width at the bottom of the body zone, BG #f2f3f5, 13px radius, padding 16px, 맑은 고딕 Bold 14pt #222222."

**Section Divider**
> "Build a A4 가로 섹션 디바이더 on #181e25 (dark) BG OR Hero Gradient linear-gradient(135deg, #1456f0 0%, #3b82f6 50%, #60a5fa 100%) for premium feel. Section number top-left in 맑은 고딕 Bold 14pt rgba(255,255,255,0.6). 한국표준협회 로고 top-right in white variant (original asset uniformly inverted to white — no other modification). Section title centered vertically, 맑은 고딕 Bold 48pt #ffffff. One-line lead under title, 맑은 고딕 Regular 20pt rgba(255,255,255,0.7), line-height 1.45. Page number bottom-left in rgba(255,255,255,0.6)."

### Iteration Checklist (run before exporting any slide)
1. 판형이 A4 가로(10.8333" × 7.5")인가?
2. 전부 맑은 고딕인가 — 다른 서체가 섞이지 않았는가?
3. 한국표준협회 로고가 우상단(오른쪽 끝 0.5" 여백)에 원본 비율·투명 배경 그대로인가 — 뒤에 흰/검은 박스, 밑줄, 그림자, 색 보정, 크롭, 회전이 없는가?
4. All five zone anchors (header / headline / subtitle / body box / footer) match previous slide coordinates?
5. Body content stays strictly inside 2.39"–6.85" — no invasion of headline/subtitle zones above or clearance buffer/footer below?
6. Lower body box filled with dense, structured content (no empty bottom 30% within the box)?
7. If the slide carries data / comparison / process / structure — is it visualized as a chart or diagram (not narrated as prose)?
8. Every data point has a source line?
9. At most one Brand Glow element on the slide?
10. Hero Gradient (if used) only on permitted locations (cover hero card / section divider bg / 1 featured KPI), max 1 per slide and max 3 across the deck, never on chart bars or text?
11. 헤드라인 Bold 32pt / 부제 Regular 16pt #45515e / 본문 Regular 12–13pt — 위계가 서는가?
12. No emojis anywhere?
13. All chart/data text is live (not rasterized)?
