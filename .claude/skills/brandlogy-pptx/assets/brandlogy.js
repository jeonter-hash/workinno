/**
 * brandlogy.js — A4 가로 덱 헬퍼 (pptxgenjs 위에서 동작)
 *
 * 목적: 디자인 시스템의 "고정 좌표 5존 · 토큰 · 가드레일"을 코드로 고정해서
 *       장표마다 좌표를 다시 타이핑하다 생기는 표류를 없앤다.
 *
 * 사용법:
 *   const B = require('./brandlogy.js');
 *   const deck = B.createDeck({ logo: 'ksa_logo.png', title: '2026 사업계획' });
 *   const s = deck.slide({ chapter: '01 시장 진단', source: '출처: 통계청(2025)' });
 *   B.headline(s, '국내 수요는 3년째 두 자릿수로 커지고 있다');
 *   B.subtitle(s, '2023–2025 연평균 성장률 14.2%, 상위 3개 채널이 성장의 71%를 견인');
 *   B.kpiRow(s, [{ value: '14.2%', label: '연평균 성장률(CAGR)' }, ...], { y: B.BAND.A.kpi.y });
 *   await deck.save('out.pptx');
 *
 * 이후 반드시:
 *   python3 scripts/postprocess.py out.pptx     # Hero Gradient 센티넬 → 벡터 gradFill
 *   python3 scripts/check_layout.py out.pptx       # 존 고정·폰트·경계 자동 점검
 *   python3 /mnt/skills/public/pptx/scripts/office/validate.py out.pptx
 */

'use strict';

const PptxGenJS = require('pptxgenjs');

// ─────────────────────────────────────────────────────────── 단위
// A4 가로 = 9906000 × 6858000 EMU = 10.8333" × 7.5" (27.52 × 19.05cm)
const SLIDE_W = 10.8333;
const SLIDE_H = 7.5;
const px = (n) => n / 144;          // CSS px → inch (1560×1080 기준 = 144dpi)
const pxPt = (n) => n / 2;          // CSS px → pt
const inPt = (n) => n * 72;         // inch → pt

// ─────────────────────────────────────────────────────────── 색
const C = {
  brandBlue: '1456F0', blue500: '3B82F6', blue400: '60A5FA', blue200: 'BFDBFE',
  blue600: '2563EB', blue700: '1D4ED8', brandDeep: '17437D', sky: '3DAEFF',
  pink: 'EA5EC1',
  ink: '222222', inkDark: '18181B', surfaceDark: '181E25',
  sub: '45515E', muted: '8E8E93', helper: '5F5F5F',
  white: 'FFFFFF', surface: 'F0F0F0', divider: 'F2F3F5', border: 'E5E7EB',
  successBg: 'E8FFEA', successFg: '16A34A',
  // Hero Gradient 센티넬 — postprocess.py가 이 채움색을 찾아 gradFill로 교체한다
  GRADIENT: '0A0B0C',
};

// ─────────────────────────────────────────────────────────── 폰트 (맑은 고딕 전용)
// 맑은 고딕은 Semilight / Regular / Bold 세 단계뿐이다. 사양의 6단 웨이트를
// 아래처럼 접어서 쓰고, 500 vs 600처럼 접혀서 사라진 대비는 크기·색으로 만든다.
let WEIGHT_MODE = 'full';
const FACE = {
  full: { 300:['맑은 고딕 Semilight',false], 400:['맑은 고딕',false], 500:['맑은 고딕',false],
          600:['맑은 고딕',true], 700:['맑은 고딕',true], 800:['맑은 고딕',true] },
  // Semilight가 없는 PC(구형 Windows·macOS)용 — 300도 Regular로 접는다
  flat: { 300:['맑은 고딕',false], 400:['맑은 고딕',false], 500:['맑은 고딕',false],
          600:['맑은 고딕',true], 700:['맑은 고딕',true], 800:['맑은 고딕',true] },
};
const FONT = '맑은 고딕';
function setWeightMode(mode) {
  if (!FACE[mode]) throw new Error(`weight mode는 'full' 또는 'flat'`);
  WEIGHT_MODE = mode;
}
/** 웨이트 → { fontFace, bold } */
function w(weight) {
  const f = FACE[WEIGHT_MODE][weight];
  if (!f) throw new Error(`지원하지 않는 웨이트: ${weight} (300/400/500/600/700/800만)`);
  return { fontFace: f[0], bold: f[1] };
}

// ─────────────────────────────────────────────────────────── 존 (전 장표 고정)
const Z = {
  header:   { x: 0.5, y: 0.40, w: 9.8333, h: 0.30 },
  chapter:  { x: 0.5, y: 0.40, w: 6.0,    h: 0.30 },
  logo:     { x: 9.1133, y: 0.44, w: 1.22, h: 0.24 },  // 폭은 로고 원본 비율로 자동 보정된다
  headline: { x: 0.5, y: 1.00, w: 9.8333, h: 0.75 },
  subtitle: { x: 0.5, y: 1.63, w: 9.8333, h: 0.40 },
  body:     { x: 0.5, y: 2.39, w: 9.8333, h: 4.46 },
  footer:   { x: 0.5, y: 7.05, w: 9.8333, h: 0.25 },
};
const LOGO_H = 0.24, LOGO_MAX_W = 1.9;
const BODY_TOP = 2.39, BODY_BOTTOM = 6.85, EPS = 0.004;

// 12열 그리드
const GRID = { cols: 12, gutter: 0.2, col: (9.8333 - 11 * 0.2) / 12 };
const colX = (i) => +(Z.body.x + i * (GRID.col + GRID.gutter)).toFixed(4);
const colW = (n) => +(n * GRID.col + (n - 1) * GRID.gutter).toFixed(4);
/** 균등 n분할 → [{x,w}, ...] (n은 12의 약수: 2,3,4,6) */
function split(n) {
  if (12 % n !== 0) throw new Error('split은 2, 3, 4, 6만 (12열 균등 분할)');
  const span = 12 / n;
  return Array.from({ length: n }, (_, i) => ({ x: colX(i * span), w: colW(span) }));
}

// 패턴별 밴드 (합계 = 4.46")
const BAND = {
  A: { kpi:   { y: 2.39, h: 1.60 }, detail: { y: 4.23, h: 2.62 } },
  B: { cols:  { y: 2.39, h: 3.66 }, callout: { y: 6.25, h: 0.60 }, colsFull: { y: 2.39, h: 4.46 } },
  C: { figure:{ y: 2.39, h: 3.56 }, caption: { y: 6.15, h: 0.70 } },
  D: { stages:{ y: 2.69, h: 1.90 }, outcome: { y: 4.79, h: 2.06 } },
  E: { quote: { y: 2.39, h: 4.46 }, cards: [{ y: 2.39, h: 1.36 }, { y: 3.94, h: 1.36 }, { y: 5.49, h: 1.36 }] },
  F: { kpi:   { y: 2.39, h: 1.30 }, figure: { y: 3.89, h: 1.86 }, evidence: { y: 5.95, h: 0.90 } },
};

// 반경 (CSS px → inch)
const R = { tag: px(4), button: px(8), card: px(13), large: px(16), xl: px(20), hero: px(24), badge: px(32) };

// 그림자 — 매번 새 객체를 반환한다(pptxgenjs가 옵션 객체를 제자리에서 변형하므로 재사용 금지)
const SHADOW = {
  standard:   () => ({ type: 'outer', color: '000000', opacity: 0.08, blur: 3,    offset: 2,   angle: 90 }),
  softGlow:   () => ({ type: 'outer', color: '000000', opacity: 0.08, blur: 11.3, offset: 0,   angle: 90 }),
  brandGlow:  () => ({ type: 'outer', color: '2C1E74', opacity: 0.16, blur: 7.5,  offset: 0,   angle: 90 }),
  brandGlowOffset: () => ({ type: 'outer', color: '2C1E74', opacity: 0.11, blur: 8.8, offset: 3.4, angle: 17 }),
  elevated:   () => ({ type: 'outer', color: '242424', opacity: 0.08, blur: 8,    offset: 6,   angle: 90 }),
};

// ─────────────────────────────────────────────────────────── 로고
/** PNG·JPEG 헤더에서 픽셀 크기를 읽는다(의존성 없음). 실패하면 null */
function imageSize(file) {
  try {
    const b = require('fs').readFileSync(file);
    if (b.slice(1, 4).toString() === 'PNG') {
      return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
    }
    if (b[0] === 0xff && b[1] === 0xd8) {            // JPEG: SOF 마커에서 크기를 읽는다
      let i = 2;
      while (i < b.length - 9) {
        if (b[i] !== 0xff) { i += 1; continue; }
        const m = b[i + 1];
        if (m >= 0xc0 && m <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(m)) {
          return { w: b.readUInt16BE(i + 7), h: b.readUInt16BE(i + 5) };
        }
        if (m === 0xd8 || m === 0xd9 || (m >= 0xd0 && m <= 0xd7)) { i += 2; continue; }
        i += 2 + b.readUInt16BE(i + 2);
      }
    }
    return null;
  } catch (e) { return null; }
}
const pngSize = imageSize;   // 옛 이름 유지

/**
 * 로고 배치 박스. 높이 0.24"에 맞추고 폭은 **원본 비율로 자동 계산**한다
 * (임의 폭을 강제하면 로고가 늘어난다 — 로고 무결성 규칙 위반).
 * 아주 가로로 긴 로고는 폭 상한(1.9")에 맞춰 높이를 줄인다. 오른쪽 끝은 항상 0.5" 여백.
 */
function logoBox(file) {
  const sz = file ? imageSize(file) : null;
  let h = LOGO_H;
  let wd = sz ? +(h * (sz.w / sz.h)).toFixed(4) : Z.logo.w;
  if (wd > LOGO_MAX_W) { h = +(LOGO_MAX_W * (sz.h / sz.w)).toFixed(4); wd = LOGO_MAX_W; }
  return { x: +(SLIDE_W - 0.5 - wd).toFixed(4), y: +(Z.logo.y + (LOGO_H - h) / 2).toFixed(4), w: wd, h };
}

/**
 * 동봉된 한국표준협회 로고를 찾는다. 헬퍼를 build 폴더로 복사해 쓰는 워크플로에서도
 * 로고가 빠지지 않도록 여러 위치를 순서대로 뒤진다. 못 찾으면 null(로고 없이 진행).
 * 환경변수 KSA_LOGO로 직접 지정할 수도 있다.
 */
function defaultLogo() {
  const path = require('path'), fs = require('fs'), os = require('os');
  const names = ['ksa_logo.png', 'ksa_logo.jpg'];
  const dirs = [__dirname, path.join(__dirname, 'assets'), process.cwd()];
  // cwd에서 위로 올라가며 .claude/skills/brandlogy-pptx/assets 를 찾는다
  let cur = process.cwd();
  for (let i = 0; i < 6; i++) {
    dirs.push(path.join(cur, '.claude', 'skills', 'brandlogy-pptx', 'assets'));
    const up = path.dirname(cur);
    if (up === cur) break;
    cur = up;
  }
  dirs.push(path.join(os.homedir(), '.claude', 'skills', 'brandlogy-pptx', 'assets'));

  if (process.env.KSA_LOGO && fs.existsSync(process.env.KSA_LOGO)) return process.env.KSA_LOGO;
  for (const d of dirs) {
    for (const n of names) {
      const p = path.join(d, n);
      try { if (fs.existsSync(p)) return p; } catch (e) { /* 접근 불가 경로는 건너뛴다 */ }
    }
  }
  return null;
}

// ─────────────────────────────────────────────────────────── 가드레일
function assertBody(y, h, what = '본문 요소') {
  if (y < BODY_TOP - EPS) throw new Error(`${what}가 본문 상단(2.39")을 침범: y=${y}`);
  if (y + h > BODY_BOTTOM + EPS) throw new Error(`${what}가 본문 하단(6.85")을 침범: y+h=${(y + h).toFixed(3)} — 내용을 줄이거나 장표를 분할할 것`);
  return true;
}
function noEmoji(t) {
  const s = Array.isArray(t) ? t.map((x) => (x && x.text) || x).join(' ') : String(t == null ? '' : t);
  if (/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE0F}]/u.test(s)) throw new Error(`이모지 금지: ${s.slice(0, 40)}`);
  return t;
}
function useGradient(slide) {
  const d = slide._deck;
  if (slide._grad >= 1) throw new Error('Hero Gradient는 장표당 1개까지');
  if (d._grad >= 3) throw new Error('Hero Gradient는 덱 전체 3개까지');
  slide._grad += 1; d._grad += 1;
  return C.GRADIENT;
}
function useBrandGlow(slide) {
  if (slide._glow >= 1) throw new Error('Brand Glow는 장표당 1개까지');
  slide._glow += 1;
  return SHADOW.brandGlow();
}

// ─────────────────────────────────────────────────────────── 덱
function createDeck(opts = {}) {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'A4_LANDSCAPE', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'A4_LANDSCAPE';
  if (opts.title) pres.title = opts.title;
  if (opts.author) pres.author = opts.author;
  if (opts.subject) pres.subject = opts.subject;
  if (opts.company) pres.company = opts.company;

  const deck = {
    pres, _grad: 0, _page: 0,
    logo: opts.logo === undefined ? defaultLogo() : opts.logo,  // 동봉된 KSA 로고를 기본값으로 (원본 그대로)
    logoWhite: opts.logoWhite || null, // 어두운 배경용 흰 변형(균일 반전본)

    /** 표준 5존 프레임이 적용된 본문 장표 */
    slide(o = {}) {
      const s = pres.addSlide();
      s.background = { color: C.white };
      s._deck = deck; s._grad = 0; s._glow = 0;
      deck._page += 1;
      s._page = o.page == null ? deck._page : o.page;
      frame(s, o);
      return s;
    },

    /** 프레임 없는 특수 장표(표지·섹션 디바이더·클로징) */
    bareSlide(o = {}) {
      const s = pres.addSlide();
      s.background = { color: o.bg || C.white };
      s._deck = deck; s._grad = 0; s._glow = 0;
      deck._page += 1;
      s._page = o.page == null ? deck._page : o.page;
      return s;
    },

    async save(path) {
      await pres.writeFile({ fileName: path });
      if (deck._grad > 0) {
        console.log(`[deck] Hero Gradient ${deck._grad}개 — 반드시 실행: python3 scripts/postprocess.py ${path}`);
      }
      if (!deck.logo) {
        console.log('[deck] 경고: 로고 파일이 지정되지 않아 장표에 로고가 없다. 한국표준협회(KSA) 누끼 PNG를 지정할 것.');
      }
      return path;
    },
  };
  return deck;
}

// ─────────────────────────────────────────────────────────── 존 요소
/** 헤더(챕터명 + 로고) · 푸터(페이지 + 출처) */
function frame(slide, o = {}) {
  const deck = slide._deck;
  if (o.chapter) {
    slide.addText(noEmoji(o.chapter), {
      ...Z.chapter, ...w(600), fontSize: 12, color: C.muted,
      valign: 'middle', align: 'left', margin: 0, lineSpacingMultiple: 1.3,
    });
  }
  if (deck.logo) {
    // 원본 그대로 — 배경/테두리/그림자/보정 금지. 폭은 원본 비율로 자동 계산한다
    slide.addImage({ path: deck.logo, ...logoBox(deck.logo), altText: '한국표준협회 로고' });
  }
  slide.addText(String(o.page == null ? slide._page : o.page), {
    ...Z.footer, w: 4.0, ...w(500), fontSize: 10, color: C.muted,
    valign: 'middle', align: 'left', margin: 0,
  });
  if (o.source) {
    slide.addText(noEmoji(o.source), {
      ...Z.footer, x: Z.footer.x + 4.333, w: 8.0, ...w(400), fontSize: 9.5, color: C.muted,
      valign: 'middle', align: 'right', margin: 0,
    });
  }
  return slide;
}

function headline(slide, text, o = {}) {
  const size = o.fontSize || 32;   // A4 가로에서는 32pt 한 줄(한글 22자 내외)이 기본
  slide.addText(noEmoji(text), {
    ...Z.headline, ...w(700), fontSize: size, color: o.color || C.ink,
    valign: 'top', align: 'left', margin: 0,
    lineSpacingMultiple: 1.2, charSpacing: o.charSpacing == null ? -0.65 : o.charSpacing, // ≈ -0.02em @32pt
  });
  return slide;
}

function subtitle(slide, text, o = {}) {
  slide.addText(noEmoji(text), {
    ...Z.subtitle, ...w(500), fontSize: 16, color: o.color || C.sub,
    valign: 'top', align: 'left', margin: 0, lineSpacingMultiple: 1.45,
  });
  return slide;
}

// ─────────────────────────────────────────────────────────── 본문 컴포넌트
/** 카드 껍데기. kind: 'standard' | 'data' | 'featured' | 'gradient' | 'plain' */
function card(slide, o) {
  const { x, y, w: cw, h } = o;
  assertBody(y, h, o.what || '카드');
  const kind = o.kind || 'standard';
  const shape = { x, y, w: cw, h, rectRadius: o.radius || (kind === 'gradient' || kind === 'featured' ? R.hero : R.card) };
  if (kind === 'gradient') {
    shape.fill = { color: useGradient(slide) };
    shape.shadow = useBrandGlow(slide);
  } else if (kind === 'featured') {
    shape.fill = { color: o.fill || C.white };
    shape.shadow = useBrandGlow(slide);
  } else if (kind === 'data') {
    shape.fill = { color: o.fill || C.white };
    shape.line = { color: C.divider, width: 1 };
  } else if (kind === 'plain') {
    shape.fill = { color: o.fill || C.white };
  } else {
    shape.fill = { color: o.fill || C.white };
    shape.shadow = SHADOW.standard();
  }
  if (o.line) shape.line = o.line;
  slide.addShape('roundRect', shape);
  return { x, y, w: cw, h };
}

/** KPI 타일 1장 */
function kpiCard(slide, o) {
  const pad = o.pad == null ? px(20) : o.pad;
  const gradient = !!o.gradient;
  card(slide, { ...o, kind: gradient ? 'gradient' : (o.featured ? 'featured' : 'standard'), what: 'KPI 카드' });
  const numColor = gradient ? C.white : (o.color || C.brandBlue);
  const labColor = gradient ? 'FFFFFF' : C.sub;
  const numSize = o.valueSize || 36;   // A4 가로 폭(9.83") 기준
  slide.addText(noEmoji(o.value), {
    x: o.x + pad, y: o.y + pad, w: o.w - 2 * pad, h: o.h - 2 * pad - 0.26,
    ...w(700), fontSize: numSize, color: numColor,
    valign: 'bottom', align: 'left', margin: 0, lineSpacingMultiple: 1.1,
  });
  slide.addText(noEmoji(o.label), {
    x: o.x + pad, y: o.y + o.h - pad - 0.24, w: o.w - 2 * pad, h: 0.24,
    ...w(500), fontSize: 11.5, color: labColor, transparency: gradient ? 15 : 0,
    valign: 'middle', align: 'left', margin: 0, lineSpacingMultiple: 1.3,
  });
  return o;
}

/** KPI 3–4장 한 줄 */
function kpiRow(slide, items, o = {}) {
  if (items.length < 2 || items.length > 4) throw new Error('KPI 행은 2–4장');
  const band = o.y == null ? BAND.A.kpi : { y: o.y, h: o.h == null ? BAND.A.kpi.h : o.h };
  const cells = split(items.length === 3 ? 3 : (items.length === 2 ? 2 : 4));
  return items.map((it, i) => kpiCard(slide, { ...cells[i], y: band.y, h: band.h, ...it }));
}

/** 차트 컨테이너(제목 + 차트 영역 + 출처). 반환값의 area에 addChart를 그린다 */
function dataCard(slide, o) {
  const pad = o.pad == null ? px(20) : o.pad;
  card(slide, { ...o, kind: o.kind || 'data', what: '데이터 카드' });
  let top = o.y + pad, bottom = o.y + o.h - pad;
  if (o.title) {
    slide.addText(noEmoji(o.title), {
      x: o.x + pad, y: top, w: o.w - 2 * pad, h: 0.26,
      ...w(600), fontSize: 14, color: C.ink, valign: 'middle', align: 'left', margin: 0,
    });
    top += 0.32;
  }
  if (o.source) {
    slide.addText(noEmoji(o.source), {
      x: o.x + pad, y: bottom - 0.20, w: o.w - 2 * pad, h: 0.20,
      ...w(400), fontSize: 9, color: C.muted, valign: 'middle', align: 'left', margin: 0,
    });
    bottom -= 0.26;
  }
  return { area: { x: o.x + pad * 0.5, y: top, w: o.w - pad, h: +(bottom - top).toFixed(4) } };
}

/** 차트 기본 옵션 (§8) — 넘겨받은 값으로 덮어쓴다 */
function chartOpts(over = {}) {
  const base = {
    chartColors: [C.brandBlue, C.blue400, C.blue200, C.brandDeep],
    showLegend: false,
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 10, catAxisLabelColor: C.sub,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 10, valAxisLabelColor: C.sub,
    dataLabelFontFace: FONT, dataLabelFontSize: 11, dataLabelFontBold: true, dataLabelColor: C.ink,
    showValue: true, dataLabelPosition: 'outEnd',
    valGridLine: { color: C.border, size: 1 },
    catGridLine: { style: 'none' },
    valAxisLineShow: false, catAxisLineShow: false,
    chartArea: { fill: { color: C.white } },
  };
  const o = { ...base, ...over };
  if (o.barGrouping === 'stacked' || o.barGrouping === 'percentStacked') {
    if (!['ctr', 'inEnd', 'inBase'].includes(o.dataLabelPosition)) o.dataLabelPosition = 'ctr'; // outEnd는 파일을 깨뜨림
  }
  return o;
}

/** 본문 소제목(H2/H3) */
function h2(slide, text, o) {
  assertBody(o.y, o.h == null ? 0.3 : o.h, '본문 중제목');
  slide.addText(noEmoji(text), {
    x: o.x, y: o.y, w: o.w, h: o.h == null ? 0.3 : o.h,
    ...w(600), fontSize: o.fontSize || 18, color: o.color || C.ink,
    valign: 'top', align: 'left', margin: 0, lineSpacingMultiple: 1.4,
  });
}

/** 불릿 본문 */
function bullets(slide, items, o) {
  assertBody(o.y, o.h, '불릿');
  const runs = items.map((t, i) => ({
    text: noEmoji(typeof t === 'string' ? t : t.text),
    options: { bullet: true, breakLine: i < items.length - 1, ...(typeof t === 'object' && t.bold ? w(700) : {}) },
  }));
  slide.addText(runs, {
    x: o.x, y: o.y, w: o.w, h: o.h, ...w(400), fontSize: o.fontSize || 13, color: C.ink,
    valign: 'top', align: 'left', margin: 0, lineSpacingMultiple: 1.5, paraSpaceAfter: pxPt(10),
  });
}

/** "So What" 콜아웃 — 본문 하단 밴드 */
function soWhat(slide, text, o = {}) {
  const band = { x: Z.body.x, w: Z.body.w, ...(BAND.B.callout), ...o };
  card(slide, { ...band, kind: 'plain', fill: C.divider, radius: R.card, what: 'So What 콜아웃' });
  slide.addText(noEmoji(text), {
    x: band.x + px(16), y: band.y, w: band.w - 2 * px(16), h: band.h,
    ...w(600), fontSize: 14, color: C.ink, valign: 'middle', align: 'left', margin: 0,
  });
  return band;
}

/** 태그/배지 필 */
function pill(slide, text, o) {
  const h = o.h == null ? 0.26 : o.h;
  slide.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h, rectRadius: h / 2,
    fill: { color: o.fill || C.surface }, line: o.line || { color: o.fill || C.surface, width: 0 },
  });
  slide.addText(noEmoji(text), {
    x: o.x, y: o.y, w: o.w, h, ...w(600), fontSize: o.fontSize || 10.5,
    color: o.color || C.inkDark, valign: 'middle', align: 'center', margin: 0,
  });
}

/** 캡션 / 출처 한 줄 */
function caption(slide, text, o) {
  slide.addText(noEmoji(text), {
    x: o.x, y: o.y, w: o.w, h: o.h == null ? 0.2 : o.h,
    ...w(400), fontSize: o.fontSize || 9, color: C.muted,
    valign: 'middle', align: o.align || 'left', margin: 0,
  });
}

// ─────────────────────────────────────────────────────────── 특수 장표
/** 표지: 5존 프레임 + 본문 자리에 Hero Gradient 카드 하나 */
function cover(deck, o) {
  const s = deck.slide({ chapter: o.chapter || '', page: o.page, source: o.source });
  headline(s, o.title, { fontSize: o.titleSize || 32 });
  if (o.subtitle) subtitle(s, o.subtitle);
  const box = { x: Z.body.x, y: Z.body.y, w: Z.body.w, h: Z.body.h, ...(o.box || {}) };
  card(s, { ...box, kind: 'gradient', radius: R.hero, what: '표지 히어로 카드' });
  const pad = px(48);
  slide_text(s, o.kpi, { x: box.x + pad, y: box.y + box.h / 2 - 0.85, w: box.w - 2 * pad, h: 0.95 },
    { ...w(700), fontSize: 44, color: C.white, lineSpacingMultiple: 1.1 });
  slide_text(s, o.kpiLabel, { x: box.x + pad, y: box.y + box.h / 2 + 0.12, w: box.w - 2 * pad, h: 0.3 },
    { ...w(500), fontSize: 12, color: C.white, transparency: 15 });
  return s;
}

/** 섹션 디바이더: 프레임을 의도적으로 깨는 어두운/그라디언트 전면 장표 */
function divider(deck, o) {
  const useGrad = !!o.gradient;
  const s = deck.bareSlide({ bg: useGrad ? C.white : (o.bg || C.surfaceDark), page: o.page });
  if (useGrad) {
    s.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: useGradient(s) }, line: { width: 0 } });
  }
  if (o.number) {
    s.addText(noEmoji(o.number), { ...Z.chapter, ...w(600), fontSize: 14, color: C.white, transparency: 40,
      valign: 'middle', align: 'left', margin: 0 });
  }
  if (deck.logoWhite) s.addImage({ path: deck.logoWhite, ...logoBox(deck.logoWhite), altText: '한국표준협회 로고' });
  s.addText(noEmoji(o.title), { x: 0.5, y: 3.0, w: 9.8333, h: 1.1, ...w(700), fontSize: 48, color: C.white,
    valign: 'bottom', align: 'left', margin: 0, lineSpacingMultiple: 1.15, charSpacing: -1.2 });
  if (o.lead) {
    s.addText(noEmoji(o.lead), { x: 0.5, y: 4.2, w: 8.5, h: 0.5, ...w(500), fontSize: 20, color: C.white,
      transparency: 30, valign: 'top', align: 'left', margin: 0, lineSpacingMultiple: 1.45 });
  }
  s.addText(String(o.page == null ? s._page : o.page), { ...Z.footer, w: 4.0, ...w(500), fontSize: 10,
    color: C.white, transparency: 40, valign: 'middle', align: 'left', margin: 0 });
  return s;
}

function slide_text(slide, text, box, style) {
  if (text == null) return;
  slide.addText(noEmoji(text), { ...box, valign: 'middle', align: 'left', margin: 0, ...style });
}

module.exports = {
  PptxGenJS, SLIDE_W, SLIDE_H, px, pxPt, inPt,
  C, Z, GRID, BAND, R, SHADOW, colX, colW, split,
  w, FONT, setWeightMode, assertBody, noEmoji, imageSize, pngSize, logoBox, defaultLogo, LOGO_H,
  createDeck, frame, headline, subtitle,
  card, kpiCard, kpiRow, dataCard, chartOpts, h2, bullets, soWhat, pill, caption,
  cover, divider,
  BODY_TOP, BODY_BOTTOM,
};
