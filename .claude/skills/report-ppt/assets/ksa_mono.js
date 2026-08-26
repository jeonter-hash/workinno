/**
 * ksa_mono.js — KSA 모노톤 A4 가로 공용 컴포넌트
 * 두 스킬(발표용 present / 보고서용 report)이 같은 프레임·도식을 공유하고,
 * 밀도(글자 크기·여백·행 높이)만 모드로 갈아끼운다.
 */
'use strict';
const P = require('pptxgenjs');
const fs = require('fs'), path = require('path');

/* ── 무채색 팔레트 (색상 없음) ───────────────────────────── */
const K = { ink:'111111', g1:'3A3A3A', g2:'6B6B6B', g3:'9E9E9E', g4:'C7C7C7', g5:'E4E4E4', g6:'F4F4F4', g7:'FAFAFA', w:'FFFFFF' };

/**
 * 색 체계. 기본은 무채색(mono)이고, 'accent'는 KSA 로고에서 뽑은 두 색만 더한다.
 *   dark  구조색 — 표 머리글·체브론 진행 단계·트리 루트·조직 최상위
 *   acc   강조색 — 장표당 1~2곳만 (우선 항목·마일스톤·핵심 KPI)
 * 회색 스케일과 본문 글자색은 두 모드가 같다. 색이 의미를 갖는 자리에만 쓴다.
 */
const PALETTE = {
  mono:   { dark:K.ink,   mid:K.g2,     acc:K.ink,   accSoft:K.g6,   onDark:K.w },
  accent: { dark:'1F3864', mid:'4A6491', acc:'C4303C', accSoft:'F2E4E6', onDark:K.w },
};
const FONT = '맑은 고딕';
const TEXT_MIN = K.g2;          // 흰 배경 위 글자의 최소 명도 — 이보다 흐리면 쓰지 않는다

/* ── 판형·존 (A4 가로) ──────────────────────────────────── */
const W = 10.8333, H = 7.5, M = 0.5, CW = W - 2*M;
const BB = 6.85;              // 본문 하한 (벽) — 모드 공통
const TITLE_Y = 0.94;         // 타이틀 상자 상단 — 모드 공통
const TITLE_MAX = 30;         // 타이틀 글자 수 상한 (1줄)

/* 헤드 존의 좌표는 **렌더 실측**으로 잡았다.
 * 맑은 고딕은 상자 위에서 글자까지 여백이 생기고 줄 간격도 계산값보다 넓다.
 * 그래서 pt×행간÷72로 계산하면 어긋난다. 실측값은 다음과 같다(105dpi 렌더 측정).
 *
 *   타이틀 20pt : 상자 위에서 0.136" 아래부터 글자, 0.383"에서 끝
 *   타이틀 24pt : 0.155" / 0.459"
 *   메시지 12pt : 0.086"부터, 줄 간격 0.266", 2줄 글자 높이 0.400"
 *   메시지 13pt : 0.102"부터, 줄 간격 0.288", 2줄 글자 높이 0.438"
 *
 * 이 값으로 **타이틀↔메시지 간격 = 메시지↔본문 간격 = 0.14"**가 되게 역산했다.
 * 좌표를 손대려면 scripts/measure_head.py 로 다시 재고 나서 고칠 것.
 */
const Z = {
  chapter:{ x:M, y:0.40, w:6.2, h:0.30 },
  rule:   { y:0.82 },
  foot:   { x:M, y:7.05, w:CW, h:0.25 },
};
/** 모드별 헤드·본문 존 */
const zones = T => ({
  title: { x:M, y:TITLE_Y, w:CW, h:T.titleH },
  msg:   { x:M, y:T.msgY,  w:CW, h:T.msgH  },   // 2줄분 고정 — 1줄이어도 본문은 안 올라온다
  body:  { x:M, y:T.bt,    w:CW, h:BB-T.bt },
});
const col = (CW-11*0.2)/12;
const cx = i => +(M + i*(col+0.2)).toFixed(4);
const cw = n => +(n*col + (n-1)*0.2).toFixed(4);
const split = n => { if (12%n) throw new Error('split은 2·3·4·6'); const sp=12/n;
  return Array.from({length:n},(_,i)=>({ x:cx(i*sp), w:cw(sp) })); };

/* ── 밀도 프로파일 ──────────────────────────────────────── */
const MODE = {
  present: { name:'발표용', chapter:14, h1:24, sub:13, h2:14, body:11.5, small:10, cap:9.5,
             kpi:34, kpiLabel:10.5, tableHead:10.5, tableBody:10.5, rowH:0.44, pad:0.17, lh:1.35, gap:0.20,
             titleH:0.46, msgY:1.43, msgH:0.62, bt:2.11, msgMax:120 },
  report:  { name:'보고서용', chapter:12, h1:20, sub:11.5, h2:12, body:9.5, small:9, cap:9,
             kpi:24, kpiLabel:9, tableHead:9.5, tableBody:9.5, rowH:0.335, pad:0.12, lh:1.28, gap:0.14,
             titleH:0.40, msgY:1.37, msgH:0.58, bt:2.00, msgMax:130 },
  // dense — 20행 표처럼 내용이 많을 때만. 한 장 수용량이 약 18% 늘어난다(11행 → 13행).
  // 8.5pt가 하한이며 그 아래는 인쇄·투사 어디서도 안전하지 않다.
  // dense는 보고서와 같은 헤드 상자를 쓴다 — 글자만 작아지므로 간격이 조금 더 벌어진다.
  report_dense: { name:'보고서용(고밀도)', chapter:11.5, h1:19, sub:11, h2:11.5, body:9, small:8.5, cap:8.5,
             kpi:22, kpiLabel:8.5, tableHead:9, tableBody:9, rowH:0.295, pad:0.10, lh:1.25, gap:0.12,
             titleH:0.40, msgY:1.37, msgH:0.58, bt:2.00, msgMax:135 },
};

/* ── 글자 폭 추정 (밑줄을 글자 폭에 맞추는 데 사용) ───────── */
function textW(str, pt){
  let em=0;
  for (const ch of String(str)){
    const c=ch.codePointAt(0);
    if (ch===' ') em+=0.30;
    else if (c>0x1100) em+=1.0;
    else if (ch>='A'&&ch<='Z') em+=0.62;
    else if (ch>='0'&&ch<='9') em+=0.55;
    else em+=0.50;
  }
  return +(em*pt/72).toFixed(4);
}
/** 상자 폭 기준 줄 수 추정 — 도형 밖으로 글자가 넘치는지 미리 본다 */
const lines = (str, pt, boxW) => Math.max(1, Math.ceil(textW(str,pt)/(boxW-0.06)));
/** 글자가 들어갈 최소 높이(인치) */
const needH = (str, pt, boxW, lh=1.3) => +(lines(str,pt,boxW)*pt*lh/72 + 0.06).toFixed(3);

function createDeck(o={}){
  // o.docTitle — 본문 헤더 우측에 반복 표기할 문서 제목 (없으면 o.title)
  // o.dense    — 보고서용에서만. 표가 20행에 이르는 등 내용이 많을 때 한 단계 낮춘 밀도를 쓴다.
  const key = (o.mode||'present') === 'report' && o.dense ? 'report_dense' : (o.mode||'present');
  const T = MODE[key];
  if (!T) throw new Error("mode는 'present' 또는 'report'");
  const C = PALETTE[o.palette || 'mono'];
  if (!C) throw new Error("palette는 'mono' 또는 'accent'");
  const pres = new P();
  pres.defineLayout({ name:'A4L', width:W, height:H }); pres.layout='A4L';
  // 본문 마스터 — 헤더 헤어라인 + PowerPoint 자동 슬라이드 번호(하단 가운데)
  const docTitle = o.docTitle || o.title || '문서 제목';
  pres.defineSlideMaster({
    title:'KSA_BODY', background:{ color:K.w },
    objects:[
      { line:{ x:M, y:Z.rule.y, w:CW, h:0, line:{ color:K.g4, width:0.75 } } },
      // 우측 상단 문서 제목 — 마스터에 있으므로 [보기 → 슬라이드 마스터]에서 한 번 고치면 전 장표에 반영된다
      { text:{ text:docTitle, options:{ x:W-M-5.0, y:Z.chapter.y, w:5.0, h:Z.chapter.h,
        fontFace:FONT, fontSize:T.chapter, color:K.g3, align:'right', valign:'middle', margin:0 } } },
    ],
    slideNumber:{ x:0, y:7.05, w:W, h:0.25, align:'center',
                  fontFace:FONT, fontSize:T.cap, color:K.g2 },
  });
  // 표지·간지 마스터 — 번호 없음
  pres.defineSlideMaster({ title:'KSA_TITLE', background:{ color:K.w } });
  if (o.title) pres.title=o.title;
  if (o.author) pres.author=o.author;
  const deck = { pres, T, C, page:0, bt:T.bt, BB, docTitle: o.docTitle || o.title || '',
    slide(opt={}){
      const s = pres.addSlide({ masterName:'KSA_BODY' }); s._T=T; s._C=C; s._deck=deck;
      deck.page += 1; s._page = opt.page==null? deck.page : opt.page;
      frame(s, opt); return s;
    },
    bare(opt={}){ const s=pres.addSlide({ masterName:'KSA_TITLE' });
      if (opt.bg) s.background={color:opt.bg}; s._T=T; s._C=C; s._deck=deck;
      deck.page+=1; s._page=opt.page==null?deck.page:opt.page; return s; },
    save(f){ return pres.writeFile({ fileName:f }).then(()=>f); },
  };
  return deck;
}

/* ── 5존 프레임 ─────────────────────────────────────────── */
function frame(s, o={}){
  const T=s._T, d=s._deck;
  if (o.chapter) s.addText(o.chapter, { ...Z.chapter, fontFace:FONT, fontSize:T.chapter, bold:true,
    color:K.g2, valign:'middle', margin:0, charSpacing:0.2 });
  // 우측 상단(문서 제목)은 슬라이드 마스터가 그린다 — 여기서는 비워 둔다.
  // 헤더 헤어라인과 페이지 번호는 마스터가 그린다. 출처는 본문 안 caption/source로 붙인다.
  return s;
}
/**
 * 헤드 — 타이틀과 메시지로 나눈다.
 *   head(s, { title:'업무별 시간 구조 및 절감 여력',
 *             message:'상위 4개 업무가 …를 차지하는 반면 잔여 8개 업무는 … 제한적임' })
 *
 * 타이틀은 주제를 가리키는 짧은 명사구(30자·1줄). 챕터에 이미 있는 대분류는 넣지 않는다.
 * 메시지는 판단을 담은 한 문장이며 명사형 개조식(~함/~임/~해야 함/…)으로 끝낸다.
 * 자세한 문법과 예문은 references/headline.md 를 볼 것.
 *
* 문자열 하나만 넘기면 타이틀만 그린다(목차·간지용).
 * 줄 수는 글자 폭을 실제로 계산해 판정하므로 숫자·영문이 많으면 더 들어간다.
 */
function head(s, o, opt={}){
  const T = s._T, Zn = zones(T);
  const title = typeof o === 'string' ? o : o.title;
  const message = typeof o === 'string' ? null : o.message;
  if (!title) throw new Error('head: 타이틀이 없다');

  if (title.length > TITLE_MAX)
    throw new Error(`헤드 타이틀이 ${title.length}자 — 상한 ${TITLE_MAX}자. `
      + `${title.length - TITLE_MAX}자를 줄일 것. 대분류는 챕터에 있으므로 넣지 않는다: "${title}"`);
  if (lines(title, opt.size||T.h1, Zn.title.w) > 1)
    throw new Error(`헤드 타이틀이 한 줄을 넘는다 — 더 짧게 쓸 것: "${title}"`);

  s.addText(title, { ...Zn.title, fontFace:FONT, fontSize:opt.size||T.h1, bold:true,
    color:K.ink, valign:'top', margin:0, charSpacing:-0.4, lineSpacingMultiple:1.2 });
  if (!message) return s;

  if (message.length > T.msgMax)
    throw new Error(`헤드 메시지가 ${message.length}자 — ${T.name} 상한 ${T.msgMax}자(2줄). `
      + `${message.length - T.msgMax}자를 줄일 것: "${message.slice(0,40)}…"`);
  const n = lines(message, T.sub, Zn.msg.w);
  if (n > 2)
    throw new Error(`헤드 메시지가 ${n}줄 — 2줄까지만 쓴다. 문장을 나누지 말고 줄일 것: "${message.slice(0,40)}…"`);

  s.addText(message, { ...Zn.msg, fontFace:FONT, fontSize:T.sub, color:K.g1,
    valign:'top', margin:0, lineSpacingMultiple:1.30 });
  return s;
}

/* ── 기본 조각 ──────────────────────────────────────────── */
/** 본문 벽 검사 — 본문 상단은 모드마다 다르므로 슬라이드에서 읽는다 */
function guard(s,y,h,what='요소'){
  const bt = s && s._T ? s._T.bt : 2.00;
  if (y < bt-0.005) throw new Error(`${what}가 본문 상단(${bt}")을 침범: y=${y}`);
  if (y+h > BB+0.005) throw new Error(`${what}가 본문 하단(${BB}")을 침범: ${(y+h).toFixed(3)}`);
}
function box(s,o){
  guard(s,o.y,o.h,o.what||'박스');
  const noLine = o.line==='none' || o.lw===0;
  s.addShape(o.round?'roundRect':'rect', { x:o.x, y:o.y, w:o.w, h:o.h, ...(o.round?{rectRadius:o.round}:{}),
    fill:{ color:o.fill||K.w },
    line: noLine ? { type:'none' } : { color:o.line||K.g4, width:o.lw==null?0.75:o.lw } });
  return o;
}
function txt(s,t,o){
  s.addText(t, { x:o.x, y:o.y, w:o.w, h:o.h, fontFace:FONT, fontSize:o.sz||s._T.body,
    bold:!!o.b, color:o.c||K.ink, align:o.align||'left', valign:o.valign||'top', margin:0,
    lineSpacingMultiple:o.lh||s._T.lh, ...(o.wrap===false?{wrap:false}:{}), ...(o.space?{paraSpaceAfter:o.space}:{}) });
}
/** 밑줄 — 반드시 위 글자의 실제 폭에 맞춘다 */
const underline = (s,t,pt,o) => s.addShape('line', { x:o.x, y:o.y, w:textW(t,pt), h:0,
  line:{ color:o.color||K.ink, width:o.width||2 } });
/** 구분선 — 위/아래 글자와 최소 0.09" 띄운다 */
const hr = (s,o) => s.addShape('line', { x:o.x, y:o.y, w:o.w, h:0, line:{ color:o.color||K.g4, width:o.width||0.75 } });

/* ── 컴포넌트 ───────────────────────────────────────────── */
function sectionTitle(s, t, o){
  const T=s._T;
  txt(s,t,{ ...o, sz:o.sz||T.h2, b:true, c:K.ink, h:o.h||0.28 });
  hr(s,{ x:o.x, y:o.y+(o.h||0.28)+0.04, w:o.w, color:K.ink, width:1.25 });
}
function kpiRow(s, items, o){
  const T=s._T, cells=split(items.length), h=o.h||1.15;
  items.forEach((it,i)=>{
    const c=cells[i], dark=!!it.dark;
    const dc = it.acc ? s._C.acc : s._C.dark;
    box(s,{ x:c.x, y:o.y, w:c.w, h, fill:dark?dc:K.w, line:dark?dc:K.g4, what:'KPI' });
    const p=T.pad, valSz=Math.min(it.size||T.kpi, Math.floor((c.w-2*p)*72/(0.62*String(it.v).length)));
    txt(s,it.v,{ x:c.x+p, y:o.y+p*0.7, w:c.w-2*p, h:h-2*p-0.30, sz:valSz, b:true,
      c:dark?K.w:K.ink, valign:'bottom', lh:1.1, wrap:false });
    txt(s,it.l,{ x:c.x+p, y:o.y+h-p-0.28, w:c.w-2*p, h:0.28, sz:T.kpiLabel, c:dark?K.g5:K.g2 });
  });
}
/** 표 — 헤더 먹, 본문 교차 음영, 세로줄 없음 */
function table(s, o){
  const T=s._T, colW=o.colW, rowH=o.rowH||T.rowH, headH=o.headH||rowH;
  const total=+colW.reduce((a,b)=>a+b,0).toFixed(4), h=headH+rowH*o.rows.length;
  guard(s,o.y,h,'표');
  if (o.x + total > W - M + 0.005)
    throw new Error(`표 폭 합계 ${total}"가 우측 여백을 넘음 — x=${o.x}에서 쓸 수 있는 폭은 ${+(W-M-o.x).toFixed(4)}"`);
  o.rows.forEach((r,i)=>{ if (r.length!==colW.length)
    throw new Error(`표 ${i+1}행의 칸 수(${r.length})가 열 수(${colW.length})와 다름`); });
  s.addShape('rect',{ x:o.x, y:o.y, w:total, h:headH, fill:{color:s._C.dark}, line:{type:'none'} });
  let cxp=o.x;
  o.head.forEach((t,i)=>{
    txt(s,t,{ x:cxp+0.08, y:o.y, w:colW[i]-0.16, h:headH, sz:T.tableHead, b:true, c:K.w,
      valign:'middle', align:'center', lh:1.2 });        // 표 머리글은 항상 가운데 정렬
    cxp+=colW[i];
  });
  o.rows.forEach((r,ri)=>{
    const y=o.y+headH+ri*rowH;
    if (ri%2===1) s.addShape('rect',{ x:o.x, y, w:total, h:rowH, fill:{color:K.g7}, line:{type:'none'} });
    if (o.rules) hr(s,{ x:o.x, y:y+rowH, w:total, color:K.g5 });   // 행 구분선은 기본 없음
    let cp=o.x;
    r.forEach((t,i)=>{
      const bold=o.boldCol && o.boldCol.includes(i);
      txt(s,String(t),{ x:cp+0.08, y, w:colW[i]-0.16, h:rowH, sz:T.tableBody, b:bold,
        c:t==='—'?K.g3:K.ink, valign:'middle', align:(o.align&&o.align[i])||'left', lh:1.2 });
      cp+=colW[i];
    });
  });
  hr(s,{ x:o.x, y:o.y+h, w:total, color:K.g4 });                  // 표 하단 마감선만 유지
  return { h, bottom:o.y+h };
}
/**
 * 네이티브 표 — PowerPoint의 진짜 표(<a:tbl>)로 만든다.
 * 도형 표(table)와 API는 같고, 차이는 산출물의 성격이다.
 *   · 네이티브: PowerPoint에서 행 추가·삭제·정렬·셀 병합 가능, 글이 길면 행 높이가 자동으로 늘어남
 *   · 도형    : 좌표를 1/1000인치까지 통제, 대신 편집은 도형 단위
 * rowH는 최소 높이이며 글이 길면 PowerPoint가 행을 키우므로, 본문 하단 벽에 여유를 두고 쓴다.
 */
function tableNative(s, o){
  const T=s._T, colW=o.colW, rowH=o.rowH||T.rowH;
  const total=+colW.reduce((a,b)=>a+b,0).toFixed(4);
  // 셀 글이 열 폭을 넘으면 PowerPoint가 행을 늘린다 → 늘어난 높이를 미리 계산해 벽을 지킨다
  const rowLines = (r, sz) => Math.max(1, ...r.map((t,i)=> lines(String(t), sz, colW[i]-0.16)));
  const grow = (n, sz) => +(n*sz*1.25/72 + 0.10).toFixed(3);
  const headH = Math.max(o.headH||rowH, grow(rowLines(o.head, T.tableHead), T.tableHead));
  const rowHs = o.rows.map(r => Math.max(rowH, grow(rowLines(r, T.tableBody), T.tableBody)));
  const tall = +(headH + rowHs.reduce((a,b)=>a+b,0)).toFixed(3);
  guard(s,o.y, tall, `표(slide ${s._page})`);
  if (o.x + total > W - M + 0.005)
    throw new Error(`표 폭 합계 ${total}"가 우측 여백을 넘음 — x=${o.x}에서 쓸 수 있는 폭은 ${+(W-M-o.x).toFixed(4)}"`);
  o.rows.forEach((r,i)=>{ if (r.length!==colW.length)
    throw new Error(`표 ${i+1}행의 칸 수(${r.length})가 열 수(${colW.length})와 다름`); });
  const hair = c => ({ type:'solid', color:c, pt:0.75 });
  const none = { type:'none' };
  const headRow = o.head.map((t)=>({ text:String(t), options:{
    fill:{ color:s._C.dark }, color:K.w, bold:true, fontSize:T.tableHead,
    align:'center', valign:'middle',              // 표 머리글은 항상 가운데 정렬
    border:[none,none,hair(s._C.dark),none] } }));
  const bodyRows = o.rows.map((r,ri)=> r.map((t,i)=>({ text:String(t), options:{
    fill:{ color: ri%2 ? K.g7 : K.w }, color: t==='—'?K.g3:K.ink,
    bold: !!(o.boldCol && o.boldCol.includes(i)), fontSize:T.tableBody,
    align:(o.align&&o.align[i])||'left', valign:'middle',
    border:[none,none,hair(ri===o.rows.length-1?K.g3:K.g5),none] } })));
  s.addTable([headRow, ...bodyRows], {
    x:o.x, y:o.y, w:total, colW, rowH:[headH, ...rowHs],
    fontFace:FONT, margin:[2,5,2,5], autoPage:false, border:none,
  });
  return { h:tall, bottom:+(o.y+tall).toFixed(3), grown: tall > headH + rowH*o.rows.length + 0.005 };
}

function bullets(s, items, o){
  const T=s._T, sz=o.sz||T.body, lh=o.lh||T.lh, space=(o.space==null?4:o.space)/72;
  // 글이 상자를 넘치면 아래 요소를 덮는다 — 미리 계산해 막는다
  const need = items.reduce((a,t)=>{
    const str = typeof t==='string' ? t : t.text;
    return a + lines(str, sz, o.w-0.12) * sz * lh / 72;
  }, 0) + (items.length-1)*space + 0.06;
  if (o.h && need > o.h + 0.02)
    throw new Error(`불릿 ${items.length}개가 높이 ${o.h}"를 넘는다(필요 ${need.toFixed(2)}") — `
      + `높이를 키우거나 문장을 줄일 것`);
  const runs=items.map((t,i)=>({ text:typeof t==='string'?t:t.text,
    options:{ bullet:{ code:'2013' }, breakLine:i<items.length-1, ...(typeof t==='object'&&t.b?{bold:true}:{}) } }));
  s.addText(runs, { x:o.x, y:o.y, w:o.w, h:o.h, fontFace:FONT, fontSize:o.sz||T.body, color:K.ink,
    valign:'top', margin:0, lineSpacingMultiple:o.lh||T.lh, paraSpaceAfter:o.space==null?4:o.space });
}
/* ── 차트 ────────────────────────────────────────────────
 * 수치는 표로 늘어놓기보다 차트로 보이는 편이 빠르다. 다만 이 체계는 무채색이므로
 * **계열을 명도로만 구분한다.** 색이 없으니 계열이 다섯을 넘으면 읽히지 않아 막는다.
 *
 * 공통 규칙
 *   - 격자선을 그리지 않는다. 값은 데이터 레이블로 직접 붙인다
 *   - 축선은 헤어라인(#C7C7C7), 축 글자는 cap 크기
 *   - 계열이 하나면 범례를 두지 않는다
 *   - hi 로 지정한 계열·항목만 먹(또는 강조색)으로 반전한다
 *
 * 차트를 만든 뒤에는 **반드시 scripts/postprocess.py 를 돌린다.**
 * pptxgenjs가 선언되지 않은 <c:axId>를 쓰고 한글 폰트(<a:ea>)를 빠뜨려
 * PowerPoint가 복구 대화상자를 띄운다.
 */
const SERIES_MAX = 5;
const seriesRamp = C => [C.dark, C.mid, K.g3, K.g4, K.g5];

function chartFrame(s, o, what){
  guard(s, o.y, o.h, what);
  if (o.x + o.w > W - M + 0.005)
    throw new Error(`${what} 우측 끝 ${(o.x+o.w).toFixed(3)}"가 여백을 넘음 — 폭은 ${(W-M-o.x).toFixed(3)}" 이하`);
}
/** 축·레이블·범례 공통 서식 */
function chartOpts(s, o, n){
  const T = s._T;
  return {
    x:o.x, y:o.y, w:o.w, h:o.h,
    chartColors: o.colors || seriesRamp(s._C).slice(0, n),
    showLegend: o.legend == null ? n > 1 : o.legend,
    legendPos: o.legendPos || 'b', legendFontFace: FONT, legendFontSize: T.cap, legendColor: K.g2,
    showValue: o.value == null ? true : o.value,
    dataLabelFontFace: FONT, dataLabelFontSize: o.labelSize || T.cap, dataLabelColor: K.g1,
    dataLabelPosition: o.labelPos || undefined,
    ...(o.fmt ? { dataLabelFormatCode: o.fmt } : {}),
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: T.cap, catAxisLabelColor: K.g2,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: T.cap, valAxisLabelColor: K.g2,
    catAxisLineColor: K.g4, valAxisLineColor: K.g4,
    valGridLine: { style:'none' }, catGridLine: { style:'none' },
    valAxisHidden: o.valAxis == null ? true : !o.valAxis,   // 값은 레이블로 읽는다
    border: { pt:0, color:K.w }, fill: K.w,
    ...(o.max != null ? { valAxisMaxVal:o.max } : {}),
    ...(o.min != null ? { valAxisMinVal:o.min } : {}),
    ...(o.title ? { showTitle:true, title:o.title, titleFontFace:FONT,
                    titleFontSize:s._T.h2, titleColor:K.ink, titleAlign:'left' } : {}),
  };
}
function toSeries(o){
  const ss = o.series || [{ name: o.name || '값', values: o.values }];
  if (ss.length > SERIES_MAX)
    throw new Error(`차트 계열이 ${ss.length}개 — 무채색은 명도로만 구분하므로 ${SERIES_MAX}개까지만 쓴다. `
      + `계열을 묶거나 차트를 나눌 것`);
  return ss.map(x => ({ name:x.name, labels:o.cats, values:x.values }));
}

/** 막대 — horizontal:true 면 가로막대. stacked:true 면 누적 */
function barChart(s, o){
  chartFrame(s, o, '막대 차트');
  const data = toSeries(o);
  s.addChart('bar', data, { ...chartOpts(s, o, data.length),
    barDir: o.horizontal ? 'bar' : 'col',
    // 가로막대는 PowerPoint가 첫 항목을 맨 아래에 놓는다 — 순위 차트가 뒤집히므로 되돌린다
    ...(o.horizontal ? { catAxisOrientation:'maxMin' } : {}),
    barGrouping: o.stacked ? 'stacked' : 'clustered',
    barGapWidthPct: o.gap == null ? 55 : o.gap,
    dataLabelPosition: o.stacked ? 'ctr' : (o.horizontal ? 'outEnd' : 'outEnd'),
  });
  return { bottom:+(o.y+o.h).toFixed(3) };
}
/** 선 — 추세. marker:false 면 표식 없음 */
function lineChart(s, o){
  chartFrame(s, o, '선 차트');
  const data = toSeries(o);
  s.addChart('line', data, { ...chartOpts(s, o, data.length),
    lineSize: o.lineSize || 2.25,
    lineDataSymbol: o.marker === false ? 'none' : (o.marker || 'circle'),
    lineDataSymbolSize: o.markerSize || 6,
    valAxisHidden: o.valAxis == null ? false : !o.valAxis,   // 추세는 축이 있어야 읽힌다
    showValue: o.value == null ? false : o.value,
  });
  return { bottom:+(o.y+o.h).toFixed(3) };
}
/** 원 — hole 을 주면 도넛. 조각이 다섯을 넘으면 '기타'로 묶는다.
 *  원은 상자의 **짧은 변**에 갇히므로 정사각형에 가까운 상자에 넣는다.
 *  가로로 긴 상자에 넣을 때는 범례를 오른쪽에 두어(기본값) 원이 왼쪽 정사각형을 쓰게 한다. */
function pieChart(s, o){
  chartFrame(s, o, '원 차트');
  if (o.labels.length > SERIES_MAX)
    throw new Error(`원 차트 조각이 ${o.labels.length}개 — 명도로 구분하려면 ${SERIES_MAX}개까지. `
      + `작은 조각은 '기타'로 묶을 것`);
  const data = [{ name:o.name || '구성', labels:o.labels, values:o.values }];
  const pct = o.percent == null ? true : o.percent;   // 구성비가 기본 — 값과 비율을 함께 찍지 않는다
  s.addChart(o.hole ? 'doughnut' : 'pie', data, { ...chartOpts(s, o, o.labels.length),
    holeSize: o.hole || undefined,
    // 조각 위에 찍으면 먹 조각에서 글자가 묻힌다 — 바깥에 둔다(도넛은 가운데가 비어 안쪽이 안전)
    dataLabelPosition: o.hole ? 'ctr' : (o.labelPos || 'outEnd'),
    showPercent: pct, showValue: pct ? false : (o.value == null ? true : o.value),
    showLegend: o.legend == null ? true : o.legend,
    legendPos: o.legendPos || 'r',
    valAxisHidden: true, catAxisHidden: true,
  });
  return { bottom:+(o.y+o.h).toFixed(3) };
}
/** 혼합 — 막대(왼쪽 축) + 선(오른쪽 축). 규모와 비율을 한 장에 겹칠 때 */
function comboChart(s, o){
  chartFrame(s, o, '혼합 차트');
  const T = s._T, ramp = seriesRamp(s._C);
  const bars  = (o.bars  || []).map((x,i) => ({ name:x.name, labels:o.cats, values:x.values }));
  const lines = (o.lines || []).map((x,i) => ({ name:x.name, labels:o.cats, values:x.values }));
  if (bars.length + lines.length > SERIES_MAX)
    throw new Error(`혼합 차트 계열이 ${bars.length+lines.length}개 — ${SERIES_MAX}개까지만 쓴다`);
  const common = chartOpts(s, o, bars.length + lines.length);
  s.addChart([
    { type:'bar',  data:bars,  options:{ barDir:'col', barGrouping:'clustered',
        barGapWidthPct: o.gap == null ? 55 : o.gap,
        chartColors: ramp.slice(0, bars.length), showValue:true,
        dataLabelFontFace:FONT, dataLabelFontSize:T.cap, dataLabelColor:K.g1, dataLabelPosition:'outEnd' } },
    { type:'line', data:lines, options:{ chartColors: ramp.slice(bars.length, bars.length+lines.length),
        lineSize:2.25, lineDataSymbol:'circle', lineDataSymbolSize:6, secondaryValAxis:!!o.rightAxis,
        secondaryCatAxis:!!o.rightAxis, showValue:false } },
  ], { ...common, valAxes:undefined, catAxes:undefined,
       valAxisHidden:false, valAxisMajorTickMark:'none',
       ...(o.rightAxis ? { valAxes:[
            { valAxisTitle:o.unit || '', showValAxisTitle:!!o.unit, valAxisLabelFontFace:FONT,
              valAxisLabelFontSize:T.cap, valAxisLabelColor:K.g2, valAxisLineColor:K.g4, valGridLine:{style:'none'} },
            { valAxisTitle:o.unit2 || '', showValAxisTitle:!!o.unit2, valAxisLabelFontFace:FONT,
              valAxisLabelFontSize:T.cap, valAxisLabelColor:K.g2, valAxisLineColor:K.g4, valGridLine:{style:'none'} },
          ], catAxes:[
            { catAxisLabelFontFace:FONT, catAxisLabelFontSize:T.cap, catAxisLabelColor:K.g2, catAxisLineColor:K.g4 },
            { catAxisHidden:true },
          ] } : {}) });
  return { bottom:+(o.y+o.h).toFixed(3) };
}
/** 차트 아래 단위·출처 한 줄 */
const chartNote = (s, t, o) => footnote(s, t, { x:o.x, y:o.y, w:o.w });

/** 배경 명도에 맞는 글자색 — 밝은 회색 위에 흰 글자를 얹지 않는다 */
function onTone(hex){
  const r=parseInt(hex.slice(0,2),16), g=parseInt(hex.slice(2,4),16), b=parseInt(hex.slice(4,6),16);
  return (r*0.299 + g*0.587 + b*0.114) < 140 ? K.w : K.ink;
}

/* ── 관계 도식 ───────────────────────────────────────────
 * 모두 무채색이며 그라데이션·3D·베벨을 쓰지 않는다. 겹침과 명도만으로 뜻을 만든다.
 * 각 함수는 { bottom }을 돌려주므로 그 값으로 다음 요소의 y를 잡는다.
 */

/** 벤 — 집합·그룹·공통. 원을 반투명으로 겹쳐 교집합이 저절로 진해지게 한다.
 *   items: ['영역 A','영역 B','영역 C']  (2개 또는 3개)
 *   center: 교집합에 넣을 말 (선택)
 */
function venn(s, o){
  const T = s._T, items = o.items;
  if (items.length < 2 || items.length > 3) throw new Error('벤 다이어그램은 2개 또는 3개 영역만');
  guard(s, o.y, o.h, '벤');
  const r = o.r || Math.min(o.h*0.36, o.w*0.22);
  const cxm = o.x + o.w/2, cym = o.y + o.h*0.47, d = r*0.62;
  const at = items.length === 3 ? [-90, 30, 150] : [180, 0];
  const pos = at.map(a => ({ x:cxm + d*Math.cos(a*Math.PI/180) * (items.length===2?0.85:1),
                             y:cym + d*Math.sin(a*Math.PI/180) }));
  pos.forEach(pt => s.addShape('ellipse', { x:+(pt.x-r).toFixed(4), y:+(pt.y-r).toFixed(4), w:r*2, h:r*2,
    fill:{ color:K.ink, transparency:80 }, line:{ color:K.g3, width:0.75 } }));
  // 영역 이름은 원 바깥으로 빼서 교집합 글자와 겹치지 않게 한다
  items.forEach((t, i) => {
    const a = at[i]*Math.PI/180, off = d + r*1.16;   // 원 테두리 밖으로 밀어 겹치지 않게 한다
    const lx = cxm + off*Math.cos(a), ly = cym + off*Math.sin(a);
    const w = 1.9;
    txt(s, t, { x:+(lx - w/2).toFixed(4), y:+(ly - 0.16).toFixed(4), w, h:0.32,
      sz:o.sz||T.body, b:true, align:'center', valign:'middle' });
  });
  if (o.center)
    txt(s, o.center, { x:+(cxm-1.0).toFixed(4), y:+(cym-0.16).toFixed(4), w:2.0, h:0.32,
      sz:o.sz||T.body, b:true, c:K.w, align:'center', valign:'middle' });
  return { bottom:+(o.y+o.h).toFixed(3) };
}

/** 허브&스포크 — 분석. 가운데 요인 하나에 주변 요소를 붙인다.
 *   hub: '중심', items: ['요소1', … ]  (3~8개)
 */
function hubSpoke(s, o){
  const T = s._T, n = o.items.length;
  if (n < 3 || n > 8) throw new Error('허브&스포크는 요소 3~8개');
  guard(s, o.y, o.h, '허브&스포크');
  const cxm = o.x + o.w/2, cym = o.y + o.h/2;
  const R = Math.min(o.h*0.33, o.w*0.20);          // 위성 중심까지 거리
  const hr = o.hubR || Math.min(0.48, R*0.62);     // 허브 반지름
  const sr = o.spokeR || hr*0.78;                  // 위성 반지름
  for (let i=0;i<n;i++){
    const a = (-90 + i*360/n) * Math.PI/180;
    const px = cxm + R*Math.cos(a), py = cym + R*Math.sin(a);
    const L = Math.hypot(px-cxm, py-cym) - hr - sr;
    if (L > 0.02){                                  // 허브와 위성을 잇는 선
      const mx = cxm + (hr + L/2)*Math.cos(a), my = cym + (hr + L/2)*Math.sin(a);
      s.addShape('line', { x:+(mx-L/2).toFixed(4), y:+my.toFixed(4), w:+L.toFixed(4), h:0,
        rotate:+( -90 + i*360/n ).toFixed(2), line:{ color:K.g3, width:0.75 } });
    }
    s.addShape('ellipse', { x:+(px-sr).toFixed(4), y:+(py-sr).toFixed(4), w:sr*2, h:sr*2,
      fill:{ color:K.g6 }, line:{ color:K.g3, width:0.75 } });
    txt(s, o.items[i], { x:+(px-sr).toFixed(4), y:+(py-sr).toFixed(4), w:sr*2, h:sr*2,
      sz:o.sz||T.small, align:'center', valign:'middle', lh:1.15 });
  }
  s.addShape('ellipse', { x:+(cxm-hr).toFixed(4), y:+(cym-hr).toFixed(4), w:hr*2, h:hr*2,
    fill:{ color:s._C.dark }, line:{ type:'none' } });
  txt(s, o.hub, { x:+(cxm-hr).toFixed(4), y:+(cym-hr).toFixed(4), w:hr*2, h:hr*2,
    sz:o.hubSz||T.body, b:true, c:K.w, align:'center', valign:'middle', lh:1.15 });
  return { bottom:+(o.y+o.h).toFixed(3) };
}

/** 순환 — 반복되는 고리. 원 위에 단계를 놓고 이웃끼리 화살표로 잇는다.
 *   items: ['단계1', … ]  (3~6개)
 */
function cycle(s, o){
  const T = s._T, n = o.items.length;
  if (n < 3 || n > 6) throw new Error('순환 도식은 단계 3~6개');
  guard(s, o.y, o.h, '순환');
  const cxm = o.x + o.w/2, cym = o.y + o.h/2;
  const R = Math.min(o.h*0.34, o.w*0.21);
  const nr = o.nodeR || Math.min(0.56, R*0.70);
  const pts = Array.from({length:n}, (_,i) => {
    const a = (-90 + i*360/n) * Math.PI/180;
    return { x:cxm + R*Math.cos(a), y:cym + R*Math.sin(a) };
  });
  for (let i=0;i<n;i++){                            // 이웃을 잇는 화살표
    const a = pts[i], b = pts[(i+1)%n];
    const dx = b.x-a.x, dy = b.y-a.y, dist = Math.hypot(dx,dy);
    const L = dist - nr*2 - 0.10;
    if (L <= 0.05) continue;
    const mx = (a.x+b.x)/2, my = (a.y+b.y)/2, ah = o.arrowH || 0.16;
    s.addShape('rightArrow', { x:+(mx-L/2).toFixed(4), y:+(my-ah/2).toFixed(4), w:+L.toFixed(4), h:ah,
      rotate:+(Math.atan2(dy,dx)*180/Math.PI).toFixed(2),
      fill:{ color:K.g4 }, line:{ type:'none' } });
  }
  pts.forEach((pt,i) => {
    const dark = o.hi === i;
    s.addShape('ellipse', { x:+(pt.x-nr).toFixed(4), y:+(pt.y-nr).toFixed(4), w:nr*2, h:nr*2,
      fill:{ color:dark ? s._C.dark : K.g6 }, line:{ color:dark ? s._C.dark : K.g3, width:0.75 } });
    txt(s, o.items[i], { x:+(pt.x-nr).toFixed(4), y:+(pt.y-nr).toFixed(4), w:nr*2, h:nr*2,
      sz:o.sz||T.small, b:true, c:dark?K.w:K.ink, align:'center', valign:'middle', lh:1.15 });
  });
  if (o.center)
    txt(s, o.center, { x:+(cxm-R*0.62).toFixed(4), y:+(cym-0.20).toFixed(4), w:R*1.24, h:0.40,
      sz:o.centerSz||T.body, b:true, c:K.g2, align:'center', valign:'middle' });
  return { bottom:+(o.y+o.h).toFixed(3) };
}

/** 피라미드 — 기반·기초·삼위일체. 위로 갈수록 좁아지고 진해진다.
 *   items: ['최상위','중간','기반']  (위→아래, 2~5층)
 *   notes: 층별 오른쪽 설명 (선택)
 */
function pyramid(s, o){
  const T = s._T, n = o.items.length;
  if (n < 2 || n > 5) throw new Error('피라미드는 2~5층');
  guard(s, o.y, o.h, '피라미드');
  const gap = o.gap == null ? 0.05 : o.gap;
  const lh = (o.h - gap*(n-1)) / n;
  const noteW = o.notes ? (o.noteW || 3.2) : 0;
  const pw = o.w - (noteW ? noteW + 0.22 : 0);       // 피라미드 폭
  const cxm = o.x + pw/2;
  const tones = [s._C.dark, K.g2, K.g3, K.g4, K.g5];
  let y = o.y;
  for (let i=0;i<n;i++){
    const topW = pw * i / n, botW = pw * (i+1) / n;
    const tone = tones[Math.min(i, tones.length-1)];
    const fg = onTone(tone);
    if (i === 0) {
      s.addShape('triangle', { x:+(cxm-botW/2).toFixed(4), y:+y.toFixed(4), w:+botW.toFixed(4), h:+lh.toFixed(4),
        fill:{ color:tone }, line:{ type:'none' } });
    } else {
      // pptxgenjs는 도형의 조절점(adj)을 넘기지 못한다. 도형 이름에 값을 실어 보내고
      // postprocess.py 가 실제 adj로 바꿔 넣는다(표지 평행사변형과 같은 방식).
      const inset = (botW - topW)/2;
      const adj = Math.round(inset / Math.min(botW, lh) * 100000);
      s.addShape('trapezoid', { x:+(cxm-botW/2).toFixed(4), y:+y.toFixed(4), w:+botW.toFixed(4), h:+lh.toFixed(4),
        fill:{ color:tone }, line:{ type:'none' }, objectName:`KSA_TRAP_${adj}` });
    }
    const tw = i === 0 ? botW*0.62 : (topW + botW)/2;
    txt(s, o.items[i], { x:+(cxm-tw/2).toFixed(4), y:+(y + lh*(i===0?0.42:0.12)).toFixed(4),
      w:+tw.toFixed(4), h:+(lh*(i===0?0.55:0.76)).toFixed(4),
      sz:o.sz||T.body, b:true, c:fg, align:'center', valign:'middle', lh:1.2 });
    if (o.notes && o.notes[i])
      txt(s, o.notes[i], { x:+(o.x + pw + 0.22).toFixed(4), y:+y.toFixed(4), w:noteW, h:+lh.toFixed(4),
        sz:o.noteSz||T.small, c:K.g1, valign:'middle', lh:1.3 });
    y += lh + gap;
  }
  return { bottom:+(o.y+o.h).toFixed(3) };
}

/** 계단 — 발전·성장. 왼쪽에서 오른쪽으로 올라가며 진해진다.
 *   items: [{t:'단계', d:'설명', v:'수치'}, … ]  (3~6개)
 */
function steps(s, o){
  const T = s._T, n = o.items.length;
  if (n < 3 || n > 6) throw new Error('계단 도식은 3~6단');
  guard(s, o.y, o.h, '계단');
  const gap = o.gap == null ? 0.10 : o.gap;
  const sw = (o.w - gap*(n-1)) / n;
  const base = o.y + o.h;
  // 단이 몇이든 마지막이 가장 진하게 오도록 뒤에서 n개를 쓴다
  const ramp = [K.g5, K.g4, K.g3, K.g2, s._C.dark];
  const tones = ramp.slice(Math.max(0, ramp.length - n));
  o.items.forEach((it, i) => {
    const h = o.h * (0.34 + 0.66*(i+1)/n);
    const x = o.x + i*(sw+gap), y = base - h;
    const tone = tones[i], fg = onTone(tone);
    s.addShape('rect', { x:+x.toFixed(4), y:+y.toFixed(4), w:+sw.toFixed(4), h:+h.toFixed(4),
      fill:{ color:tone }, line:{ type:'none' } });
    txt(s, it.t, { x:+x.toFixed(4), y:+(y+0.10).toFixed(4), w:+sw.toFixed(4), h:0.30,
      sz:o.sz||T.body, b:true, c:fg, align:'center', valign:'middle' });
    if (it.v) txt(s, it.v, { x:+x.toFixed(4), y:+(y+0.42).toFixed(4), w:+sw.toFixed(4), h:0.30,
      sz:o.vSz||T.h2, b:true, c:fg, align:'center', valign:'middle' });
    if (it.d) txt(s, it.d, { x:+x.toFixed(4), y:+(base+0.06).toFixed(4), w:+sw.toFixed(4), h:0.44,
      sz:o.dSz||T.small, c:K.g2, align:'center', valign:'top', lh:1.25 });
  });
  return { bottom:+(base + (o.items.some(i=>i.d) ? 0.50 : 0)).toFixed(3) };
}

/** 하비볼 한 개 — 지름 d, 값 v(0~4). 채운 몫만큼 시계 방향으로 칠한다.
 *  글리프(◔ ◕)는 맑은 고딕에 없어 다른 서체로 대체되며 작게 깨진다. 그래서 도형으로 그린다.
 *  부채꼴의 각도는 pptxgenjs가 넘기지 못하므로 도형 이름에 실어 postprocess.py 가 넣는다.
 */
function harveyBall(s, o){
  const d = o.d, v = o.v, c = o.color || K.ink, line = o.line || K.g3;
  const x = +(o.x - d/2).toFixed(4), y = +(o.y - d/2).toFixed(4);
  if (v >= 4) {
    s.addShape('ellipse', { x, y, w:d, h:d, fill:{ color:c }, line:{ color:c, width:0.75 } });
    return;
  }
  // 부채꼴은 adj1=0(3시)에서 시계 방향으로 칠하고, 도형을 돌려 시작 위치를 맞춘다.
  //   rotate -90 → 12–3시,  0 → 3–6시,  90 → 6–9시,  180 → 9–12시
  // 270°짜리 부채꼴은 렌더러가 제대로 그리지 못하므로 3/4은 반원 + 사분원으로 나눠 그린다.
  const wedge = (deg, rot) => s.addShape('pie', { x, y, w:d, h:d, rotate:rot,
    fill:{ color:c }, line:{ type:'none' }, objectName:`KSA_PIE_${deg*60000}` });
  if (v === 1) wedge(90, -90);            // 12–3시
  else if (v === 2) wedge(180, -90);      // 오른쪽 반
  else if (v === 3) { wedge(180, -90); wedge(90, 90); }   // 오른쪽 반 + 6–9시
  s.addShape('ellipse', { x, y, w:d, h:d, fill:{ type:'none' }, line:{ color:line, width:0.75 } });
}

/** 하비볼 평가표 — 비교평가. 명도가 아니라 **채운 몫**으로 5단계를 읽힌다.
 *   rows: [{ t:'항목', v:[0,2,4, … ] }]  v는 0~4
 *   cols: ['대안 A','대안 B', … ]
 *  값 칸은 네이티브 표에 비워 두고 그 위에 공을 얹는다. 항목 이름이 두 줄이 되면
 *  행이 자라 공의 위치가 어긋나므로 한 줄을 넘기지 못하게 막는다.
 */
const HARVEY_STEPS = ['없음','낮음','보통','높음','매우 높음'];
function harvey(s, o){
  const T = s._T;
  const n = o.cols.length;
  const first = o.firstW || 3.0;
  const rest = +((o.w - first) / n).toFixed(4);
  const colW = o.colW || [first, ...Array(n).fill(rest)];
  const headH = o.headH || Math.max(0.40, T.rowH);
  const rowH = o.rowH || Math.max(0.46, T.rowH + 0.12);   // 공이 들어갈 만큼 높이를 준다
  o.rows.forEach(r => {
    if (r.v.length !== n) throw new Error(`하비볼 '${r.t}'의 값 ${r.v.length}개가 열 ${n}개와 다름`);
    r.v.forEach(v => { if (!Number.isInteger(v) || v < 0 || v > 4)
      throw new Error(`하비볼 값은 0~4 정수 (받은 값 ${v})`); });
    if (lines(r.t, T.tableBody, colW[0] - 0.16) > 1)
      throw new Error(`하비볼 항목 '${r.t}'이 한 줄을 넘는다 — 공 위치가 어긋나므로 짧게 쓸 것`);
  });
  const res = tableNative(s, { x:o.x, y:o.y, w:o.w, colW, headH, rowH,
    head: [o.corner || '평가 항목', ...o.cols],
    align: ['left', ...Array(n).fill('center')],
    rows: o.rows.map(r => [r.t, ...Array(n).fill('')]) });
  if (res.grown) throw new Error('하비볼 표의 행이 자랐다 — 열 폭을 넓히거나 이름을 줄일 것');
  const d = o.ball || Math.min(0.22, rowH * 0.44);
  o.rows.forEach((r, i) => {
    const cy = o.y + headH + rowH*i + rowH/2;
    r.v.forEach((v, j) => {
      const cx = o.x + colW.slice(0, j+1).reduce((p, q) => p+q, 0) + colW[j+1]/2;
      harveyBall(s, { x:+cx.toFixed(4), y:+cy.toFixed(4), d, v });
    });
  });
  // 범례도 글리프가 아니라 같은 공으로 그린다
  if (o.legend !== false) {
    const ly = +(res.bottom + 0.16).toFixed(4);
    let lx = o.x + d/2;
    HARVEY_STEPS.forEach((t, v) => {
      harveyBall(s, { x:+lx.toFixed(4), y:ly, d:d*0.86 });
      harveyBall(s, { x:+lx.toFixed(4), y:ly, d:d*0.86, v });
      const tw = textW(t, T.cap) + 0.10;
      txt(s, t, { x:+(lx + d*0.62).toFixed(4), y:+(ly - 0.11).toFixed(4), w:+tw.toFixed(4), h:0.22,
        sz:T.cap, c:K.g2, valign:'middle', wrap:false });
      lx += d*0.62 + tw + 0.22;
    });
    return { bottom:+(ly + 0.16).toFixed(3) };
  }
  return { bottom:res.bottom };
}

/** 원인 → 결과 — 전후 관계. 왼쪽 현상에서 오른쪽 결과로 잇는다.
 *   rows: [{ cause:'현상', effect:'결과', note:'매개' }]  (1~4행)
 */
function causeEffect(s, o){
  const T = s._T, n = o.rows.length;
  if (n < 1 || n > 4) throw new Error('원인→결과는 1~4행');
  guard(s, o.y, o.h, '원인→결과');
  const gap = o.gap == null ? 0.14 : o.gap;
  const rh = (o.h - 0.34 - gap*(n-1)) / n;          // 0.34 = 머리 밴드
  const aw = o.arrowW || 0.62;
  const cw2 = (o.w - aw - 0.28) / 2;
  const lx = o.x, rx = o.x + cw2 + aw + 0.28;
  [[lx, o.causeLabel || '현상', K.g2], [rx, o.effectLabel || '결과', s._C.dark]]
    .forEach(([x, t, c]) => {
      s.addShape('rect', { x:+x.toFixed(4), y:+o.y.toFixed(4), w:+cw2.toFixed(4), h:0.30,
        fill:{ color:c }, line:{ type:'none' } });
      txt(s, t, { x:+x.toFixed(4), y:+o.y.toFixed(4), w:+cw2.toFixed(4), h:0.30,
        sz:o.headSz||T.h2, b:true, c:K.w, align:'center', valign:'middle' });
    });
  let y = o.y + 0.34;
  o.rows.forEach((r) => {
    box(s, { x:lx, y:+y.toFixed(4), w:+cw2.toFixed(4), h:+rh.toFixed(4), fill:K.w, line:K.g4, what:'현상' });
    txt(s, r.cause, { x:+(lx+T.pad).toFixed(4), y:+y.toFixed(4), w:+(cw2-2*T.pad).toFixed(4), h:+rh.toFixed(4),
      sz:o.sz||T.body, valign:'middle', lh:1.3 });
    s.addShape('rightArrow', { x:+(lx+cw2+0.14).toFixed(4), y:+(y + rh/2 - 0.13).toFixed(4), w:aw, h:0.26,
      fill:{ color:K.g4 }, line:{ type:'none' } });
    if (r.note) txt(s, r.note, { x:+(lx+cw2+0.14).toFixed(4), y:+(y + rh/2 - 0.44).toFixed(4), w:aw, h:0.28,
      sz:Math.max(8.5, T.cap), c:K.g2, align:'center', valign:'middle', wrap:false });
    box(s, { x:rx, y:+y.toFixed(4), w:+cw2.toFixed(4), h:+rh.toFixed(4), fill:K.g7, line:K.g4, what:'결과' });
    txt(s, r.effect, { x:+(rx+T.pad).toFixed(4), y:+y.toFixed(4), w:+(cw2-2*T.pad).toFixed(4), h:+rh.toFixed(4),
      sz:o.sz||T.body, b:true, valign:'middle', lh:1.3 });
    y += rh + gap;
  });
  return { bottom:+(o.y+o.h).toFixed(3) };
}

function chevrons(s, steps, o){
  const T=s._T, n=steps.length, step=(o.w)/n, chW=step+0.30, h=o.h||0.80;
  steps.forEach((st,i)=>{
    const x=o.x+i*(step-0.05);
    s.addShape(i===0?'homePlate':'chevron',{ x, y:o.y, w:chW, h,
      fill:{ color: st.tone==='light'?K.g4 : st.tone==='mid'?s._C.mid : (st.tone==='acc'? s._C.acc : s._C.dark) }, line:{ color:K.w, width:1 } });
    txt(s,st.t,{ x:x+(i?0.30:0.14), y:o.y, w:chW-0.55, h, sz:o.sz||T.body, b:true,
      c: st.tone==='light'?K.ink:K.w, align:'center', valign:'middle', wrap:false });
    if (st.d) txt(s,st.d,{ x:x+(i?0.26:0.10), y:o.y+h+0.10, w:step-0.10, h:0.46, sz:T.small, c:K.g2, lh:1.25 });
  });
}
function waterfall(s, o){
  const T=s._T, steps=o.steps, maxV=o.max, plotH=o.h;
  const gap=o.w/steps.length, bw=gap*0.54, sc=plotH/maxV;
  let run=0;
  steps.forEach((st,i)=>{
    const bx=o.x+i*gap+(gap-bw)/2; let top,hh,fill,ln;
    if (st.base){ top=o.y+plotH-st.v*sc; hh=st.v*sc; run=st.v; fill=st.goal? s._C.acc : s._C.dark; ln=fill; }
    else { const from=run, to=run+st.v; run=to; top=o.y+plotH-Math.max(from,to)*sc; hh=Math.abs(st.v)*sc; fill=K.g5; ln=K.g2; }
    s.addShape('rect',{ x:bx, y:top, w:bw, h:hh, fill:{color:fill}, line:{color:ln,width:0.75} });
    txt(s,String(st.v),{ x:bx-0.18, y:top-0.24, w:bw+0.36, h:0.22, sz:T.small, b:true, align:'center', valign:'bottom' });
    txt(s,st.t,{ x:bx-0.22, y:o.y+plotH+0.05, w:bw+0.44, h:0.26, sz:T.small, c:K.g2, align:'center', lh:1.15 });
  });
  hr(s,{ x:o.x, y:o.y+plotH, w:o.w, color:K.g3 });
  return o.y+plotH+0.31;                 // x축 라벨 하단
}
function tree(s, o){
  const T=s._T, r=o.root, mids=o.mids;
  s.addShape('rect',{ x:r.x, y:r.y, w:r.w, h:r.h, fill:{color:s._C.dark}, line:{type:'none'} });
  txt(s,r.t,{ x:r.x, y:r.y, w:r.w, h:r.h, sz:T.body, b:true, c:K.w, align:'center', valign:'middle', lh:1.2 });
  const stubX=r.x+r.w+0.22;
  hr(s,{ x:r.x+r.w, y:r.y+r.h/2, w:0.22, color:K.g3 });
  mids.forEach(m=>{
    hr(s,{ x:stubX, y:Math.min(m.y+m.h/2, r.y+r.h/2), w:0, color:K.g3 });
    s.addShape('line',{ x:stubX, y:Math.min(m.y+m.h/2,r.y+r.h/2), w:0, h:Math.abs(m.y+m.h/2-(r.y+r.h/2)), line:{color:K.g3,width:0.75} });
    hr(s,{ x:stubX, y:m.y+m.h/2, w:0.18, color:K.g3 });
    s.addShape('rect',{ x:m.x, y:m.y, w:m.w, h:m.h, fill:{color:K.g5}, line:{color:K.g2,width:0.75} });
    txt(s,m.t,{ x:m.x, y:m.y, w:m.w, h:m.h, sz:T.small, b:true, align:'center', valign:'middle', lh:1.2 });
    (m.leaves||[]).forEach(lf=>{
      s.addShape('rect',{ x:lf.x, y:lf.y, w:lf.w, h:lf.h, fill:{color:K.w}, line:{color:K.g4,width:0.75} });
      txt(s,lf.t,{ x:lf.x+0.08, y:lf.y, w:lf.w-0.16, h:lf.h, sz:T.small, valign:'middle', lh:1.15 });
      const sx=m.x+m.w;
      hr(s,{ x:sx, y:m.y+m.h/2, w:0.16, color:K.g3 });
      s.addShape('line',{ x:sx+0.16, y:Math.min(lf.y+lf.h/2,m.y+m.h/2), w:0, h:Math.abs(lf.y+lf.h/2-(m.y+m.h/2)), line:{color:K.g3,width:0.75} });
      hr(s,{ x:sx+0.16, y:lf.y+lf.h/2, w:lf.x-(sx+0.16), color:K.g3 });
    });
  });
}
function matrix(s, o){
  const T=s._T;
  s.addShape('rect',{ x:o.x, y:o.y, w:o.w, h:o.h, fill:{color:K.w}, line:{color:K.g4,width:0.75} });
  s.addShape('rect',{ x:o.x, y:o.y, w:o.w/2, h:o.h/2, fill:{color:K.g6}, line:{color:K.g4,width:0.75} });
  s.addShape('line',{ x:o.x+o.w/2, y:o.y, w:0, h:o.h, line:{color:K.g4,width:0.75} });
  hr(s,{ x:o.x, y:o.y+o.h/2, w:o.w, color:K.g4 });
  const q=o.quadrants;
  txt(s,q[0],{ x:o.x+0.10, y:o.y+0.08, w:o.w/2-0.2, h:0.24, sz:T.small, b:true });
  txt(s,q[1],{ x:o.x+o.w/2+0.10, y:o.y+0.08, w:o.w/2-0.2, h:0.24, sz:T.small, b:true, c:K.g2 });
  txt(s,q[2],{ x:o.x+0.10, y:o.y+o.h/2+0.08, w:o.w/2-0.2, h:0.24, sz:T.small, b:true, c:K.g2 });
  txt(s,q[3],{ x:o.x+o.w/2+0.10, y:o.y+o.h/2+0.08, w:o.w/2-0.2, h:0.24, sz:T.small, b:true, c:K.g2 });
  txt(s,o.yLabel,{ x:o.x-0.62, y:o.y, w:0.58, h:o.h, sz:T.small, c:K.g2, align:'center', valign:'middle', lh:1.25 });
  txt(s,o.xLabel,{ x:o.x, y:o.y+o.h+0.06, w:o.w, h:0.24, sz:T.small, c:K.g2, align:'center' });
  o.points.forEach(p=>{
    const px=o.x+p.px*o.w, py=o.y+p.py*o.h, d=p.d||0.34, hot=p.px<0.5&&p.py<0.5;
    s.addShape('ellipse',{ x:px-d/2, y:py-d/2, w:d, h:d, fill:{color:hot?s._C.acc:K.g4}, line:{color:hot?s._C.acc:K.g2,width:0.75} });
    // 오른쪽 끝 버블은 옆에 라벨을 둘 자리가 없다 → 버블 위에 가운데 정렬로 얹는다
    const room = (o.labelRoom == null ? 1.96 : o.labelRoom);
    const st = { sz:T.small, b:hot, c:hot?s._C.acc:K.g2, valign:'middle' };
    if (px + d/2 + 0.06 + textW(p.t, T.small) > o.x + o.w + room)
      txt(s, p.t, { x:px-1.0, y:py-d/2-0.26, w:2.0, h:0.24, ...st, align:'center' });
    else
      txt(s, p.t, { x:px+d/2+0.06, y:py-0.12, w:Math.min(1.9, textW(p.t,T.small)+0.12), h:0.24, ...st });
  });
}
function gantt(s, o){
  const T=s._T, labW=o.labW||1.6, noteW=o.noteW||1.8;
  const gx=o.x+labW, gw=o.w-labW-noteW, rowH=o.rowH||0.56;
  o.cols.forEach((c,i)=>{
    const x=gx+i*(gw/o.cols.length);
    txt(s,c,{ x, y:o.y, w:gw/o.cols.length, h:0.26, sz:T.small, c:K.g2, align:'center' });
    s.addShape('line',{ x, y:o.y+0.30, w:0, h:rowH*o.rows.length+0.06, line:{color:K.g5,width:0.75} });
  });
  s.addShape('line',{ x:gx+gw, y:o.y+0.30, w:0, h:rowH*o.rows.length+0.06, line:{color:K.g5,width:0.75} });
  hr(s,{ x:gx, y:o.y+0.30, w:gw, color:K.g3 });
  o.rows.forEach((r,i)=>{
    const y=o.y+0.36+i*rowH;
    txt(s,r.t,{ x:o.x, y:y+0.04, w:labW-0.10, h:0.30, sz:T.small, b:true, valign:'middle' });
    hr(s,{ x:gx, y:y+rowH-0.04, w:gw, color:K.g6 });
    s.addShape('roundRect',{ x:gx+r.a*gw, y:y+0.08, w:(r.b-r.a)*gw, h:rowH-0.24, rectRadius:0.04,
      fill:{ color: r.tone==='light'?K.g4 : r.tone==='mid'?s._C.mid : s._C.dark }, line:{type:'none'} });
    const mx=gx+r.b*gw;
    s.addShape('diamond',{ x:mx-0.085, y:y+(rowH-0.24)/2, w:0.17, h:0.17, fill:{color:K.w}, line:{color:s._C.acc,width:1.25} });
    if (r.ms) txt(s,r.ms,{ x:gx+gw+0.22, y:y+0.06, w:noteW-0.24, h:0.30, sz:T.small, c:K.g2, valign:'middle' });
  });
  return o.y+0.36+rowH*o.rows.length;
}
function layers(s, o){          // 3단 계층(교육체계 등)
  const T=s._T, n=o.items.length, h=o.h/n;
  o.items.forEach((it,i)=>{
    const inset=(o.taper||0)*i, x=o.x+inset, w=o.w-2*inset;
    s.addShape('rect',{ x, y:o.y+i*h, w, h:h-0.06,
      fill:{ color: i===0?s._C.dark : i===1?s._C.mid : K.g5 }, line:{ color: i===2?K.g2:s._C.dark, width:0.75 } });
    txt(s,it.t,{ x:x+0.16, y:o.y+i*h, w:w*0.32, h:h-0.06, sz:T.body, b:true, c:i===2?K.ink:K.w, valign:'middle' });
    txt(s,it.d,{ x:x+w*0.34, y:o.y+i*h, w:w*0.64, h:h-0.06, sz:T.small, c:i===2?K.g1:K.g5, valign:'middle', lh:1.2 });
  });
}
/**
 * 패널 — 제목 박스. 제목은 **좌우 가운데 정렬**하고 얇은 밑줄을 둔 뒤,
 * 본문 항목을 박스 높이에 맞춰 **고르게 분산**해 아래가 비지 않게 한다.
 *   items: ['항목', {t:'항목', sub:'딸린 설명'}]
 */
function panel(s, o){
  const T=s._T, pad=o.pad==null?T.pad:o.pad;
  box(s,{ x:o.x, y:o.y, w:o.w, h:o.h, fill:o.fill||K.w, line:o.line||K.g4, what:'패널' });
  const innerW = o.w - 2*pad;
  let top = o.y + pad;
  if (o.title){
    txt(s, o.title, { x:o.x+pad, y:top, w:innerW, h:0.30, sz:o.titleSize||T.h2, b:true,
      align:'center', valign:'middle' });                     // 제목은 가운데 정렬
    hr(s, { x:o.x+pad, y:top+0.34, w:innerW, color:K.g4 });
    top += 0.46;
  }
  const items = (o.items||[]).map(it => typeof it==='string' ? {t:it} : it);
  if (!items.length) return;
  const sz = o.sz || T.body, lh = o.lh || T.lh, subSz = o.subSz || Math.max(8.5, sz-0.5);
  const lineH = sz*lh/72, subH = subSz*lh/72;
  // 각 항목이 차지할 실제 높이
  const heights = items.map(it =>
    lines(it.t, sz, innerW-0.22)*lineH + (it.sub ? lines(it.sub, subSz, innerW-0.40)*subH : 0));
  const used = heights.reduce((a,b)=>a+b,0);
  const room = (o.y + o.h - pad) - top;
  if (used > room + 0.005)
    throw new Error(`패널 '${o.title||''}' 안의 글(${used.toFixed(2)}")이 남은 높이 ${room.toFixed(2)}"를 넘는다`
      + ` — 높이를 ${(o.h + used - room).toFixed(2)}"로 키우거나 항목을 줄일 것`);
  // 남는 자리를 항목 사이에 고르게 나눠 박스를 채운다 (과하게 벌어지지 않게 상한)
  const gap = items.length>1 ? Math.max(0.02, Math.min((room-used)/(items.length-1), lineH*1.6)) : 0;
  let y = top + Math.max(0, (room - used - gap*(items.length-1))/2);
  items.forEach(it => {
    const hT = lines(it.t, sz, innerW-0.22)*lineH;
    txt(s, '–', { x:o.x+pad, y, w:0.16, h:hT, sz, c:K.g2 });
    txt(s, it.t, { x:o.x+pad+0.22, y, w:innerW-0.22, h:hT, sz, lh });
    y += hT;
    if (it.sub){
      const hS = lines(it.sub, subSz, innerW-0.40)*subH;
      txt(s, it.sub, { x:o.x+pad+0.40, y, w:innerW-0.40, h:hS, sz:subSz, c:K.g2, lh });
      y += hS;
    }
    y += gap;
  });
}

function callout(s, t, o){
  const T=s._T;
  const cc = o.acc ? s._C.acc : s._C.dark;
  box(s,{ x:o.x, y:o.y, w:o.w, h:o.h, fill:o.dark?cc:K.g6, line:o.dark?cc:K.g4, what:'콜아웃' });
  txt(s,t,{ x:o.x+T.pad, y:o.y, w:o.w-2*T.pad, h:o.h, sz:o.sz||T.body, b:true,
    c:o.dark?K.w:K.ink, valign:'middle', lh:1.3 });
}
/** 출처 — 도식·표 바로 아래에 붙인다. 우측 하단 푸터에 두지 않는다. */
function source(s, t, o){
  txt(s, t, { x:o.x, y:o.y, w:o.w, h:o.h||0.22, sz:Math.max(9, s._T.cap), c:K.g2, lh:1.2,
    align:o.align||'left' });
}

function footnote(s, t, o){
  txt(s,t,{ x:o.x, y:o.y, w:o.w, h:o.h||0.22, sz:Math.max(9,s._T.cap), c:K.g2, lh:1.2,
    align:o.align||'left' });
}
function pill(s, t, o){
  const T=s._T, h=o.h||0.26;
  s.addShape('roundRect',{ x:o.x, y:o.y, w:o.w, h, rectRadius:h/2,
    fill:{color:o.dark?K.ink:K.g6}, line:{color:o.dark?K.ink:K.g4, width:0.75} });
  txt(s,t,{ x:o.x, y:o.y, w:o.w, h, sz:o.sz||T.small, b:true, c:o.dark?K.w:K.ink, align:'center', valign:'middle', wrap:false });
}

/* ── 표지·목차 ─────────────────────────────────────────── */
/**
 * 표지 — 제목 전용. 우측에 격자 모티프(좌상 → 우하로 차오르는 체계의 은유)를 깔고
 * 좌측에 제목·부제를, 하단에 기관·일자를 둔다. 목차는 넣지 않는다(다음 장이 목차다).
 * 사진을 구할 수 없을 때만 쓴다. 기본 표지는 cover()의 사진 표지다.
 */
function coverPlain(deck, o){
  const s = deck.bare({}); const C = s._C;
  const gx=5.75, gy=1.30, cell=0.40, gap=0.07, ROW=11, COL=10;
  for (let r=0;r<ROW;r++) for (let c=0;c<COL;c++){
    const t=(r/(ROW-1))*0.55 + (c/(COL-1))*0.45;
    const fill = t>0.86?K.ink : t>0.68?K.g2 : t>0.50?K.g4 : t>0.32?K.g5 : K.g7;
    s.addShape('rect',{ x:gx+c*(cell+gap), y:gy+r*(cell+gap), w:cell, h:cell, fill:{color:fill}, line:{type:'none'} });
  }
  if (C.acc !== K.ink)   // 강조 팔레트일 때만 포인트 셀 하나
    s.addShape('rect',{ x:gx+7*(cell+gap), y:gy+2*(cell+gap), w:cell, h:cell, fill:{color:C.acc}, line:{type:'none'} });
  s.addShape('rect',{ x:0, y:2.10, w:5.55, h:2.55, fill:{color:K.w}, line:{type:'none'} });   // 제목 자리 확보
  s.addShape('rect',{ x:0, y:6.35, w:W,    h:1.15, fill:{color:K.w}, line:{type:'none'} });   // 하단 정보 자리
  txt(s, o.org||'한국표준협회', { x:M, y:0.40, w:5, h:0.30, sz:12, b:true, c:K.g2, valign:'middle' });
  s.addShape('rect',{ x:M, y:2.45, w:0.10, h:1.55, fill:{color:C.acc}, line:{type:'none'} });
  txt(s, o.title, { x:M+0.32, y:2.42, w:4.6, h:1.40, sz:38, b:true, lh:1.15 });
  if (o.subtitle) txt(s, o.subtitle, { x:M+0.32, y:3.85, w:4.6, h:0.70, sz:12.5, c:K.g1, lh:1.4 });
  txt(s, `${o.org||'한국표준협회'}${o.team? ' · '+o.team : ''}`, { x:M, y:6.60, w:5, h:0.28, sz:10.5, c:K.g2 });
  if (o.date) txt(s, o.date, { x:W-M-3, y:6.60, w:3, h:0.28, sz:10.5, c:K.g2, align:'right' });
  return s;
}

/**
 * 목차 — 본문 장표 프레임을 그대로 쓴다. '목차'는 헤드라인에 한 번만 쓰고
 * 좌상단 챕터명은 비워 중복을 없앤다. 항목은 {t, p} 또는 {n, t, p}(장) 형식.
 */

/* ── 사진 표지 ───────────────────────────────────────────
 * 상단 4분할 사진 띠(사선 분할) + 좌하단 먹 패널 + 우하단 대형 사진.
 * 사진은 각각 **독립된 그림 도형**이라 PowerPoint에서 [그림 바꾸기]로 갈아끼운다.
 * 사진을 주지 않으면 assets/photo_placeholder.png 자리표시가 들어간다.
 *   photos: [띠1, 띠2, 띠3, 띠4, 대형] — 없거나 모자라면 자리표시로 채운다
 * 사선은 후처리(postprocess.py)가 띠 사진의 도형을 평행사변형으로 바꿔 만든다.
 */
const COVER = {
  bandY: 0.12, bandH: 2.61, lowY: 2.89,
  x0: M, x1: +(W-M).toFixed(4),
  lean: 0.2126,                    // 사선 기울기 tan(12°)
  darkR: 4.66, wedge: 1.04,        // 먹 패널 오른쪽 끝 · 사선 폭
  heroX: 3.60,                     // 대형 사진 왼쪽 끝 (먹 패널에 가려지는 구간)
};

function placeholderPath(){
  const p = path.join(__dirname, 'photo_placeholder.png');
  return fs.existsSync(p) ? p : null;
}

function cover(deck, o){
  const s = deck.bare({}); const C = s._C;
  const K_ = K, c = COVER;
  const ph = placeholderPath();
  const photos = Array.from({length:5}, (_,i) => (o.photos && o.photos[i]) || ph);
  if (photos.some(p => !p))
    throw new Error('표지 사진이 없고 assets/photo_placeholder.png도 없다 — 자리표시 파일을 두거나 photos를 넘길 것');

  const dx = +(c.bandH * c.lean).toFixed(4);
  const span = +(c.x1 - c.x0).toFixed(4);
  const wt = +((span - dx)/4).toFixed(4);
  const y0 = c.bandY, y1 = +(y0 + c.bandH).toFixed(4);

  // 띠 — 0번은 왼쪽 모서리를 살리려 직사각형, 1~3번은 평행사변형(후처리에서 도형 변경)
  s.addImage({ path:photos[0], x:c.x0, y:y0, w:+(wt+dx).toFixed(4), h:c.bandH,
    sizing:{ type:'cover', w:+(wt+dx).toFixed(4), h:c.bandH }, altText:'KSA_COVER_RECT 표지 사진 1' });
  for (let i=1;i<4;i++){
    const x = +(c.x0 + i*wt).toFixed(4);
    s.addImage({ path:photos[i], x, y:y0, w:+(wt+dx).toFixed(4), h:c.bandH,
      sizing:{ type:'cover', w:+(wt+dx).toFixed(4), h:c.bandH },
      altText:`KSA_COVER_BAND 표지 사진 ${i+1}` });
  }
  // 사진 사이 흰 사선
  for (let i=1;i<4;i++)
    s.addShape('line',{ x:+(c.x0 + i*wt).toFixed(4), y:y0, w:dx, h:c.bandH, flipV:true,
      line:{ color:K_.w, width:7 } });

  // 하단 — 대형 사진 위에 먹 패널과 삼각형을 얹어 사선을 만든다
  // 먹 패널 색 — 무채색이면 #3A3A3A, 강조 팔레트면 네이비
  const panel = o.panel || (C.dark === K_.ink ? K_.g1 : C.dark);
  const lh = +(H - c.lowY).toFixed(4);
  // 대형 사진은 먹 패널 아래로 조금만 물린다 — 사선 자리만 덮으면 되고, 그래야 사진의 가운데가 보인다
  const hx = c.heroX, hw = +(c.x1 - hx).toFixed(4);
  s.addImage({ path:photos[4], x:hx, y:c.lowY, w:hw, h:lh,
    sizing:{ type:'cover', w:hw, h:lh }, altText:'KSA_COVER_HERO 표지 대형 사진' });
  s.addShape('rect',{ x:c.x0, y:c.lowY, w:+(c.darkR-c.x0).toFixed(4), h:lh,
    fill:{color:panel}, line:{type:'none'} });
  s.addShape('rtTriangle',{ x:c.darkR, y:c.lowY, w:c.wedge, h:lh, flipV:true,
    fill:{color:panel}, line:{type:'none'} });
  s.addShape('line',{ x:c.darkR, y:c.lowY, w:c.wedge, h:lh, flipV:true,
    line:{ color:K_.w, width:7 } });

  // 글자 — 먹 패널 위
  const tx = 0.60;
  txt(s, o.org || '', { x:tx, y:3.78, w:3.95, h:0.30, sz:14, b:true, c:K_.w, valign:'middle' });
  txt(s, o.title,     { x:tx, y:4.10, w:4.30, h:1.00, sz:26, b:true, c:K_.w, valign:'top', lh:1.15 });
  hr(s, { x:tx, y:5.42, w:3.90, color:K_.w, width:1 });
  txt(s, o.date || '', { x:tx+0.02, y:5.56, w:3.5, h:0.28, sz:12, c:K_.w, valign:'middle' });
  txt(s, o.by || '한국표준협회', { x:tx+0.02, y:5.94, w:3.5, h:0.28, sz:12, c:K_.w, valign:'middle' });
  return s;
}

function toc(deck, o){
  const s = deck.slide({ chapter: '' }); const C = s._C;
  head(s, o.title || '목차');            // '목차'는 타이틀에 한 번만 — 메시지는 두지 않는다
  const CH_H = 0.39, SUB_H = 0.27, GAP = 0.17;   // 장 구분선이 앞 항목 글자와 떨어지도록
  const BTt = s._T.bt;
  let y = BTt + 0.06;
  hr(s, { x:M, y:BTt, w:CW, color:K.ink, width:1 });   // 본문 앵커 겸 첫 장 구분선
  o.items.forEach((it, i) => {
    const top = !!it.n, sz = top ? 14 : 11.5;
    if (top && i > 0) { y += GAP; hr(s, { x:M, y:y-0.05, w:CW, color:K.ink, width:1 }); }
    if (top) txt(s, it.n, { x:M, y:y+0.05, w:0.6, h:0.30, sz:12.5, b:true, c:C.acc });
    txt(s, it.t, { x:M+(top?0.90:1.20), y:y+0.04, w:6.0, h:0.30, sz, b:top, c:top?K.ink:K.g1 });
    const lx = M + (top?0.90:1.20) + textW(it.t, sz) + 0.16;
    s.addShape('line',{ x:lx, y:y+0.23, w:Math.max(0.2, W-M-0.85-lx), h:0,
      line:{ color:K.g4, width:0.75, dashType:'sysDot' } });
    txt(s, String(it.p), { x:W-M-0.7, y:y+0.04, w:0.7, h:0.30, sz:top?13:11.5, b:top,
      c:top?K.ink:K.g1, align:'right' });
    y += top ? CH_H : SUB_H;
  });
  guard(s, BTt, y - BTt, '목차 목록');
  return s;
}

module.exports = { P, K, FONT, TEXT_MIN, W, H, M, CW, BB, TITLE_Y, TITLE_MAX, Z, zones, col, cx, cw, split, MODE,
  textW, lines, needH, PALETTE, createDeck, head, guard, box, txt,
  underline, hr, sectionTitle, kpiRow, table, tableNative, bullets, chevrons, waterfall, tree, matrix,
  gantt, layers, callout, panel, footnote, source, pill, cover, coverPlain, toc,
  barChart, lineChart, pieChart, comboChart, chartNote, seriesRamp, SERIES_MAX,
  venn, hubSpoke, cycle, pyramid, steps, harvey, harveyBall, causeEffect, HARVEY_STEPS, onTone };
