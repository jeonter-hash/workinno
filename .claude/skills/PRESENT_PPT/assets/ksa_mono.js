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
const FONT = '맑은 고딕';
const TEXT_MIN = K.g2;          // 흰 배경 위 글자의 최소 명도 — 이보다 흐리면 쓰지 않는다

/* ── 판형·존 (A4 가로) ──────────────────────────────────── */
const W = 10.8333, H = 7.5, M = 0.5, CW = W - 2*M;
const BT = 2.39, BB = 6.85;                    // 본문 상·하한 (벽)
const Z = {
  chapter:{ x:M, y:0.40, w:6.2, h:0.30 },
  rule:   { y:0.82 },
  logo:   { y:0.44, h:0.24 },
  head:   { x:M, y:1.00, w:CW, h:0.75 },
  sub:    { x:M, y:1.63, w:CW, h:0.40 },
  body:   { x:M, y:BT,   w:CW, h:BB-BT },
  foot:   { x:M, y:7.05, w:CW, h:0.25 },
};
const col = (CW-11*0.2)/12;
const cx = i => +(M + i*(col+0.2)).toFixed(4);
const cw = n => +(n*col + (n-1)*0.2).toFixed(4);
const split = n => { if (12%n) throw new Error('split은 2·3·4·6'); const sp=12/n;
  return Array.from({length:n},(_,i)=>({ x:cx(i*sp), w:cw(sp) })); };

/* ── 밀도 프로파일 ──────────────────────────────────────── */
const MODE = {
  present: { name:'발표용', chapter:14, h1:24, sub:13, h2:14, body:11.5, small:10, cap:9.5,
             kpi:34, kpiLabel:10.5, tableHead:10.5, tableBody:10.5, rowH:0.44, pad:0.17, lh:1.35, gap:0.20 },
  report:  { name:'보고서용', chapter:12, h1:20, sub:11.5, h2:12, body:9.5, small:9, cap:9,
             kpi:24, kpiLabel:9, tableHead:9.5, tableBody:9.5, rowH:0.32, pad:0.12, lh:1.28, gap:0.14 },
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

/* ── 이미지 크기(로고 비율 자동) ────────────────────────── */
function imgSize(f){
  try{ const b=fs.readFileSync(f);
    if (b.slice(1,4).toString()==='PNG') return { w:b.readUInt32BE(16), h:b.readUInt32BE(20) };
    if (b[0]===0xff&&b[1]===0xd8){ let i=2;
      while(i<b.length-9){ if(b[i]!==0xff){i++;continue;} const m=b[i+1];
        if(m>=0xc0&&m<=0xcf&&![0xc4,0xc8,0xcc].includes(m)) return { w:b.readUInt16BE(i+7), h:b.readUInt16BE(i+5) };
        if(m===0xd8||m===0xd9||(m>=0xd0&&m<=0xd7)){i+=2;continue;} i+=2+b.readUInt16BE(i+2); } }
  }catch(e){}
  return null;
}
function findLogo(explicit){
  if (explicit && fs.existsSync(explicit)) return explicit;
  const cands=['ksa_logo.png','ksa_logo.jpg'];
  const dirs=[__dirname, path.join(__dirname,'assets'), process.cwd()];
  for(const d of dirs) for(const n of cands){ const p=path.join(d,n); if(fs.existsSync(p)) return p; }
  return null;
}

/* ── 덱 ─────────────────────────────────────────────────── */
function createDeck(o={}){
  const T = MODE[o.mode||'present'];
  if (!T) throw new Error("mode는 'present' 또는 'report'");
  const pres = new P();
  pres.defineLayout({ name:'A4L', width:W, height:H }); pres.layout='A4L';
  if (o.title) pres.title=o.title;
  if (o.author) pres.author=o.author;
  const logo = findLogo(o.logo);
  const deck = { pres, T, logo, page:0,
    slide(opt={}){
      const s = pres.addSlide(); s.background={color:K.w}; s._T=T; s._deck=deck;
      deck.page += 1; s._page = opt.page==null? deck.page : opt.page;
      frame(s, opt); return s;
    },
    bare(opt={}){ const s=pres.addSlide(); s.background={color:opt.bg||K.w}; s._T=T; s._deck=deck;
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
  if (d.logo){ const sz=imgSize(d.logo), h=Z.logo.h, w=sz? +(h*sz.w/sz.h).toFixed(4):1.22;
    s.addImage({ path:d.logo, x:+(W-M-w).toFixed(4), y:Z.logo.y, w, h, altText:'한국표준협회 로고' }); }
  s.addShape('line', { x:M, y:Z.rule.y, w:CW, h:0, line:{ color:K.g4, width:0.75 } });
  s.addText(String(o.page==null? s._page : o.page), { ...Z.foot, w:2, fontFace:FONT, fontSize:T.cap,
    color:K.g2, valign:'middle', margin:0 });
  if (o.source) s.addText(o.source, { x:M+3.6, y:Z.foot.y, w:CW-3.6, h:Z.foot.h, fontFace:FONT,
    fontSize:T.cap, color:K.g2, align:'right', valign:'middle', margin:0 });
  return s;
}
/** 로고를 우상단(오른쪽 여백 0.5")에 원본 비율로 배치 */
function logoAt(s, o={}){
  const d=s._deck; if(!d||!d.logo) return null;
  const sz=imgSize(d.logo), h=o.h||Z.logo.h, w=sz? +(h*sz.w/sz.h).toFixed(4):1.22;
  const x=o.x==null? +(W-M-w).toFixed(4) : o.x;
  s.addImage({ path:d.logo, x, y:o.y==null?Z.logo.y:o.y, w, h, altText:'한국표준협회 로고' });
  return { x, w, h };
}

const head = (s,t,o={}) => s.addText(t, { ...Z.head, fontFace:FONT, fontSize:o.size||s._T.h1, bold:true,
  color:K.ink, valign:'top', margin:0, charSpacing:-0.4, lineSpacingMultiple:1.25 });
const sub  = (s,t) => s.addText(t, { ...Z.sub, fontFace:FONT, fontSize:s._T.sub, color:K.g2,
  valign:'top', margin:0, lineSpacingMultiple:1.4 });

/* ── 기본 조각 ──────────────────────────────────────────── */
function guard(y,h,what='요소'){
  if (y < BT-0.005) throw new Error(`${what}가 본문 상단(2.39")을 침범: y=${y}`);
  if (y+h > BB+0.005) throw new Error(`${what}가 본문 하단(6.85")을 침범: ${(y+h).toFixed(3)}`);
}
function box(s,o){
  guard(o.y,o.h,o.what||'박스');
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
    box(s,{ x:c.x, y:o.y, w:c.w, h, fill:dark?K.ink:K.w, line:dark?K.ink:K.g4, what:'KPI' });
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
  guard(o.y,h,'표');
  if (o.x + total > W - M + 0.005)
    throw new Error(`표 폭 합계 ${total}"가 우측 여백을 넘음 — x=${o.x}에서 쓸 수 있는 폭은 ${+(W-M-o.x).toFixed(4)}"`);
  o.rows.forEach((r,i)=>{ if (r.length!==colW.length)
    throw new Error(`표 ${i+1}행의 칸 수(${r.length})가 열 수(${colW.length})와 다름`); });
  s.addShape('rect',{ x:o.x, y:o.y, w:total, h:headH, fill:{color:K.ink}, line:{type:'none'} });
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
  guard(o.y, tall, '표');
  if (o.x + total > W - M + 0.005)
    throw new Error(`표 폭 합계 ${total}"가 우측 여백을 넘음 — x=${o.x}에서 쓸 수 있는 폭은 ${+(W-M-o.x).toFixed(4)}"`);
  o.rows.forEach((r,i)=>{ if (r.length!==colW.length)
    throw new Error(`표 ${i+1}행의 칸 수(${r.length})가 열 수(${colW.length})와 다름`); });
  const hair = c => ({ type:'solid', color:c, pt:0.75 });
  const none = { type:'none' };
  const headRow = o.head.map((t)=>({ text:String(t), options:{
    fill:{ color:K.ink }, color:K.w, bold:true, fontSize:T.tableHead,
    align:'center', valign:'middle',              // 표 머리글은 항상 가운데 정렬
    border:[none,none,hair(K.ink),none] } }));
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
  const T=s._T;
  const runs=items.map((t,i)=>({ text:typeof t==='string'?t:t.text,
    options:{ bullet:{ code:'2013' }, breakLine:i<items.length-1, ...(typeof t==='object'&&t.b?{bold:true}:{}) } }));
  s.addText(runs, { x:o.x, y:o.y, w:o.w, h:o.h, fontFace:FONT, fontSize:o.sz||T.body, color:K.ink,
    valign:'top', margin:0, lineSpacingMultiple:o.lh||T.lh, paraSpaceAfter:o.space==null?4:o.space });
}
function chevrons(s, steps, o){
  const T=s._T, n=steps.length, step=(o.w)/n, chW=step+0.30, h=o.h||0.80;
  steps.forEach((st,i)=>{
    const x=o.x+i*(step-0.05);
    s.addShape(i===0?'homePlate':'chevron',{ x, y:o.y, w:chW, h,
      fill:{ color: st.tone==='light'?K.g4 : st.tone==='mid'?K.g2 : K.ink }, line:{ color:K.w, width:1 } });
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
    if (st.base){ top=o.y+plotH-st.v*sc; hh=st.v*sc; run=st.v; fill=K.ink; ln=K.ink; }
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
  s.addShape('rect',{ x:r.x, y:r.y, w:r.w, h:r.h, fill:{color:K.ink}, line:{type:'none'} });
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
    s.addShape('ellipse',{ x:px-d/2, y:py-d/2, w:d, h:d, fill:{color:hot?K.ink:K.g4}, line:{color:hot?K.ink:K.g2,width:0.75} });
    txt(s,p.t,{ x:px+d/2+0.06, y:py-0.12, w:1.9, h:0.24, sz:T.small, b:hot, c:hot?K.ink:K.g2, valign:'middle' });
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
      fill:{ color: r.tone==='light'?K.g4 : r.tone==='mid'?K.g2 : K.ink }, line:{type:'none'} });
    const mx=gx+r.b*gw;
    s.addShape('diamond',{ x:mx-0.085, y:y+(rowH-0.24)/2, w:0.17, h:0.17, fill:{color:K.w}, line:{color:K.ink,width:1} });
    if (r.ms) txt(s,r.ms,{ x:gx+gw+0.22, y:y+0.06, w:noteW-0.24, h:0.30, sz:T.small, c:K.g2, valign:'middle' });
  });
  return o.y+0.36+rowH*o.rows.length;
}
function layers(s, o){          // 3단 계층(교육체계 등)
  const T=s._T, n=o.items.length, h=o.h/n;
  o.items.forEach((it,i)=>{
    const inset=(o.taper||0)*i, x=o.x+inset, w=o.w-2*inset;
    s.addShape('rect',{ x, y:o.y+i*h, w, h:h-0.06,
      fill:{ color: i===0?K.ink : i===1?K.g2 : K.g5 }, line:{ color: i===2?K.g2:K.ink, width:0.75 } });
    txt(s,it.t,{ x:x+0.16, y:o.y+i*h, w:w*0.32, h:h-0.06, sz:T.body, b:true, c:i===2?K.ink:K.w, valign:'middle' });
    txt(s,it.d,{ x:x+w*0.34, y:o.y+i*h, w:w*0.64, h:h-0.06, sz:T.small, c:i===2?K.g1:K.g5, valign:'middle', lh:1.2 });
  });
}
function callout(s, t, o){
  const T=s._T;
  box(s,{ x:o.x, y:o.y, w:o.w, h:o.h, fill:o.dark?K.ink:K.g6, line:o.dark?K.ink:K.g4, what:'콜아웃' });
  txt(s,t,{ x:o.x+T.pad, y:o.y, w:o.w-2*T.pad, h:o.h, sz:o.sz||T.body, b:true,
    c:o.dark?K.w:K.ink, valign:'middle', lh:1.3 });
}
function footnote(s, t, o){
  txt(s,t,{ x:o.x, y:o.y, w:o.w, h:o.h||0.22, sz:Math.max(9,s._T.cap), c:K.g2, lh:1.2 });
}
function pill(s, t, o){
  const T=s._T, h=o.h||0.26;
  s.addShape('roundRect',{ x:o.x, y:o.y, w:o.w, h, rectRadius:h/2,
    fill:{color:o.dark?K.ink:K.g6}, line:{color:o.dark?K.ink:K.g4, width:0.75} });
  txt(s,t,{ x:o.x, y:o.y, w:o.w, h, sz:o.sz||T.small, b:true, c:o.dark?K.w:K.ink, align:'center', valign:'middle', wrap:false });
}

module.exports = { P, K, FONT, TEXT_MIN, W, H, M, CW, BT, BB, Z, col, cx, cw, split, MODE,
  textW, lines, needH, imgSize, findLogo, createDeck, frame, head, sub, guard, box, txt,
  underline, hr, logoAt, sectionTitle, kpiRow, table, tableNative, bullets, chevrons, waterfall, tree, matrix,
  gantt, layers, callout, footnote, pill };
