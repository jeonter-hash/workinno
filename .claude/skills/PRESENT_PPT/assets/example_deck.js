/** 발표용 샘플 12장 — 큰 글자·적은 항목·도식 중심 */
'use strict';
const B = require('./ksa_mono.js');
const { K, M, CW, BT, BB, W } = B;
const OUT = process.argv[2] || 'present_sample.pptx';
const deck = B.createDeck({ mode:'present', title:'AX 전환 로드맵 (발표용)', author:'한국표준협회' });
const T = deck.T;
const S = (ch, src) => deck.slide({ chapter:ch, source:src });

/* 1 표지 */
{
  const s = deck.bare({});
  s.addShape('rect',{ x:0, y:0, w:0.26, h:B.H, fill:{color:K.ink}, line:{type:'none'} });
  s.addText('한국표준협회', { x:M+0.3, y:0.40, w:6, h:0.30, fontFace:B.FONT, fontSize:14, bold:true, color:K.g2, valign:'middle', margin:0 });
  B.logoAt(s,{});
  s.addShape('line',{ x:M+0.3, y:0.82, w:CW-0.3, h:0, line:{color:K.g4,width:0.75} });
  const title='AX 전환 로드맵';
  s.addText(title,{ x:M+0.3, y:1.72, w:CW-0.6, h:0.95, fontFace:B.FONT, fontSize:44, bold:true, color:K.ink, valign:'bottom', margin:0, charSpacing:-1.2 });
  s.addShape('line',{ x:M+0.3, y:2.80, w:B.textW(title,44), h:0, line:{color:K.ink,width:2.5} });
  s.addText('업무 재설계부터 인력 전환까지 12개월 실행 설계',{ x:M+0.3, y:2.99, w:7.5, h:0.4, fontFace:B.FONT, fontSize:15, color:K.g1, margin:0 });
  [['01','현황 진단'],['02','전환 과제'],['03','실행 로드맵'],['04','성과 관리']].forEach(([n,t],i)=>{
    const x=M+0.3+i*2.45;
    s.addShape('line',{ x, y:4.05, w:2.15, h:0, line:{color:K.g4,width:0.75} });
    s.addText(n,{ x, y:4.15, w:2.15, h:0.34, fontFace:B.FONT, fontSize:20, bold:true, color:K.g2, margin:0 });
    s.addText(t,{ x, y:4.52, w:2.15, h:0.30, fontFace:B.FONT, fontSize:13, bold:true, color:K.ink, margin:0 });
  });
  s.addShape('rect',{ x:M+0.3, y:5.35, w:CW-0.6, h:1.15, fill:{color:K.g6}, line:{color:K.g4,width:0.75} });
  s.addText([{text:'핵심 결론  ',options:{fontFace:B.FONT,fontSize:12,bold:true,color:K.g2}},
             {text:'전환의 병목은 인력이 아닌 업무 정의이며, 표준 업무 단위 확정이 교육·배치 설계의 선행 조건임',options:{fontFace:B.FONT,fontSize:13,bold:true,color:K.ink}}],
    { x:M+0.55, y:5.35, w:CW-1.1, h:1.15, valign:'middle', margin:0, lineSpacingMultiple:1.35 });
  s.addText('1',{ x:M, y:7.05, w:2, h:0.25, fontFace:B.FONT, fontSize:9.5, color:K.g2, valign:'middle', margin:0 });
}

/* 2 목차 */
{
  const s=S('목차','');
  B.head(s,'보고 순서');
  B.sub(s,'4개 장, 12개 장표 — 각 장 말미에 결론 1건 제시');
  const items=[['01','현황 진단','업무시간 구조·절감 여력 산정'],['02','전환 과제','3대 경로·5단계 게이트 설계'],
               ['03','실행 로드맵','우선순위·추진 체계·12개월 일정'],['04','성과 관리','교육 체계·성과지표 운영']];
  items.forEach(([n,t,d],i)=>{
    const y=BT+i*1.14;
    B.box(s,{ x:M, y, w:CW, h:1.02, fill:i===0?K.g6:K.w, line:K.g4 });
    B.txt(s,n,{ x:M+0.24, y:y+0.16, w:0.9, h:0.5, sz:26, b:true, c:K.g2 });
    B.txt(s,t,{ x:M+1.20, y:y+0.16, w:3.0, h:0.36, sz:16, b:true });
    B.txt(s,d,{ x:M+1.20, y:y+0.56, w:5.6, h:0.30, sz:T.small, c:K.g2 });
    B.pill(s,`${(i+1)*3-2}–${(i+1)*3}쪽`,{ x:M+CW-1.35, y:y+0.36, w:1.1, h:0.28 });
  });
}

/* 3 핵심 요약 */
{
  const s=S('요약','출처: 업무분석 결과(예시), 2026-01');
  B.head(s,'절감 여력 8,100시간, 상위 4개 업무에 65% 집중');
  B.sub(s,'전면 착수 대신 상위 과제 집중 추진 권고 — 1년 내 6,300시간 절감 가능');
  B.kpiRow(s,[{v:'12,400h',l:'연간 총 업무시간'},{v:'8,100h',l:'정형·반복 구간'},
              {v:'4개',l:'절감 상위 업무'},{v:'65%',l:'상위 4개 집중도',dark:true}],{ y:BT, h:1.20 });
  const cells=B.split(3);
  ['진단 — 12개 업무 실측, 절감 경로 3종 확인','설계 — 5단계 게이트, 통과 기준 사전 합의',
   '실행 — 4개 과제 우선, 분기별 산출물 1건'].forEach((t,i)=>{
    const y=BT+1.42;
    B.box(s,{ x:cells[i].x, y, w:cells[i].w, h:1.34, fill:K.w, line:K.g4 });
    B.txt(s,`0${i+1}`,{ x:cells[i].x+T.pad, y:y+0.12, w:0.7, h:0.28, sz:12, b:true, c:K.g2 });
    B.txt(s,t,{ x:cells[i].x+T.pad, y:y+0.46, w:cells[i].w-2*T.pad, h:0.74, sz:T.body, b:true, lh:1.35 });
  });
  B.callout(s,'권고 — 상위 4개 과제에 자원을 집중하고, 잔여 과제는 2차 연도로 이월',{ x:M, y:BT+3.02, w:CW, h:0.62, dark:true });
  B.footnote(s,'※ 절감 시간은 실측 기준이며 인건비 환산은 별도 산정 필요',{ x:M, y:BT+3.78, w:CW });
}

/* 4 워터폴 */
{
  const s=S('01 현황 진단','출처: 업무분석 로그(예시)');
  B.head(s,'절감 경로별 기여도 확인');
  B.sub(s,'문서 자동화 21천 시간으로 최대 — 4개 경로 합산 시 목표 수준 도달');
  B.box(s,{ x:M, y:BT, w:CW, h:4.46, fill:K.w, line:K.g4 });
  B.sectionTitle(s,'업무시간 절감 경로 (천 시간)',{ x:M+0.24, y:BT+0.18, w:5.2 });
  const axB=B.waterfall(s,{ x:M+0.4, y:BT+0.80, w:CW-0.8, h:2.30, max:130,
    steps:[{t:'현재',v:124,base:true},{t:'문서 자동화',v:-21},{t:'데이터 연계',v:-14},
           {t:'판정 표준화',v:-11},{t:'교육 이관',v:-6},{t:'목표',v:72,base:true}] });
  B.hr(s,{ x:M+0.24, y:axB+0.14, w:CW-0.48, color:K.g4 });
  B.txt(s,'상위 4개 업무만으로 목표의 82% 달성 가능 — 잔여 과제는 2차 추진으로 이월 권고',
    { x:M+0.24, y:axB+0.22, w:CW-0.48, h:0.32, sz:T.body, b:true, c:K.g1 });
}

/* 5 드라이버 트리 */
{
  const s=S('01 현황 진단','출처: 업무분석 로그(예시)');
  B.head(s,'절감 시간의 3대 경로: 자동화·표준화·이관');
  B.sub(s,'3개 경로 외 활동은 절감 산정 대상에서 제외');
  B.box(s,{ x:M, y:BT, w:6.4, h:4.46, fill:K.w, line:K.g5 });
  const mids=[['자동화\n3,600h',['문서 생성  1,900h','데이터 입력  1,700h']],
              ['표준화\n2,700h',['판정 기준  1,500h','서식 통일  1,200h']],
              ['이관\n1,800h',['1차 검토  1,100h','단순 응대  700h']]];
  B.tree(s,{ root:{ x:M+0.22, y:BT+1.85, w:1.35, h:0.72, t:'절감\n8,100h' },
    mids: mids.map(([t,leaves],i)=>{
      const y=BT+0.55+i*1.42;
      return { x:M+2.05, y, w:1.45, h:0.66, t,
        leaves: leaves.map((lt,k)=>({ x:M+3.92, y:y-0.30+k*0.72, w:1.85, h:0.56, t:lt })) };
    }) });
  const rx=M+6.7, rw=W-M-(M+6.7);
  ['자동화 — 검증된 도구 적용으로 3개월 내 완료 가능','표준화 — 부서 합의가 병목, 게이트 설정 필요','이관 — 교육 병행 없이는 원복 우려']
   .forEach((t,i)=>{
    const y=BT+i*1.55;
    B.box(s,{ x:rx, y, w:rw, h:1.36, fill:i===1?K.ink:K.g6, line:i===1?K.ink:K.g4 });
    B.txt(s,`0${i+1}`,{ x:rx+0.2, y:y+0.14, w:0.6, h:0.28, sz:11, b:true, c:i===1?K.g4:K.g2 });
    B.txt(s,t,{ x:rx+0.2, y:y+0.46, w:rw-0.4, h:0.78, sz:T.body, b:true, c:i===1?K.w:K.ink, lh:1.35 });
  });
}

/* 6 체브론 + 게이트표 */
{
  const s=S('02 전환 과제','평가: 컨설턴트 5인 합의(예시)');
  B.head(s,'전환 5단계 구분 및 단계별 통과 기준 설정');
  B.sub(s,'단계별 산출물을 다음 단계 입력으로 연결 — 게이트 미통과 시 다음 단계 진입 유보');
  B.chevrons(s,[{t:'1. 진단',d:'업무 목록·시간 실측'},{t:'2. 설계',d:'표준 업무 단위 확정'},
                {t:'3. 시범',d:'2개 부서 파일럿'},{t:'4. 확산',d:'전사 이관·교육',tone:'mid'},
                {t:'5. 정착',d:'성과 측정·보정',tone:'light'}],{ x:M, y:BT, w:CW, h:0.84 });
  B.tableNative(s,{ x:M, y:BT+1.72, colW:[4.6,1.6,1.6,2.0333], headH:0.40, rowH:0.50,
    head:['통과 기준','난이도','자원 소요','선행 조건'], align:['left','center','center','center'], boldCol:[3],
    rows:[['업무 목록 100% 실측 완료','●','◐','—'],
          ['표준 업무 단위 정의서 승인','●','●','진단 완료'],
          ['파일럿 2개 부서 KPI 달성','◐','●','설계 승인'],
          ['전 부서 이관율 90% 이상','●','●','파일럿 성공']] });
  B.footnote(s,'● 높음   ◐ 보통   ○ 낮음',{ x:M, y:BT+4.16, w:4 });
}

/* 7 2×2 매트릭스 */
{
  const s=S('03 실행 로드맵','평가 기준: 절감효과 × 실행난이도(예시)');
  B.head(s,'우선 착수 대상 4개 과제 선정');
  B.sub(s,'효과 대·난이도 소 과제 우선 추진 — 우측 하단 영역은 당해연도 추진 유보');
  B.matrix(s,{ x:M+0.62, y:BT+0.30, w:5.85, h:3.72, yLabel:'절감\n효과\n↑', xLabel:'실행 난이도 →',
    quadrants:['즉시 착수','선별 추진','자동화 위임','추진 유보'],
    points:[{px:0.22,py:0.20,t:'문서 자동화',d:0.52},{px:0.36,py:0.34,t:'데이터 연계',d:0.44},
            {px:0.16,py:0.44,t:'서식 통일',d:0.34},{px:0.44,py:0.14,t:'판정 표준화',d:0.40},
            {px:0.68,py:0.30,t:'응대 이관',d:0.32},{px:0.58,py:0.66,t:'예외 처리',d:0.28},
            {px:0.82,py:0.78,t:'전면 재개발',d:0.36}] });
  const rx=M+7.0, rw=W-M-(M+7.0);
  B.sectionTitle(s,'우선순위 4개 과제',{ x:rx, y:BT, w:rw, sz:13 });
  [['문서 자동화','1,900h · 3개월'],['판정 표준화','1,500h · 4개월'],['데이터 연계','1,700h · 5개월'],['서식 통일','1,200h · 2개월']]
   .forEach(([n,d],i)=>{
    const y=BT+0.52+i*0.86;
    B.hr(s,{ x:rx, y, w:rw, color:K.g5 });
    B.txt(s,String(i+1),{ x:rx, y:y+0.12, w:0.3, h:0.28, sz:11, b:true, c:K.g2 });
    B.txt(s,n,{ x:rx+0.34, y:y+0.10, w:rw-0.34, h:0.30, sz:11.5, b:true });
    B.txt(s,d,{ x:rx+0.34, y:y+0.42, w:rw-0.34, h:0.28, sz:T.small, c:K.g2 });
  });
  B.callout(s,'4개 과제 합계 6,300시간 · 1년 내 완결',{ x:rx, y:BT+4.02, w:rw, h:0.44, dark:true, sz:10.5 });
}

/* 8 추진 체계 (조직도) */
{
  const s=S('03 실행 로드맵','');
  B.head(s,'전환추진단 중심의 3층 실행 체계 운영');
  B.sub(s,'의사결정·실행·지원을 분리하고, 부서 담당자를 실행층에 상시 배치');
  const cxm=M+CW/2;
  const top={ x:cxm-1.7, y:BT, w:3.4, h:0.58 };
  B.box(s,{ ...top, fill:K.ink, line:K.ink });
  B.txt(s,'경영진 협의체 (분기)',{ ...top, sz:T.body, b:true, c:K.w, align:'center', valign:'middle' });
  const mid={ x:cxm-1.7, y:BT+0.88, w:3.4, h:0.58 };
  B.box(s,{ ...mid, fill:K.g5, line:K.g2 });
  B.txt(s,'전환추진단 (월)',{ ...mid, sz:T.body, b:true, align:'center', valign:'middle' });
  s.addShape('line',{ x:cxm, y:top.y+top.h, w:0, h:mid.y-(top.y+top.h), line:{color:K.g3,width:0.75} });
  const cells=B.split(4);
  ['업무 표준화반','자동화 실행반','교육·이관반','성과 관리반'].forEach((t,i)=>{
    const y=BT+1.76, c=cells[i];
    B.box(s,{ x:c.x, y, w:c.w, h:0.56, fill:K.w, line:K.g4 });
    B.txt(s,t,{ x:c.x, y, w:c.w, h:0.56, sz:T.small, b:true, align:'center', valign:'middle' });
    s.addShape('line',{ x:c.x+c.w/2, y:mid.y+mid.h+0.18, w:0, h:y-(mid.y+mid.h+0.18), line:{color:K.g3,width:0.75} });
  });
  s.addShape('line',{ x:cells[0].x+cells[0].w/2, y:mid.y+mid.h+0.18, w:cells[3].x+cells[3].w/2-(cells[0].x+cells[0].w/2), h:0, line:{color:K.g3,width:0.75} });
  s.addShape('line',{ x:cxm, y:mid.y+mid.h, w:0, h:0.18, line:{color:K.g3,width:0.75} });
  B.tableNative(s,{ x:M, y:BT+2.66, colW:[2.2,4.2,3.4333], headH:0.40, rowH:0.44,
    head:['조직','역할','산출물'], align:['left','left','left'],
    rows:[['경영진 협의체','자원 배분·중단 판단','분기 심의 결과'],
          ['전환추진단','과제 관리·게이트 운영','월간 추진 보고'],
          ['실행반(4개)','과제 수행·현업 협의','과제별 산출물']] });
}

/* 9 간트 */
{
  const s=S('03 실행 로드맵','기간은 착수일 기준 상대 표기(예시)');
  B.head(s,'12개월 4개 분기 실행 체계 운영');
  B.sub(s,'분기별 산출물 1건 원칙 — 분기 말 확인 불가 과제는 로드맵 제외');
  B.gantt(s,{ x:M, y:BT, w:CW, labW:1.7, noteW:1.9, rowH:0.60, cols:['1분기','2분기','3분기','4분기'],
    rows:[{t:'업무 표준화',a:0.00,b:0.45,ms:'정의서 승인'},
          {t:'문서 자동화',a:0.20,b:0.55,ms:'전 부서 적용'},
          {t:'데이터 연계',a:0.35,b:0.75,ms:'API 오픈',tone:'mid'},
          {t:'교육·이관',a:0.50,b:1.00,ms:'이관율 90%',tone:'mid'},
          {t:'성과 측정',a:0.75,b:1.00,ms:'보정 완료',tone:'light'}] });
  const bY=BB-0.94;
  B.box(s,{ x:M, y:bY, w:CW, h:0.94, fill:K.g6, line:K.g4 });
  [['분기 산출물','정의서 → 적용 보고 →\nAPI 문서 → 성과 리포트'],['확인 주체','전환추진단 월 단위\n경영진 분기 단위'],
   ['중단 기준','2개 분기 연속 미달 시\n과제 재설계']].forEach(([k,v],i)=>{
    const x=M+0.24+i*(CW/3);
    B.txt(s,k,{ x, y:bY+0.13, w:CW/3-0.45, h:0.24, sz:T.small, b:true, c:K.g2 });
    B.txt(s,v,{ x, y:bY+0.40, w:CW/3-0.45, h:0.48, sz:10.5, b:true, lh:1.28 });
  });
}

/* 10 교육 체계 */
{
  const s=S('04 성과 관리','');
  B.head(s,'3단 교육 체계로 전환 역량 내재화');
  B.sub(s,'전 직원 공통 → 직무별 심화 → 추진단 전문 순으로 단계 구성');
  B.layers(s,{ x:M+0.9, y:BT, w:CW-1.8, h:2.46, taper:0.55,
    items:[{t:'전문 과정',d:'전환추진단 15명 · 과제 설계·게이트 운영 · 40시간'},
           {t:'심화 과정',d:'직무별 60명 · 표준 업무 단위 적용 실습 · 24시간'},
           {t:'공통 과정',d:'전 직원 320명 · AX 개념·업무 변화 이해 · 8시간'}] });
  B.tableNative(s,{ x:M, y:BT+2.70, colW:[2.0,1.6,1.6,2.2,2.4333], headH:0.40, rowH:0.44,
    head:['과정','대상','시간','핵심 내용','완료 기준'], align:['left','center','center','left','left'],
    rows:[['공통','320명','8h','AX 개념·업무 변화','수료율 90%'],
          ['심화','60명','24h','표준 업무 단위 실습','과제 제출'],
          ['전문','15명','40h','게이트 운영·성과 측정','인증 취득']] });
}

/* 11 성과 관리 */
{
  const s=S('04 성과 관리','측정 주기: 월 1회 / 보고: 분기 1회');
  B.head(s,'5개 지표로 전환 성과 상시 점검');
  B.sub(s,'선행지표(이관율·수료율)와 결과지표(절감시간·오류율)를 분리 관리');
  B.kpiRow(s,[{v:'6,300h',l:'연간 절감 목표'},{v:'90%',l:'업무 이관율'},
              {v:'-40%',l:'처리 오류율'},{v:'92%',l:'교육 수료율',dark:true}],{ y:BT, h:1.15 });
  B.tableNative(s,{ x:M, y:BT+1.36, colW:[2.6,1.5,1.5,1.5,2.7333], headH:0.38, rowH:0.44,
    head:['지표','구분','현재','목표','측정 방법'], align:['left','center','center','center','left'],
    rows:[['업무시간 절감','결과','0h','6,300h','업무 로그 집계'],
          ['업무 이관율','선행','12%','90%','부서 확인서'],
          ['처리 오류율','결과','8.4%','5.0%','품질 점검 표본'],
          ['교육 수료율','선행','—','92%','학습관리시스템'],
          ['표준 준수율','선행','—','85%','분기 감사']] });
  B.callout(s,'선행지표가 2개 분기 연속 미달할 경우 과제 재설계 절차 착수',{ x:M, y:BB-0.56, w:CW, h:0.56 });
}

/* 12 클로징 */
{
  const s=deck.bare({ bg:K.ink });
  s.addText('03',{ x:M, y:0.40, w:4, h:0.30, fontFace:B.FONT, fontSize:14, bold:true, color:K.g4, valign:'middle', margin:0 });
  s.addText('다음 단계',{ x:M, y:2.10, w:CW, h:0.9, fontFace:B.FONT, fontSize:40, bold:true, color:K.w, valign:'bottom', margin:0, charSpacing:-1 });
  s.addShape('line',{ x:M, y:3.15, w:B.textW('다음 단계',40), h:0, line:{color:K.w,width:2.5} });
  [['1주','전환추진단 구성 및 킥오프'],['4주','표준 업무 단위 정의서 초안'],['8주','파일럿 2개 부서 선정·착수']]
   .forEach(([k,v],i)=>{
    const y=3.70+i*0.78;
    s.addText(k,{ x:M, y, w:1.2, h:0.34, fontFace:B.FONT, fontSize:16, bold:true, color:K.g4, margin:0 });
    s.addText(v,{ x:M+1.4, y, w:CW-1.4, h:0.34, fontFace:B.FONT, fontSize:16, color:K.w, margin:0 });
    s.addShape('line',{ x:M, y:y+0.52, w:CW, h:0, line:{color:K.g1,width:0.75} });
  });
  s.addText('12',{ x:M, y:7.05, w:2, h:0.25, fontFace:B.FONT, fontSize:9.5, color:K.g3, valign:'middle', margin:0 });
}

deck.save(OUT).then(f=>console.log('wrote', f));
