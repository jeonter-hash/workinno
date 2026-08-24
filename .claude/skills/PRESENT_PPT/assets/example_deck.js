/** 발표용 샘플 12장 — 큰 글자·적은 항목·도식 중심 */
'use strict';
const B = require('./ksa_mono.js');
const { K, M, CW, BT, BB, W } = B;
const OUT = process.argv[2] || 'present_sample.pptx';
const deck = B.createDeck({ mode:'present', palette:'accent', title:'AX 전환 로드맵',
  docTitle:'AX 전환 로드맵', author:'한국표준협회' });
const T = deck.T, C = deck.C;
const S = (ch) => deck.slide({ chapter:ch });

/* 1 표지 · 2 목차 */
// 표지 — 사진 4장(띠) + 대형 사진 1장. photos를 생략하면 자리표시가 들어가고,
// PowerPoint에서 각 사진을 [그림 바꾸기]로 갈아끼우면 사선이 그대로 유지된다.
B.cover(deck, { org:'○○공단', title:'AX 전환 로드맵\n(최종보고)',
  date:'2026.01', by:'한국표준협회' });
B.toc(deck, { items:[
  { n:'01', t:'현황 진단', p:3 }, { t:'업무시간 구조와 절감 여력', p:3 }, { t:'절감 경로별 기여도', p:4 }, { t:'3대 경로 분해', p:5 },
  { n:'02', t:'전환 과제', p:6 }, { t:'5단계 게이트와 통과 기준', p:6 },
  { n:'03', t:'실행 로드맵', p:7 }, { t:'우선 착수 4개 과제', p:7 }, { t:'추진 체계와 역할', p:8 }, { t:'12개월 실행 일정', p:9 },
  { n:'04', t:'성과 관리', p:10 }, { t:'3단 교육 체계', p:10 }, { t:'성과지표 5종', p:11 },
]});

/* 3 핵심 요약 */
{
  const s=S('요약');
  B.head(s,'절감 여력 8,100시간, 상위 4개 업무에 65% 집중');
  B.sub(s,'전면 착수보다 상위 4개에 자원을 모으는 편이 효과적 — 1년이면 6,300시간');
  B.kpiRow(s,[{v:'12,400h',l:'연간 총 업무시간'},{v:'8,100h',l:'정형·반복 구간'},
              {v:'4개',l:'절감 상위 업무'},{v:'65%',l:'상위 4개 집중도',dark:true}],{ y:BT, h:1.20 });
  const cells=B.split(3);
  // panel() — 제목은 가운데 정렬, 항목은 박스 높이에 맞춰 고르게 분산된다
  [['01  진단',['12개 업무의 시간을 실측','절감 경로 세 갈래를 확인']],
   ['02  설계',['5단계 게이트를 구성','통과 기준을 사전 합의']],
   ['03  실행',['상위 4개 과제부터 착수','분기마다 산출물 1건을 확인']]].forEach(([t,items],i)=>{
    B.panel(s,{ x:cells[i].x, y:BT+1.52, w:cells[i].w, h:1.72, title:t, items });
  });
  B.callout(s,'상위 4개에 자원을 모으면 1년 안에 목표의 82%에 도달 — 나머지는 2차 연도로 이월',{ x:M, y:BT+3.44, w:CW, h:0.72, dark:true });
  B.footnote(s,'주) 절감 시간은 12개 업무의 실측값이며, 인건비 환산액은 포함하지 않습니다.',{ x:M, y:BB-0.24, w:CW });
}

/* 4 워터폴 */
{
  const s=S('01 현황 진단');
  B.head(s,'절감 52천 시간 가운데 21천 시간이 문서 자동화 한 경로에 몰려 있습니다');
  B.sub(s,'네 경로를 모두 실행해야 목표 72천 시간에 닿습니다');
  B.box(s,{ x:M, y:BT, w:CW, h:4.65, fill:K.w, line:K.g4 });
  B.sectionTitle(s,'업무시간 절감 경로 (천 시간)',{ x:M+0.24, y:BT+0.18, w:5.2 });
  const axB=B.waterfall(s,{ x:M+0.4, y:BT+0.86, w:CW-0.8, h:2.70, max:130,
    steps:[{t:'현재',v:124,base:true},{t:'문서 자동화',v:-21},{t:'데이터 연계',v:-14},
           {t:'판정 표준화',v:-11},{t:'교육 이관',v:-6},{t:'목표',v:72,base:true}] });
  B.hr(s,{ x:M+0.24, y:axB+0.14, w:CW-0.48, color:K.g4 });
  B.txt(s,'상위 4개 업무만으로 목표의 82%에 도달 — 나머지는 2차 연도로 이월',
    { x:M+0.24, y:axB+0.22, w:CW-0.48, h:0.32, sz:T.body, b:true, c:K.g1 });
}

/* 5 드라이버 트리 */
{
  const s=S('01 현황 진단');
  B.head(s,'절감 시간의 3대 경로: 자동화·표준화·이관');
  B.sub(s,'절감으로 인정하는 범위는 이 세 경로로 한정합니다');
  B.box(s,{ x:M, y:BT, w:6.4, h:4.65, fill:K.w, line:K.g5 });
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
  ['자동화 — 이미 검증된 도구여서 3개월이면 적용을 마칩니다','표준화 — 부서 합의가 병목이므로 게이트로 끊어 관리합니다','이관 — 교육을 함께 하지 않으면 예전 방식으로 되돌아갑니다']
   .forEach((t,i)=>{
    const y=BT+i*1.62;
    B.box(s,{ x:rx, y, w:rw, h:1.41, fill:i===1?C.acc:K.g6, line:i===1?C.acc:K.g4 });
    B.txt(s,`0${i+1}`,{ x:rx+0.2, y:y+0.14, w:0.6, h:0.28, sz:11, b:true, c:i===1?K.g4:K.g2 });
    B.txt(s,t,{ x:rx+0.2, y:y+0.46, w:rw-0.4, h:0.78, sz:T.body, b:true, c:i===1?K.w:K.ink, lh:1.35 });
  });
}

/* 6 체브론 + 게이트표 */
{
  const s=S('02 전환 과제');
  B.head(s,'5단계로 끊고 단계마다 통과 기준을 확인합니다');
  B.sub(s,'앞 단계의 산출물이 다음 단계의 입력 — 통과하지 못한 과제는 다음으로 넘기지 않습니다');
  B.chevrons(s,[{t:'1. 진단',d:'업무 목록·시간 실측'},{t:'2. 설계',d:'표준 업무 단위 확정'},
                {t:'3. 시범',d:'2개 부서 파일럿'},{t:'4. 확산',d:'전사 이관·교육',tone:'mid'},
                {t:'5. 정착',d:'성과 측정·보정',tone:'light'}],{ x:M, y:BT, w:CW, h:0.84 });
  B.tableNative(s,{ x:M, y:BT+1.80, colW:[4.6,1.6,1.6,2.0333], headH:0.44, rowH:0.54,
    head:['통과 기준','난이도','자원 소요','선행 조건'], align:['left','center','center','center'], boldCol:[3],
    rows:[['업무 목록 100% 실측 완료','●','◐','—'],
          ['표준 업무 단위 정의서 승인','●','●','진단 완료'],
          ['파일럿 2개 부서 KPI 달성','◐','●','설계 승인'],
          ['전 부서 이관율 90% 이상','●','●','파일럿 성공']] });
  B.footnote(s,'● 높음   ◐ 보통   ○ 낮음',{ x:M, y:BB-0.22, w:4 });
}

/* 7 2×2 매트릭스 */
{
  const s=S('03 실행 로드맵');
  B.head(s,'효과가 크고 난이도가 낮은 4개를 먼저 착수합니다');
  B.sub(s,'효과가 작고 난이도가 높은 우측 하단 영역은 올해 착수하지 않습니다');
  B.matrix(s,{ x:M+0.62, y:BT+0.24, w:5.85, h:3.95, yLabel:'절감\n효과\n↑', xLabel:'실행 난이도 →',
    quadrants:['즉시 착수','선별 추진','자동화 위임','추진 유보'],
    points:[{px:0.22,py:0.20,t:'문서 자동화',d:0.52},{px:0.36,py:0.34,t:'데이터 연계',d:0.44},
            {px:0.16,py:0.44,t:'서식 통일',d:0.34},{px:0.44,py:0.14,t:'판정 표준화',d:0.40},
            {px:0.68,py:0.30,t:'응대 이관',d:0.32},{px:0.58,py:0.66,t:'예외 처리',d:0.28},
            {px:0.82,py:0.78,t:'전면 재개발',d:0.36}] });
  const rx=M+7.0, rw=W-M-(M+7.0);
  B.sectionTitle(s,'우선순위 4개 과제',{ x:rx, y:BT, w:rw, sz:13 });
  [['문서 자동화','1,900h · 3개월'],['판정 표준화','1,500h · 4개월'],['데이터 연계','1,700h · 5개월'],['서식 통일','1,200h · 2개월']]
   .forEach(([n,d],i)=>{
    const y=BT+0.56+i*0.92;
    B.hr(s,{ x:rx, y, w:rw, color:K.g5 });
    B.txt(s,String(i+1),{ x:rx, y:y+0.12, w:0.3, h:0.28, sz:11, b:true, c:K.g2 });
    B.txt(s,n,{ x:rx+0.34, y:y+0.10, w:rw-0.34, h:0.30, sz:11.5, b:true });
    B.txt(s,d,{ x:rx+0.34, y:y+0.42, w:rw-0.34, h:0.28, sz:T.small, c:K.g2 });
  });
  B.txt(s,'4개 합계 6,300시간 · 1년 내 완결',{ x:rx, y:BT+4.02, w:rw, h:0.44, sz:10.5, b:true, c:K.g1, align:'center', valign:'middle' });
  B.hr(s,{ x:rx, y:BT+4.02, w:rw, color:K.g4 });
}

/* 8 추진 체계 (조직도) */
{
  const s=S('03 실행 로드맵');
  B.head(s,'전환추진단을 가운데 두고 세 층으로 나눠 운영합니다');
  B.sub(s,'판단하는 자리와 실행하는 자리를 분리하고, 부서 담당자를 실행층에 상시 배치합니다');
  const cxm=M+CW/2;
  const top={ x:cxm-1.7, y:BT, w:3.4, h:0.62 };
  B.box(s,{ ...top, fill:C.dark, line:C.dark });
  B.txt(s,'경영진 협의체 (분기)',{ ...top, sz:T.body, b:true, c:K.w, align:'center', valign:'middle' });
  const mid={ x:cxm-1.7, y:BT+0.94, w:3.4, h:0.62 };
  B.box(s,{ ...mid, fill:K.g5, line:K.g2 });
  B.txt(s,'전환추진단 (월)',{ ...mid, sz:T.body, b:true, align:'center', valign:'middle' });
  s.addShape('line',{ x:cxm, y:top.y+top.h, w:0, h:mid.y-(top.y+top.h), line:{color:K.g3,width:0.75} });
  const cells=B.split(4);
  ['업무 표준화반','자동화 실행반','교육·이관반','성과 관리반'].forEach((t,i)=>{
    const y=BT+1.88, c=cells[i];
    B.box(s,{ x:c.x, y, w:c.w, h:0.62, fill:K.w, line:K.g4 });
    B.txt(s,t,{ x:c.x, y, w:c.w, h:0.62, sz:T.small, b:true, align:'center', valign:'middle' });
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
  const s=S('03 실행 로드맵');
  B.head(s,'12개월을 네 분기로 끊고 분기마다 산출물 하나를 확인합니다');
  B.sub(s,'분기 말에 확인할 산출물이 없는 과제는 로드맵에 넣지 않습니다');
  B.gantt(s,{ x:M, y:BT, w:CW, labW:1.7, noteW:1.9, rowH:0.68, cols:['1분기','2분기','3분기','4분기'],
    rows:[{t:'업무 표준화',a:0.00,b:0.45,ms:'정의서 승인'},
          {t:'문서 자동화',a:0.20,b:0.55,ms:'전 부서 적용'},
          {t:'데이터 연계',a:0.35,b:0.75,ms:'API 오픈',tone:'mid'},
          {t:'교육·이관',a:0.50,b:1.00,ms:'이관율 90%',tone:'mid'},
          {t:'성과 측정',a:0.75,b:1.00,ms:'보정 완료',tone:'light'}] });
  const bY=BB-0.94;
  B.box(s,{ x:M, y:bY, w:CW, h:0.94, fill:K.g6, line:K.g4 });
  [['분기 산출물','정의서 → 적용 보고 →\nAPI 문서 → 성과 리포트'],['확인 주체','전환추진단 월 단위\n경영진 분기 단위'],
   ['중단 기준','두 분기 연속 미달하면\n과제를 다시 설계']].forEach(([k,v],i)=>{
    const x=M+0.24+i*(CW/3);
    B.txt(s,k,{ x, y:bY+0.13, w:CW/3-0.45, h:0.24, sz:T.small, b:true, c:K.g2 });
    B.txt(s,v,{ x, y:bY+0.40, w:CW/3-0.45, h:0.48, sz:10.5, b:true, lh:1.28 });
  });
}

/* 10 교육 체계 */
{
  const s=S('04 성과 관리');
  B.head(s,'공통·심화·전문 세 단계로 역량을 쌓습니다');
  B.sub(s,'전 직원이 같은 개념을 익힌 뒤 직무별로 심화하고, 추진단은 운영까지 배웁니다');
  B.layers(s,{ x:M+0.9, y:BT, w:CW-1.8, h:2.70, taper:0.55,
    items:[{t:'전문 과정',d:'전환추진단 15명 · 과제 설계·게이트 운영 · 40시간'},
           {t:'심화 과정',d:'직무별 60명 · 표준 업무 단위 적용 실습 · 24시간'},
           {t:'공통 과정',d:'전 직원 320명 · AX 개념·업무 변화 이해 · 8시간'}] });
  B.tableNative(s,{ x:M, y:BT+2.86, colW:[2.0,1.6,1.6,2.2,2.4333], headH:0.44, rowH:0.44,
    head:['과정','대상','시간','핵심 내용','완료 기준'], align:['left','center','center','left','left'],
    rows:[['공통','320명','8h','AX 개념·업무 변화','수료율 90%'],
          ['심화','60명','24h','표준 업무 단위 실습','과제 제출'],
          ['전문','15명','40h','게이트 운영·성과 측정','인증 취득']] });
}

/* 11 성과 관리 */
{
  const s=S('04 성과 관리');
  B.head(s,'다섯 개 지표로 성과를 상시 점검합니다');
  B.sub(s,'미리 움직이는 선행지표(이관율·수료율)와 뒤따라오는 결과지표(절감시간·오류율)를 나눠 봅니다');
  B.kpiRow(s,[{v:'6,300h',l:'연간 절감 목표'},{v:'90%',l:'업무 이관율'},
              {v:'-40%',l:'처리 오류율'},{v:'92%',l:'교육 수료율',dark:true}],{ y:BT, h:1.15 });
  B.tableNative(s,{ x:M, y:BT+1.36, colW:[2.6,1.5,1.5,1.5,2.7333], headH:0.38, rowH:0.44,
    head:['지표','구분','현재','목표','측정 방법'], align:['left','center','center','center','left'],
    rows:[['업무시간 절감','결과','0h','6,300h','업무 로그 집계'],
          ['업무 이관율','선행','12%','90%','부서 확인서'],
          ['처리 오류율','결과','8.4%','5.0%','품질 점검 표본'],
          ['교육 수료율','선행','—','92%','학습관리시스템'],
          ['표준 준수율','선행','—','85%','분기 감사']] });
  B.callout(s,'선행지표가 두 분기 연속 목표에 미치지 못하면 과제를 다시 설계합니다.',{ x:M, y:BB-0.58, w:CW, h:0.58 });
}

deck.save(OUT).then(f=>console.log('wrote', f));
