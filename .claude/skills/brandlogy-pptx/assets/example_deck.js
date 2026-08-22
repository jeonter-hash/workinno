/**
 * example_deck.js — 헬퍼 사용 예시 (4장: 표지 · 패턴 A · 패턴 B · 섹션 디바이더)
 * 실행: node example_deck.js [로고.png] [출력.pptx]
 * 새 덱을 만들 때 이 파일을 복사해 내용만 갈아끼운다. 좌표는 절대 직접 쓰지 않는다.
 */
'use strict';
const path = require('path');
const B = require('./brandlogy.js');

const LOGO = process.argv[2] || null;                 // 사용자 제공 Brandlogy 누끼 PNG
const OUT = process.argv[3] || 'brandlogy_example.pptx';

(async () => {
  const deck = B.createDeck({ logo: LOGO, title: 'Brandlogy 예시 덱', author: 'Brandlogy' });

  // ── 1. 표지 (Hero Gradient 1/3)
  B.cover(deck, {
    chapter: 'Brandlogy',
    title: '검색 트래픽은 줄었지만, 브랜드 유입은 3년째 커지고 있다',
    subtitle: '2023–2025 자사몰 유입 구조 변화와 2026년 채널 전략',
    kpi: '+42.6%',
    kpiLabel: '브랜드 검색 유입 증가율 (2023 → 2025)',
    source: '출처: 자사 GA4 로그, 2026-01',
  });

  // ── 2. 패턴 A — KPI 스트립 + 차트
  {
    const s = deck.slide({ chapter: '01 유입 구조 진단', source: '출처: 자사 GA4 로그(2023-01~2025-12)' });
    B.headline(s, '성장은 채널 확대가 아니라 브랜드 검색 한 곳에서 나왔다');
    B.subtitle(s, '전체 세션은 정체했지만 브랜드 검색 세션만 연 20% 이상 증가, 전환율도 2.4배 높다');

    B.kpiRow(s, [
      { value: '+42.6%', label: '브랜드 검색 유입 (3년 누적)' },
      { value: '2.4x', label: '일반 검색 대비 전환율' },
      { value: '−8.1%', label: '비브랜드 검색 유입', color: B.C.sub },
      { value: '71%', label: '상위 3개 채널의 성장 기여도', gradient: true },  // 장표당 그라디언트 1개
    ]);

    const d = B.dataCard(s, {
      x: B.Z.body.x, y: B.BAND.A.detail.y, w: B.Z.body.w, h: B.BAND.A.detail.h,
      title: '채널별 세션 추이 (천 세션)',
      source: '출처: 자사 GA4 로그, 채널 정의는 2024년 개편 기준으로 소급 적용',
    });
    s.addChart('bar', [
      { name: '브랜드 검색', labels: ['2023', '2024', '2025'], values: [412, 498, 587] },
      { name: '비브랜드 검색', labels: ['2023', '2024', '2025'], values: [880, 842, 809] },
    ], B.chartOpts({ ...d.area, barDir: 'col', showLegend: true, legendPos: 'b',
      legendFontFace: 'Pretendard', legendFontSize: 10, legendColor: B.C.sub,
      chartColors: [B.C.brandBlue, B.C.blue200] }));
  }

  // ── 3. 패턴 B — 2열 비교 + So What
  {
    const s = deck.slide({ chapter: '02 2026 채널 전략', source: '출처: 내부 시뮬레이션(2026-01)' });
    B.headline(s, '예산을 채널 수가 아니라 브랜드 자산에 다시 배분해야 한다');
    B.subtitle(s, '동일 예산에서 브랜드 검색 비중을 12%p 올리면 CAC는 19% 내려간다');

    const [L, Rc] = B.split(2);
    const band = B.BAND.B.cols;
    B.card(s, { ...L, ...band, kind: 'standard' });
    B.h2(s, '현재 배분의 문제', { x: L.x + B.px(20), y: band.y + B.px(20), w: L.w - 2 * B.px(20) });
    B.bullets(s, [
      '9개 채널에 예산을 고르게 나눠 어느 채널도 임계 노출량을 넘기지 못함',
      '하위 4개 채널이 예산의 31%를 쓰고 전환의 6%만 기여',
      '브랜드 검색은 성과가 가장 좋은데 예산 비중은 14%에 정체',
    ], { x: L.x + B.px(20), y: band.y + 0.75, w: L.w - 2 * B.px(20), h: band.h - 0.95 });

    const d = B.dataCard(s, { ...Rc, ...band, title: '예산 재배분 시 CAC 변화 (원)',
      source: '출처: 2025년 채널별 CAC에 탄력성 −0.42를 적용한 내부 추정' });
    s.addChart('bar', [{ name: 'CAC', labels: ['현행', '재배분(안)'], values: [38400, 31100] }],
      B.chartOpts({ ...d.area, barDir: 'col' }));

    B.soWhat(s, 'So What — 채널을 늘리는 대신 하위 4개를 정리하고 그 예산을 브랜드 검색·리타게팅에 이관한다.');
  }

  // ── 4. 섹션 디바이더 (Hero Gradient 2/3)
  B.divider(deck, { number: '03', title: '실행 로드맵', lead: '1분기 정리, 2분기 이관, 3분기 성과 검증', gradient: true });

  await deck.save(path.resolve(OUT));
  console.log('wrote', OUT);
})();
