/**
 * report-weekly.js
 * 任意のアカウント設定 + 週次データ → 週次レポートDOCX
 *
 * 使い方:
 *   const { generateWeeklyReport } = require('./lib/report-weekly');
 *   generateWeeklyReport(account, weekData, outputPath);
 */
const { Document, Packer, PageBreak, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign } = require('docx');
const fs   = require('fs');
const path = require('path');
const {
  C, FONT, BD_NONE, BD_GRAY,
  r, p, sp, hr, h1, h2, th, td,
  kpiRow, lbox, makeHeaderFooter, PAGE_PROPS,
} = require('./docx-base');

// ────────────────────────────────────────
// 01 表紙
// ────────────────────────────────────────
function buildCover(account, weekData) {
  return [
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({ height: { value: 120, rule: 'exact' }, children: [
        new TableCell({ borders: BD_NONE, width: { size: 9360, type: WidthType.DXA },
          shading: { fill: C.primary, type: ShadingType.CLEAR }, children: [p('')] }),
      ]})],
    }),
    sp(400),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before:0, after:80 },
      children: [r('Threads アカウント運用', { color: C.gray400, size: 22 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before:0, after:60 },
      children: [r('週次レポート', { bold: true, color: C.primary, size: 72 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before:0, after:360 },
      children: [r('Weekly Performance Report', { color: C.gray400, size: 20 })] }),
    hr(C.accent, 6), sp(240),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({ children: [new TableCell({
        borders: BD_NONE, width: { size: 9360, type: WidthType.DXA },
        shading: { fill: C.light, type: ShadingType.CLEAR },
        margins: { top: 200, bottom: 200, left: 360, right: 360 },
        children: [
          p(r(account.handle, { bold: true, color: C.primary, size: 36 }), { after: 60 }),
          p(r(`${account.location}｜${account.business_type}｜目標：${account.goal}`, { size: 19, color: C.gray700 }), { after: 0 }),
        ],
      })]})],
    }),
    sp(280),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680],
      rows: [new TableRow({ children: [
        new TableCell({ borders: BD_NONE, width: { size: 4680, type: WidthType.DXA },
          children: [
            p(r('レポート期間', { size: 17, color: C.gray400 }), { after: 40 }),
            p(r(weekData.periodLabel, { bold: true, size: 21 }), { after: 0 }),
          ]
        }),
        new TableCell({ borders: BD_NONE, width: { size: 4680, type: WidthType.DXA },
          children: [
            p(r('作成日', { size: 17, color: C.gray400 }), { after: 40 }),
            p(r(todayStr(), { bold: true, size: 21 }), { after: 0 }),
          ]
        }),
      ]})],
    }),
    sp(80),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680],
      rows: [new TableRow({ children: [
        new TableCell({ borders: BD_NONE, width: { size: 4680, type: WidthType.DXA },
          children: [
            p(r('担当ディレクター', { size: 17, color: C.gray400 }), { after: 40 }),
            p(r(account.director || '未設定', { bold: true, size: 21 }), { after: 0 }),
          ]
        }),
        new TableCell({ borders: BD_NONE, width: { size: 4680, type: WidthType.DXA },
          children: [
            p(r('担当運用者', { size: 17, color: C.gray400 }), { after: 40 }),
            p(r(account.operator || '未設定', { bold: true, size: 21 }), { after: 0 }),
          ]
        }),
      ]})],
    }),
    sp(400), hr(C.gray200), sp(60),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [
      r('CONFIDENTIAL  ｜  本資料は秘密情報を含みます。無断転載・共有はお控えください。', { color: C.gray400, size: 16 }),
    ]}),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ────────────────────────────────────────
// 02 エグゼクティブサマリー
// ────────────────────────────────────────
function buildSummary(account, weekData) {
  const { kpi, prevKpi } = weekData;
  const diff = (a, b) => {
    const d = a - b;
    return (d > 0 ? '▲ +' : d < 0 ? '▼ ' : '± ') + Math.abs(d).toLocaleString();
  };
  const diffColor = (a, b) => a > b ? C.green : a < b ? C.red : C.gray400;

  return [
    h1('01', 'エグゼクティブサマリー'),
    sp(80),
    kpiRow([
      { label: '今週 投稿数',               value: String(kpi.total),              unit: '件',  sub: `先週比 ${diff(kpi.total, prevKpi.total)}`,   subColor: diffColor(kpi.total, prevKpi.total) },
      { label: '今週 いいね合計',            value: kpi.likes.toLocaleString(),     unit: '',    sub: `先週比 ${diff(kpi.likes, prevKpi.likes)}`,    subColor: diffColor(kpi.likes, prevKpi.likes) },
      { label: '今週 コメント合計',          value: String(kpi.comments),           unit: '',    sub: `先週比 ${diff(kpi.comments, prevKpi.comments)}`, subColor: diffColor(kpi.comments, prevKpi.comments) },
      { label: '今週 エンゲージメントスコア', value: kpi.score.toLocaleString(),     unit: 'pt', sub: `先週比 ${diff(kpi.score, prevKpi.score)}`,    subColor: diffColor(kpi.score, prevKpi.score) },
    ]),
    sp(100),
    p(r('※ エンゲージメントスコア = いいね×1 + コメント×3 + リポスト×5', { size: 16, color: C.gray400 }), { after: 0 }),
    sp(80),
    h2('今週のハイライト'),
    sp(60),
    buildHighlights(kpi, prevKpi),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

function buildHighlights(kpi, prevKpi) {
  const top = kpi.top3[0];
  const rows = [];

  if (kpi.total >= prevKpi.total) {
    rows.push({ icon: '✅', label: '投稿頻度の維持', detail: `${kpi.total}件の投稿を実施。安定した運用継続中。`, color: C.light });
  } else {
    rows.push({ icon: '⚠️', label: '投稿数の減少', detail: `先週比 ${prevKpi.total - kpi.total}件減少。スケジュールを確認してください。`, color: 'FEF3CD' });
  }
  if (top && top.score > 10) {
    rows.push({ icon: '🔥', label: 'バズ投稿あり', detail: `スコア${top.score}pt「${top.content.slice(0, 30)}…」が高反応。`, color: C.light });
  } else {
    rows.push({ icon: '⚠️', label: 'バズ投稿なし', detail: 'スコア10pt超の投稿がありませんでした。コメント誘発・共感型投稿を強化してください。', color: 'FEF3CD' });
  }

  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [280, 9080],
    rows: rows.map((row, i) => new TableRow({ children: [
      new TableCell({ borders: BD_NONE, width: { size: 280, type: WidthType.DXA },
        shading: { fill: row.color, type: ShadingType.CLEAR },
        margins: { top: i === 0 ? 120 : 60, bottom: 0, left: 80, right: 80 },
        children: [p(r(row.icon, { size: 20 }), { align: AlignmentType.CENTER, after: 0 })],
      }),
      new TableCell({ borders: BD_NONE, width: { size: 9080, type: WidthType.DXA },
        margins: { top: i === 0 ? 120 : 60, bottom: 0, left: 160, right: 80 },
        children: [p([
          r(row.label + '：', { bold: true, size: 20, color: C.primary }),
          r(row.detail, { size: 20 }),
        ], { after: 0 })],
      }),
    ]})),
  });
}

// ────────────────────────────────────────
// 03 KPIダッシュボード
// ────────────────────────────────────────
function buildKPI(account, weekData) {
  const { kpi, prevKpi } = weekData;
  const rows = [
    ['投稿数',                kpi.total,     prevKpi.total,     account.weekly_post_count || 21],
    ['いいね合計',             kpi.likes,     prevKpi.likes,     '—'],
    ['コメント合計',           kpi.comments,  prevKpi.comments,  '—'],
    ['リポスト合計',           kpi.reposts,   prevKpi.reposts,   '—'],
    ['エンゲージメントスコア', kpi.score,     prevKpi.score,     '—'],
    ['いいね平均（1投稿）',    kpi.total ? +(kpi.likes/kpi.total).toFixed(1) : 0,
                              prevKpi.total ? +(prevKpi.likes/prevKpi.total).toFixed(1) : 0, '5以上'],
    [account.kpi_label || 'DM問い合わせ数', '—', '—', account.kpi_target || '—'],
  ];
  const cols = [3200, 1600, 1600, 1440, 1520];

  return [
    h1('02', 'KPIダッシュボード'),
    sp(80),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: cols,
      rows: [
        new TableRow({ children: [th('指標',cols[0]), th('今週',cols[1]), th('先週',cols[2]), th('前週比',cols[3]), th('目標',cols[4])] }),
        ...rows.map(([label, thisW, lastW, target]) => {
          const diff = typeof thisW === 'number' && typeof lastW === 'number' ? thisW - lastW : null;
          const diffColor = diff === null ? C.gray400 : diff > 0 ? C.green : diff < 0 ? C.red : C.gray400;
          const diffText  = diff === null ? '—' : (diff > 0 ? '+' : '') + diff.toLocaleString();
          return new TableRow({ children: [
            td(label, cols[0], { fill: C.gray100 }),
            td(String(thisW), cols[1], { bold: true, color: C.primary, align: AlignmentType.CENTER }),
            td(String(lastW), cols[2], { color: C.gray700, align: AlignmentType.CENTER }),
            td(diffText, cols[3], { bold: true, color: diffColor, align: AlignmentType.CENTER }),
            td(String(target), cols[4], { color: target === '—' ? C.gray400 : C.accent, align: AlignmentType.CENTER }),
          ]});
        }),
      ],
    }),
    sp(80),
    p(r('※ DM問い合わせ数はオーナー様からの報告値を入力。毎週月曜朝のフォロワー数スクリーンショット共有もお願いします。', { size: 16, color: C.gray400 }), { after: 0 }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ────────────────────────────────────────
// 04 投稿実績一覧
// ────────────────────────────────────────
function buildPostList(weekData) {
  const { posts } = weekData;
  const typeColors = {
    'インサイト祭り': 'E2D9F3',
    'サロン経営Tips': 'FDDCDC',
    'あるある・共感': 'D1ECF1',
    '雑談・日常':     'FFF3CD',
  };
  const cols = [840, 520, 760, 4720, 640, 640, 640];

  const rows_data = posts.map(post => {
    const dateObj = post.date;
    const dateLabel = dateObj ? `${dateObj.getMonth()+1}/${dateObj.getDate()}` : '—';
    const timeLabel = (post.dateStr || '').match(/(\d{2}:\d{2})/) ? post.dateStr.match(/(\d{2}:\d{2})/)[1] : '—';
    const rowFill = post.score >= 3 ? 'F0FFF4' : null;
    return new TableRow({ children: [
      td(dateLabel, cols[0], { fill: rowFill, align: AlignmentType.CENTER, color: C.gray700 }),
      td(timeLabel, cols[1], { fill: rowFill, align: AlignmentType.CENTER, color: C.gray700 }),
      td(post.type, cols[2], { fill: typeColors[post.type] || C.gray100, align: AlignmentType.CENTER }),
      td(post.content.length > 50 ? post.content.slice(0, 50) + '…' : post.content, cols[3], { fill: rowFill }),
      td(String(post.likes),    cols[4], { fill: rowFill, align: AlignmentType.CENTER, bold: post.likes >= 3,    color: post.likes >= 3 ? C.green : C.black }),
      td(String(post.comments), cols[5], { fill: rowFill, align: AlignmentType.CENTER, bold: post.comments >= 1, color: post.comments >= 1 ? C.accent : C.gray400 }),
      td(String(post.reposts),  cols[6], { fill: rowFill, align: AlignmentType.CENTER, color: C.gray400 }),
    ]});
  });

  return [
    h1('03', `投稿実績一覧（${posts.length}件）`),
    sp(80),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: cols,
      rows: [
        new TableRow({ children: [th('日付',cols[0]), th('時刻',cols[1]), th('種別',cols[2]), th('投稿内容（概要）',cols[3], AlignmentType.LEFT), th('❤️',cols[4]), th('💬',cols[5]), th('🔁',cols[6])] }),
        ...rows_data,
      ],
    }),
    sp(60),
    p(r('■ 緑ハイライト行 = エンゲージメントスコア3pt以上', { size: 16, color: C.gray400 }), { after: 0 }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ────────────────────────────────────────
// 05 バズ投稿TOP3
// ────────────────────────────────────────
function buildBuzz(weekData) {
  const { kpi } = weekData;
  const top3 = kpi.top3;
  const rankColors = [C.orange, '888888', '8B5E1A'];
  const medals = ['🥇', '🥈', '🥉'];

  const cards = top3.map((post, idx) => {
    const bc = rankColors[idx];
    const lines = post.content.split('\n').filter(l => l.trim()).slice(0, 5);
    return [
      new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
        rows: [
          new TableRow({ children: [new TableCell({
            borders: { top:{style:BorderStyle.SINGLE,size:6,color:bc}, bottom:BD_NONE.bottom, left:{style:BorderStyle.SINGLE,size:6,color:bc}, right:{style:BorderStyle.SINGLE,size:6,color:bc} },
            shading: { fill: C.gray100, type: ShadingType.CLEAR },
            margins: { top:80, bottom:80, left:200, right:200 },
            children: [p([
              r(`${medals[idx]} 第${idx+1}位`, { bold:true, color:bc, size:22 }),
              r(`　${post.dateStr || ''}　　`, { size:18, color:C.gray700 }),
              r(`❤️ ${post.likes.toLocaleString()}　💬 ${post.comments}　🔁 ${post.reposts}　🏆 ${post.score.toLocaleString()}pt`, { bold:true, color:C.primary, size:18 }),
            ], { after:0 })],
          })]})
          ,
          new TableRow({ children: [new TableCell({
            borders: { top:BD_NONE.top, bottom:{style:BorderStyle.SINGLE,size:6,color:bc}, left:{style:BorderStyle.SINGLE,size:6,color:bc}, right:{style:BorderStyle.SINGLE,size:6,color:bc} },
            margins: { top:60, bottom:80, left:200, right:200 },
            children: lines.map((line, i) => p(r(line, { size:18 }), { after: i<lines.length-1 ? 40:0, line:300 })),
          })]})
        ],
      }),
      sp(100),
    ];
  });

  return [
    h1('04', 'バズ投稿 TOP 3'),
    sp(80),
    ...cards.flat(),
    top3.length === 0 ? p(r('今週はバズ投稿がありませんでした。', { color: C.gray400 })) : sp(0),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ────────────────────────────────────────
// 06 コンテンツ種別分析
// ────────────────────────────────────────
function buildContentAnalysis(weekData) {
  const { kpi } = weekData;
  const types = Object.entries(kpi.byType).sort((a,b) => b[1].count - a[1].count);
  const typeColors = { 'インサイト祭り':'E2D9F3', 'サロン経営Tips':'FDDCDC', 'あるある・共感':'D1ECF1', '雑談・日常':'FFF3CD' };
  const cols = [2200, 760, 1200, 1200, 1200, 2800];

  return [
    h1('05', 'コンテンツ種別パフォーマンス'),
    sp(80),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: cols,
      rows: [
        new TableRow({ children: [th('種別',cols[0]), th('件数',cols[1]), th('いいね合計',cols[2]), th('いいね平均',cols[3]), th('スコア平均',cols[4]), th('備考',cols[5])] }),
        ...types.map(([type, d]) => new TableRow({ children: [
          td(type, cols[0], { fill: typeColors[type] || C.gray100 }),
          td(String(d.count), cols[1], { align: AlignmentType.CENTER }),
          td(String(d.likes), cols[2], { align: AlignmentType.CENTER }),
          td(String(d.avgLikes), cols[3], { align: AlignmentType.CENTER, bold: d.avgLikes >= 2, color: d.avgLikes >= 2 ? C.green : C.black }),
          td(String(d.avgScore), cols[4], { align: AlignmentType.CENTER, bold: d.avgScore >= 2, color: d.avgScore >= 2 ? C.green : C.black }),
          td('—', cols[5], { color: C.gray400 }),
        ]})),
      ],
    }),
    sp(120),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ────────────────────────────────────────
// 07 課題・改善提案（自動生成）
// ────────────────────────────────────────
function buildIssues(account, weekData) {
  const { kpi, prevKpi } = weekData;
  const issues = [];

  if (kpi.score < prevKpi.score * 0.7) {
    issues.push({ p:'HIGH', title:'エンゲージメント大幅低下', detail:`スコアが先週比 ${Math.round((1 - kpi.score/(prevKpi.score||1))*100)}%減少。バズ投稿の種別を分析し、再現性の高いコンテンツを増やしてください。`, action:'バズ実績のある種別の投稿を来週 +2〜3本追加する。' });
  }
  if (kpi.comments === 0) {
    issues.push({ p:'HIGH', title:'コメント誘発の仕掛け不足', detail:'今週のコメント数が0件。コメントはスコア×3の重みを持つため、最優先で改善が必要です。', action:'全投稿の末尾に「〜な方コメントで教えてください🙌」を追加。' });
  }
  if (kpi.total < (account.weekly_post_count || 21)) {
    issues.push({ p:'MED', title:'投稿数不足', detail:`目標${account.weekly_post_count || 21}件に対し実績${kpi.total}件。${(account.weekly_post_count||21) - kpi.total}件不足しています。`, action:'投稿ストックを事前に作成し、投稿予約を活用する。' });
  }
  if (kpi.zeroLikeRate > 60) {
    issues.push({ p:'MED', title:'0いいね投稿が多い', detail:`今週の${kpi.zeroLikeRate}%の投稿がいいね0件。コンテンツの質・タイミングを見直してください。`, action:'配信時間帯の最適化と、共感ネタ・問いかけ型投稿の比率を上げる。' });
  }

  if (issues.length === 0) {
    issues.push({ p:'LOW', title:'大きな課題なし', detail:'今週は安定した運用でした。次のフェーズとして、DMへの誘導率を高める施策を検討しましょう。', action:`「${account.kpi_label}」の計測を強化する。` });
  }

  const pStyle = { HIGH:{fill:'FDDCDC',color:C.red}, MED:{fill:'FFF3CD',color:C.orange}, LOW:{fill:'D4EDDA',color:C.green} };
  const cols = [600, 8760];

  return [
    h1('06', '課題・改善提案'),
    sp(80),
    ...issues.map((issue, idx) => {
      const s = pStyle[issue.p];
      return [
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: cols,
          rows: [
            new TableRow({ children: [
              new TableCell({ borders: BD_NONE, width:{size:cols[0],type:WidthType.DXA}, shading:{fill:s.fill,type:ShadingType.CLEAR},
                margins:{top:120,bottom:120,left:80,right:80}, verticalAlign:VerticalAlign.CENTER,
                rowSpan: 3,
                children: [
                  p(r(issue.p, { bold:true, color:s.color, size:16 }), { align:AlignmentType.CENTER, after:20 }),
                  p(r(`#${idx+1}`, { color:s.color, size:14 }), { align:AlignmentType.CENTER, after:0 }),
                ],
              }),
              new TableCell({ borders:{top:{style:BorderStyle.SINGLE,size:4,color:s.fill},bottom:BD_NONE.bottom,left:BD_NONE.left,right:{style:BorderStyle.SINGLE,size:4,color:s.fill}},
                margins:{top:100,bottom:40,left:200,right:160},
                children: [p(r(issue.title, { bold:true, size:21 }), { after:0 })],
              }),
            ]}),
            new TableRow({ children: [
              new TableCell({ borders:{top:BD_NONE.top,bottom:BD_NONE.bottom,left:BD_NONE.left,right:{style:BorderStyle.SINGLE,size:4,color:s.fill}},
                margins:{top:40,bottom:40,left:200,right:160},
                children: [p(r(issue.detail, { size:18, color:C.gray700 }), { after:0, line:320 })],
              }),
            ]}),
            new TableRow({ children: [
              new TableCell({ borders:{top:BD_NONE.top,bottom:{style:BorderStyle.SINGLE,size:4,color:s.fill},left:BD_NONE.left,right:{style:BorderStyle.SINGLE,size:4,color:s.fill}},
                shading:{fill:C.gray100,type:ShadingType.CLEAR},
                margins:{top:80,bottom:100,left:200,right:160},
                children: [p([r('▶ アクション：', { bold:true, color:C.accent, size:17 }), r(issue.action, { size:17 })], { after:0 })],
              }),
            ]}),
          ],
        }),
        sp(120),
      ];
    }).flat(),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ────────────────────────────────────────
// 08 来週の投稿方針
// ────────────────────────────────────────
function buildNextWeekPlan(account, weekData) {
  const nextMonday = new Date(weekData.sunday);
  nextMonday.setDate(nextMonday.getDate() + 1);
  const nextSunday = new Date(nextMonday);
  nextSunday.setDate(nextMonday.getDate() + 6);
  const fmt = d => `${d.getMonth()+1}/${d.getDate()}`;
  const periodLabel = `${fmt(nextMonday)}（月）〜${fmt(nextSunday)}（日）`;

  return [
    h1('07', `来週の運用方針（${periodLabel}）`),
    sp(80),
    lbox([
      `投稿目標：${account.weekly_post_count || 21}件（${(account.post_times || ['08:00','10:00','20:00']).join(' / ')}）`,
      `KPI目標：${account.kpi_label} ${account.kpi_target}件以上`,
      `重点テーマ：コンテンツ4本柱のバランスを維持しながら、インサイト祭り週3本を徹底する`,
    ], C.primary, C.light),
    sp(120),
    h2('コンテンツ4本柱の配分目安'),
    sp(60),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [3000, 1200, 1200, 3960],
      rows: [
        new TableRow({ children: [th('コンテンツ種別',3000), th('目安本数',1200), th('目安時間帯',1200), th('ポイント',3960)] }),
        ...(account.content_pillars || []).map((pillar, i) => {
          const info = [
            { count:'週6本', time:'昼10時中心' },
            { count:'週6本', time:'朝8時中心' },
            { count:'週3本', time:'夜20時中心' },
            { count:'週3本', time:'任意' },
          ][i] || { count:'週3本', time:'任意' };
          return new TableRow({ children: [
            td(pillar, 3000, { fill: C.gray100 }),
            td(info.count, 1200, { align: AlignmentType.CENTER }),
            td(info.time,  1200, { align: AlignmentType.CENTER }),
            td('—', 3960, { color: C.gray400 }),
          ]});
        }),
      ],
    }),
    sp(120),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({ children: [new TableCell({
        borders:{top:{style:BorderStyle.SINGLE,size:4,color:C.gray200},bottom:{style:BorderStyle.SINGLE,size:4,color:C.gray200},left:{style:BorderStyle.SINGLE,size:20,color:C.orange},right:{style:BorderStyle.SINGLE,size:4,color:C.gray200}},
        shading:{fill:'FFFBF0',type:ShadingType.CLEAR},
        margins:{top:140,bottom:140,left:240,right:240},
        children:[
          p(r('⚠️ 運用者へのアクション確認事項', { bold:true, color:C.orange, size:20 }), { after:60 }),
          p(r('① 今週の投稿内容をレポートで確認し、コメント誘発の仕掛けが入っているか確認してください。', { size:18 }), { after:40, line:320 }),
          p(r('② インサイト祭り投稿は毎回違う文章で。コピペ禁止。', { size:18 }), { after:40, line:320 }),
          p(r('③ 毎週月曜朝にフォロワー数をスクリーンショットで共有してください。', { size:18 }), { after:0, line:320 }),
        ],
      })]})],
    }),
    sp(200), hr(C.gray200), sp(80),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [
      r(`次回レポート：${nextSunday.getFullYear()}年${nextSunday.getMonth()+1}月${nextSunday.getDate()+1}日（月）予定`, { bold:true, color:C.primary, size:18 }),
    ]}),
  ];
}

// ────────────────────────────────────────
// メイン関数
// ────────────────────────────────────────
function generateWeeklyReport(account, weekData, outputPath) {
  const hf = makeHeaderFooter(account, weekData.periodLabel);
  const doc = new Document({
    styles: { default: { document: { run: { font: FONT, size: 20, color: C.black } } } },
    sections: [{
      properties: { ...PAGE_PROPS },
      ...hf,
      children: [
        ...buildCover(account, weekData),
        ...buildSummary(account, weekData),
        ...buildKPI(account, weekData),
        ...buildPostList(weekData),
        ...buildBuzz(weekData),
        ...buildContentAnalysis(weekData),
        ...buildIssues(account, weekData),
        ...buildNextWeekPlan(account, weekData),
      ],
    }],
  });

  return Packer.toBuffer(doc).then(buf => {
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, buf);
    console.log(`  ✅ 週次レポート: ${outputPath}`);
  });
}

function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}年${d.getMonth()+1}月${d.getDate()}日`;
}

module.exports = { generateWeeklyReport };
