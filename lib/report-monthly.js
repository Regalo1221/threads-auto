/**
 * report-monthly.js
 * 7セクション月次レポート（build_monthly_report.js と同じ構成）
 *
 * 表紙 → 01サマリー → 02KPI実績 → 03週別推移 → 04TOP3 → 05コンテンツ分析 → 06課題 → 07方針
 */
const {
  Document, Packer, PageBreak, Paragraph, TextRun,
  Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
} = require('docx');
const fs   = require('fs');
const path = require('path');
const {
  C, FONT, BD_NONE, BD_GRAY,
  r, p, sp, hr, h1, h2, th, td,
  kpiRow, lbox, makeHeaderFooter, PAGE_PROPS,
} = require('./docx-base');
const { analyzeAccount } = require('./analyzer');
const { getMonthData }   = require('./data-loader');

// ─────────────────────────────────────────────
// ユーティリティ
// ─────────────────────────────────────────────
function pageBreak() { return new Paragraph({ children: [new PageBreak()] }); }

function deltaStr(curr, prev) {
  if (prev == null || prev === 0) return '—';
  const d = curr - prev;
  return (d >= 0 ? '▲' : '▼') + Math.abs(d).toLocaleString();
}

function rateStr(curr, target) {
  if (!target || target === 0) return '—';
  const r = Math.round((curr / target) * 100);
  if (r >= 100) return `${r}% ✅`;
  if (r >= 70)  return `${r}% △`;
  return `${r}% ▼未達`;
}

function rateColor(curr, target) {
  if (!target || target === 0) return C.gray400;
  const r = (curr / target) * 100;
  if (r >= 100) return C.green;
  if (r >= 70)  return C.orange;
  return C.red;
}

// 投稿種別ごとの「意図・結果・改善」コメント
function evalComment(type, d, total) {
  const pct = total > 0 ? Math.round((d.count / total) * 100) : 0;

  if (type === '集客・リーチ拡大') {
    const intent = '【意図】新規フォロワーへのリーチ拡大・認知獲得';
    if (pct < 10) return intent + ' / 【結果】比率が低く（' + pct + '%）、拡散機会が限定的 / 【改善】問いかけ型・拡散型の投稿を増量';
    if (d.avgLikes >= 5) return intent + ' / 【結果】高反応（平均' + d.avgLikes + 'いいね）でリーチ拡大に貢献 / 【継続】引き続きバリエーションを増やす';
    return intent + ' / 【結果】標準的な反応 / 【改善】フック文の強化で拡散力を高める';
  }
  if (type === '専門性・価値提供') {
    const intent = '【意図】専門性・信頼構築・フォロワーの教育';
    if (d.avgLikes < 1) return intent + ' / 【結果】反応が薄い（平均' + d.avgLikes + 'いいね）/ 【改善】具体的な数字・事例を加えて訴求力を高める';
    if (d.avgLikes >= 5) return intent + ' / 【結果】高反応（平均' + d.avgLikes + 'いいね）で信頼構築に貢献 / 【継続】成功パターンを横展開';
    return intent + ' / 【結果】安定した反応 / 【改善】ビフォーアフターや数字を入れて説得力アップ';
  }
  if (type === '共感・エンゲージ') {
    const intent = '【意図】フォロワーとの関係性構築・コメント誘発';
    if (d.avgLikes >= 10) return intent + ' / 【結果】高反応（平均' + d.avgLikes + 'いいね）でエンゲージ向上に貢献 / 【継続】ターゲットに刺さるあるあるを継続';
    if (d.avgLikes < 2) return intent + ' / 【結果】反応が限定的 / 【改善】ターゲット層に特化した共感ネタに絞る';
    return intent + ' / 【結果】エンゲージメント確保に貢献 / 【改善】コメントを促す問いかけを文末に追加';
  }
  if (type === '日常・雑談') {
    const intent = '【意図】人柄の露出・親近感の醸成';
    if (pct > 60) return intent + ' / 【結果】比率が高め（' + pct + '%）でキャラクターは伝わりやすい状態 / 【調整】専門性投稿とのバランスを整える';
    return intent + ' / 【結果】適切な比率でキャラクター露出に機能 / 【継続】ターゲットに共感されやすいネタを選定';
  }
  // その他
  if (d.avgLikes < 1) return '【結果】反応が薄い / 【改善】切り口・訴求の見直しを推奨';
  if (d.avgLikes >= 5) return '【結果】高反応 / 【継続】成功パターンを横展開';
  return '【結果】標準的な反応 / 【改善】具体性・エピソードを加えると改善見込み';
}

// ─────────────────────────────────────────────
// 表紙
// ─────────────────────────────────────────────
function buildCover(account, monthData) {
  return [
    p(r('Threads 運用 月次レポート', { bold: true, color: C.primary, size: 52 }), { after: 40 }),
    p(r(account.handle, { bold: true, color: C.accent, size: 28 }), { after: 40 }),
    new Paragraph({
      spacing: { before: 0, after: 200 },
      children: [
        r(`報告期間：${monthData.periodLabel}`, { size: 20, color: C.gray700 }),
        r(`　　作成日：${todayStr()}`, { size: 18, color: C.gray400 }),
      ],
    }),
    hr(C.accent, 6), sp(120),
    lbox([
      `本レポートは ${account.handle}（${account.business_type || '業種未設定'}）の月次運用実績報告です。`,
      '「当月の投稿実績」「課題分析」「翌月の運用方針」の3部構成です。',
      '翌月頭にお届けします。内容に関するご質問は担当ディレクターまでお気軽にどうぞ。',
    ], C.accent, C.light, C.gray700),
    sp(200),
  ];
}

// ─────────────────────────────────────────────
// 01 エグゼクティブサマリー
// ─────────────────────────────────────────────
function buildSummary(account, monthData, issues) {
  const { kpi } = monthData;

  // サマリーテキスト（事実ベース＋代理店目線の意図説明）
  const criticalIssues = issues.filter(i => i.severity === 'HIGH' || i.severity === 'CRITICAL');
  let summaryText = '【当月の運用概要】';
  summaryText += `総投稿数${kpi.total}件、いいね合計${kpi.likes.toLocaleString()}件、コメント合計${kpi.comments}件。`;
  if (criticalIssues.length > 0) {
    summaryText += `データ分析の結果、${criticalIssues.map(i => i.title).join('・')}が改善ポイントとして挙がりました。`;
    summaryText += '各投稿は下記に記載の意図のもと実施しており、数値を踏まえて翌月の方針を調整します。';
  } else {
    summaryText += '全体的に安定した運用を継続できています。引き続き数値改善に向けた取り組みを続けます。';
  }

  const blocks = [
    h1('01', 'エグゼクティブサマリー'),
    p(summaryText, { color: C.gray700, line: 320, after: 140, size: 19 }),
    kpiRow([
      { label: '総投稿数',     value: String(kpi.total),            unit: '件', sub: '当月合計' },
      { label: 'いいね合計',   value: kpi.likes.toLocaleString(),    unit: '',   sub: '当月合計' },
      { label: 'コメント合計', value: String(kpi.comments),          unit: '',   sub: '当月合計' },
      { label: account.kpi_label || 'KPI目標', value: '—', unit: '件', sub: `目標：${account.kpi_target || '—'}件`, subColor: C.red },
    ]),
    sp(100),
  ];

  // 注目ポイントボックス（HIGH課題がある場合）
  if (criticalIssues.length > 0) {
    const warnLines = ['📌 翌月に向けた重点改善ポイント:'];
    criticalIssues.forEach(i => warnLines.push(`　・${i.title}`));
    blocks.push(
      lbox(warnLines, C.orange, C.yellow, C.gray700),
      sp(80),
    );
  }

  blocks.push(pageBreak());
  return blocks;
}

// ─────────────────────────────────────────────
// 02 月別KPI実績
// ─────────────────────────────────────────────
function buildKpiTable(account, monthData, prevKpi) {
  const { kpi } = monthData;
  const target   = account.kpi_target || 0;
  const weekTarget = account.weekly_post_count || 0;
  const monthTarget = weekTarget * 4;

  const rows = [
    ['投稿数（件）',         kpi.total,        prevKpi ? prevKpi.total    : null, monthTarget, rateStr(kpi.total, monthTarget)],
    ['いいね合計',           kpi.likes,        prevKpi ? prevKpi.likes    : null, null,        null],
    ['いいね平均/投稿',      kpi.total > 0 ? +(kpi.likes/kpi.total).toFixed(1) : 0,
                             prevKpi && prevKpi.total > 0 ? +(prevKpi.likes/prevKpi.total).toFixed(1) : null, null, null],
    ['コメント合計',         kpi.comments,     prevKpi ? prevKpi.comments : null, null,        null],
    ['コメント平均/投稿',    kpi.total > 0 ? +(kpi.comments/kpi.total).toFixed(1) : 0,
                             prevKpi && prevKpi.total > 0 ? +(prevKpi.comments/prevKpi.total).toFixed(1) : null, null, null],
    ['リポスト合計',         kpi.reposts,      prevKpi ? prevKpi.reposts  : null, null,        null],
    ['エンゲージメントスコア', kpi.score.toLocaleString(), prevKpi ? prevKpi.score.toLocaleString() : null, null, null],
    [account.kpi_label || 'KPI（面談等）', '—', '—', String(target) + '件', rateStr(0, target)],
  ];

  const cols = [2400, 1560, 1560, 1560, 1680, 0];
  const totalW = 9360;
  // 5列目残り
  const c5 = totalW - 2400 - 1560 - 1560 - 1560;
  const colWidths = [2400, 1560, 1560, 1560, c5];

  const { year, month } = monthData;
  const prevY = month === 1 ? year - 1 : year;
  const prevM = month === 1 ? 12 : month - 1;

  return [
    h1('02', '月別 KPI実績'),
    sp(80),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: colWidths,
      rows: [
        new TableRow({ children: [
          th('指標', colWidths[0], AlignmentType.LEFT),
          th(`${year}年${month}月（当月）`, colWidths[1]),
          th(`${prevY}年${prevM}月（前月）`, colWidths[2]),
          th('目標', colWidths[3]),
          th('達成状況', colWidths[4]),
        ]}),
        ...rows.map(([label, curr, prev, tgt, rate]) => {
          const currStr = curr != null ? String(curr) : '—';
          const prevStr = prev != null ? String(prev) : '—';
          const tgtStr  = tgt  != null ? String(tgt)  : '—';
          const rateStr2 = rate || '—';
          const rc = rateStr2.includes('未達') ? C.red : rateStr2.includes('△') ? C.orange : C.black;
          return new TableRow({ children: [
            td(label, colWidths[0], { fill: C.gray100, size: 17 }),
            td(currStr, colWidths[1], { align: AlignmentType.CENTER, size: 17, bold: true, color: C.primary }),
            td(prevStr, colWidths[2], { align: AlignmentType.CENTER, size: 17, color: C.gray400 }),
            td(tgtStr,  colWidths[3], { align: AlignmentType.CENTER, size: 16, color: C.accent }),
            td(rateStr2, colWidths[4], { align: AlignmentType.CENTER, size: 16, bold: rateStr2.includes('未達'), color: rc }),
          ]});
        }),
      ],
    }),
    sp(80),
    p([r('※ ', { size: 16, color: C.gray400 }), r('エンゲージメントスコア = いいね×1 + コメント×3 + リポスト×5', { size: 16, color: C.gray400 })], { after: 160 }),
  ];
}

// ─────────────────────────────────────────────
// 03 週別エンゲージメント推移
// ─────────────────────────────────────────────
function buildWeeklyTrend(monthData) {
  const { weeklyData } = monthData;
  const cols = [1680, 1080, 1680, 1680, 1680, 1560];

  // 推移コメント自動生成
  const maxScore = weeklyData.reduce((m, w) => Math.max(m, w.score), 0);
  const minScore = weeklyData.reduce((m, w) => Math.min(m, w.score), Infinity);
  const ratio = minScore > 0 ? Math.round(maxScore / minScore) : 999;
  const warnLines = ['【推移から読み取れること】'];
  if (ratio > 10) {
    warnLines.push(`週間スコアの最大/最小比が ${ratio}倍と大きく、バズ依存の不安定な構造が見られます。`);
    warnLines.push('安定した面談獲得には継続的なコンテンツ設計が必要です。');
  } else {
    warnLines.push('週別スコアのばらつきは比較的小さく、安定した投稿ペースが保てています。');
  }

  return [
    h1('03', '週別エンゲージメント推移'),
    sp(80),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: cols,
      rows: [
        new TableRow({ children: [
          th('週', cols[0]), th('投稿数', cols[1]), th('いいね', cols[2]),
          th('コメント', cols[3]), th('スコア', cols[4]), th('主な特徴', cols[5]),
        ]}),
        ...weeklyData.map(w => {
          const bigLike  = w.likes  >= 1000;
          const bigScore = w.score  >= 1000;
          return new TableRow({ children: [
            td(w.week, cols[0], { fill: C.gray100, size: 16, bold: true }),
            td(String(w.total), cols[1], { align: AlignmentType.CENTER, size: 17 }),
            td(w.likes.toLocaleString(), cols[2], { align: AlignmentType.CENTER, size: 17, bold: bigLike, color: bigLike ? C.primary : C.black }),
            td(String(w.comments), cols[3], { align: AlignmentType.CENTER, size: 17 }),
            td(w.score.toLocaleString() + 'pt', cols[4], { align: AlignmentType.CENTER, size: 17, bold: bigScore, color: bigScore ? C.primary : C.black }),
            td(w.note || '—', cols[5], { size: 15, color: C.gray700 }),
          ]});
        }),
      ],
    }),
    sp(80),
    lbox(warnLines, C.orange, C.yellow, C.gray700),
    sp(160),
  ];
}

// ─────────────────────────────────────────────
// 04 バズ投稿 TOP3
// ─────────────────────────────────────────────
function buildBuzzTop3(monthData) {
  const top3 = (monthData.kpi.top3 || []).slice(0, 3);
  const rankColors = [C.orange, '888888', '8B5E1A'];
  const medals = ['🥇', '🥈', '🥉'];

  const cards = top3.map((post, idx) => {
    const bc = rankColors[idx];
    const lines = (post.content || '').split('\n').filter(l => l.trim()).slice(0, 6);
    return [
      new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
        rows: [
          // ヘッダー行
          new TableRow({ children: [new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 6, color: bc }, bottom: BD_NONE.bottom, left: { style: BorderStyle.SINGLE, size: 6, color: bc }, right: { style: BorderStyle.SINGLE, size: 6, color: bc } },
            shading: { fill: C.gray100, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 200, right: 200 },
            children: [p([
              r(`${medals[idx]} 第${idx + 1}位`, { bold: true, color: bc, size: 22 }),
              r(`　${post.dateStr || ''}　　`, { size: 18, color: C.gray700 }),
              r(`❤️ ${(post.likes||0).toLocaleString()}　💬 ${post.comments||0}　🔁 ${post.reposts||0}　🏆 ${(post.score||0).toLocaleString()}pt`, { bold: true, color: C.primary, size: 18 }),
            ], { after: 0 })],
          })]})
          ,
          // 投稿内容
          new TableRow({ children: [new TableCell({
            borders: { top: BD_NONE.top, bottom: { style: BorderStyle.SINGLE, size: 2, color: C.gray200 }, left: { style: BorderStyle.SINGLE, size: 6, color: bc }, right: { style: BorderStyle.SINGLE, size: 6, color: bc } },
            margins: { top: 60, bottom: 80, left: 200, right: 200 },
            children: lines.length > 0
              ? lines.map((line, i) => p(r(line, { size: 18 }), { after: i < lines.length - 1 ? 40 : 0, line: 300 }))
              : [p(r('（投稿内容なし）', { size: 18, color: C.gray400 }), { after: 0 })],
          })]})
          ,
          // 分析行
          new TableRow({ children: [new TableCell({
            borders: { top: BD_NONE.top, bottom: { style: BorderStyle.SINGLE, size: 6, color: bc }, left: { style: BorderStyle.SINGLE, size: 6, color: bc }, right: { style: BorderStyle.SINGLE, size: 6, color: bc } },
            shading: { fill: 'FEF2F2', type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 100, left: 200, right: 200 },
            children: [
              p([
                r('種別：', { bold: true, size: 17, color: C.gray700 }),
                r(post.type || '—', { size: 17 }),
              ], { after: 40 }),
              p([
                r('📌 投稿意図：', { bold: true, size: 17, color: C.accent }),
                r('（この投稿をどういう目的で作成・投稿したかを記入）', { size: 17, color: C.gray400 }),
              ], { after: 30, line: 320 }),
              p([
                r('📋 数値考察：', { bold: true, size: 17, color: C.primary }),
                r('（反応の理由・学び・次回への活かし方を記入）', { size: 17, color: C.gray400 }),
              ], { after: 0, line: 320 }),
            ],
          })]})
        ],
      }),
      sp(100),
    ];
  });

  return [
    pageBreak(),
    h1('04', 'バズ投稿 TOP 3 詳細'),
    p('エンゲージメントスコア上位3投稿の内容と、ゴールとの関連性を分析します。', { color: C.gray700, size: 18, after: 120 }),
    ...cards.flat(),
    sp(60),
  ];
}

// ─────────────────────────────────────────────
// 05 コンテンツ種別分析
// ─────────────────────────────────────────────
function buildContentAnalysis(monthData) {
  const { kpi } = monthData;
  const types  = Object.entries(kpi.byType || {}).sort((a, b) => b[1].count - a[1].count);
  const total  = types.reduce((s, [, d]) => s + d.count, 0);
  const typeColors = {
    '集客・リーチ拡大':   'E2D9F3',
    '専門性・価値提供':   'FDDCDC',
    '共感・エンゲージ':   'D1ECF1',
    '日常・雑談':         'FFF3CD',
  };
  const cols = [2200, 800, 1160, 1160, 1160, 2880];

  // コメントボックス自動生成
  const highType = types.length > 0 ? types[0][0] : '';
  const highPct  = total > 0 ? Math.round((types[0]?.[1]?.count || 0) / total * 100) : 0;
  const commentLines = [];
  if (highType === '雑談・日常' && highPct > 50) {
    commentLines.push(`【重要な逆説】いいね平均が最も高い「${highType}」（${types[0][1].avgLikes}/投稿）は面談獲得と無関係。`);
    commentLines.push('　　　　　　　数字が良く見える投稿ほど、ゴールから遠ざかっている構造を理解した上で方針を転換します。');
  } else {
    commentLines.push('コンテンツの種別比率を確認し、ゴールに直結する種別の比率を最適化してください。');
  }

  return [
    pageBreak(),
    h1('05', 'コンテンツ種別 パフォーマンス分析'),
    sp(80),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: cols,
      rows: [
        new TableRow({ children: [
          th('種別', cols[0]), th('件数', cols[1]), th('いいね合計', cols[2]),
          th('いいね平均', cols[3]), th('スコア平均', cols[4]), th('評価', cols[5]),
        ]}),
        ...types.map(([type, d]) => new TableRow({ children: [
          td(type, cols[0], { fill: typeColors[type] || C.gray100, size: 17 }),
          td(String(d.count), cols[1], { align: AlignmentType.CENTER, size: 17 }),
          td((d.likes || 0).toLocaleString(), cols[2], { align: AlignmentType.CENTER, size: 17 }),
          td(String(d.avgLikes), cols[3], { align: AlignmentType.CENTER, size: 17, bold: d.avgLikes >= 2, color: d.avgLikes >= 5 ? C.green : d.avgLikes >= 2 ? C.accent : C.black }),
          td(String(d.avgScore), cols[4], { align: AlignmentType.CENTER, size: 17, bold: d.avgScore >= 2, color: d.avgScore >= 5 ? C.green : d.avgScore >= 2 ? C.accent : C.black }),
          td(evalComment(type, d, total), cols[5], { size: 15, color: C.gray700 }),
        ]})),
      ],
    }),
    sp(80),
    lbox(commentLines, C.orange, C.yellow, C.gray700),
    sp(160),
  ];
}

// ─────────────────────────────────────────────
// 06 課題の総括
// ─────────────────────────────────────────────
function buildIssues(issues) {
  const sortOrder = { CRITICAL: 0, HIGH: 1, MED: 2, LOW: 3 };
  const sorted = [...issues].sort((a, b) => (sortOrder[a.severity] || 9) - (sortOrder[b.severity] || 9));

  const lvlColor = { CRITICAL: C.red, HIGH: C.red, MED: C.orange, LOW: C.gray400 };
  const lvlFill  = { CRITICAL: 'FEF2F2', HIGH: 'FEF2F2', MED: C.yellow, LOW: C.gray100 };

  return [
    pageBreak(),
    h1('06', '改善ポイント・翌月への学び'),
    sp(80),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [600, 1800, 6960],
      rows: [
        new TableRow({ children: [th('優先度', 600), th('課題', 1800), th('詳細', 6960)] }),
        ...sorted.map(issue => new TableRow({ children: [
          td(issue.severity, 600, { fill: lvlFill[issue.severity] || C.gray100, bold: true, color: lvlColor[issue.severity] || C.gray400, align: AlignmentType.CENTER, size: 16 }),
          td(issue.title, 1800, { bold: true, size: 17 }),
          td(issue.detail, 6960, { size: 17, color: C.gray700 }),
        ]})),
      ],
    }),
    sp(160),
  ];
}

// ─────────────────────────────────────────────
// 07 今後の運用方針
// ─────────────────────────────────────────────
function buildStrategy(account, monthData) {
  const { month } = monthData;
  const nextMonth = month === 12 ? 1 : month + 1;
  const nextYear  = month === 12 ? monthData.year + 1 : monthData.year;
  const pillars   = account.content_pillars || ['（記入）', '（記入）', '（記入）', '（記入）'];

  // ポジション定義文
  const positionLines = [
    `「${account.target_audience || 'ターゲット層未設定'}に向けて、${account.goal || '目標未設定'}を達成するアカウント」`,
    '',
    `話しかける相手：${account.target_audience || '（未設定）'}`,
    `最終的な誘導先：公式LINE → ${account.kpi_label || 'KPI目標'}`,
  ];

  return [
    h1('07', `今後の運用方針（${nextYear}年${nextMonth}月）`),
    p('上記の課題を踏まえ、翌月から以下の方針で運用を調整します。', { color: C.gray700, size: 18, after: 120 }),

    h2('アカウントのポジション再定義'),
    lbox(positionLines, C.primary, C.light, C.gray700),
    sp(100),

    h2(`コンテンツの${pillars.length}本柱（月次KPI付き）`),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [400, 2200, 1120, 1560, 4080],
      rows: [
        new TableRow({ children: [th('', 400), th('種別', 2200), th('週本数', 1120), th('月本数', 1560), th('目的・変更ポイント', 4080)] }),
        ...pillars.map((pillar, i) => new TableRow({ children: [
          td(`${['①','②','③','④','⑤','⑥'][i]}`, 400, { fill: C.light, bold: true, color: C.primary, align: AlignmentType.CENTER, size: 18 }),
          td(pillar, 2200, { bold: true, size: 17 }),
          td('—', 1120, { align: AlignmentType.CENTER, size: 17, color: C.gray400 }),
          td('—', 1560, { align: AlignmentType.CENTER, size: 17, color: C.gray400 }),
          td('（ディレクターが記入）', 4080, { size: 16, color: C.gray400 }),
        ]})),
      ],
    }),
    sp(100),

    h2(`${nextYear}年${nextMonth}月 優先アクション`),
    ...['即日', '今週中', 'MTG後', '今月末'].map(timing =>
      new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [1200, 8160],
        rows: [new TableRow({ children: [
          td(timing, 1200, { fill: C.light, bold: true, color: C.primary, align: AlignmentType.CENTER, size: 17 }),
          td('（ディレクターが記入）', 8160, { size: 18, color: C.gray400 }),
        ]})],
      })
    ).flatMap(t => [t, sp(40)]),

    sp(160), hr(C.gray200), sp(80),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [r(`次回レポート：${nextYear}年${nextMonth}月末予定　　本資料に関するご質問はお気軽にご連絡ください`, { size: 16, color: C.gray400 })],
    }),
  ];
}

// ─────────────────────────────────────────────
// メイン
// ─────────────────────────────────────────────
function generateMonthlyReport(account, monthData, outputPath) {
  // 課題検出
  const analysis = analyzeAccount(account);
  const issues   = analysis.issues;

  // 前月データ（比較用）
  let prevKpi = null;
  try {
    const prevMonth = monthData.month === 1 ? 12 : monthData.month - 1;
    const prevYear  = monthData.month === 1 ? monthData.year - 1 : monthData.year;
    const prev      = getMonthData(account, prevYear, prevMonth);
    prevKpi = prev.kpi;
  } catch (e) {
    // 前月データなし → null のまま
  }

  const hf  = makeHeaderFooter(account, monthData.periodLabel);
  const doc = new Document({
    styles: { default: { document: { run: { font: FONT, size: 19, color: C.black } } } },
    sections: [{
      properties: { ...PAGE_PROPS },
      ...hf,
      children: [
        ...buildCover(account, monthData),
        ...buildSummary(account, monthData, issues),
        ...buildKpiTable(account, monthData, prevKpi),
        ...buildWeeklyTrend(monthData),
        ...buildBuzzTop3(monthData),
        ...buildContentAnalysis(monthData),
        ...buildIssues(issues),
        ...buildStrategy(account, monthData),
      ],
    }],
  });

  return Packer.toBuffer(doc).then(buf => {
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, buf);
    console.log(`  ✅ 月次レポート: ${outputPath}`);
  });
}

function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
}

module.exports = { generateMonthlyReport };
