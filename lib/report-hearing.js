/**
 * report-hearing.js
 * 分析結果 → ヒアリングシートDOCX
 * ・セクションA〜F：全アカウント共通質問
 * ・セクションG：データから自動生成される追加質問
 */
const { Document, Packer, PageBreak, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign } = require('docx');
const fs   = require('fs');
const path = require('path');
const {
  C, FONT, BD_NONE, BD_GRAY,
  r, p, sp, hr, h1, h2,
  makeHeaderFooter, PAGE_PROPS,
} = require('./docx-base');

// ─────────────────────────────────────────────
// スタイル定数
// ─────────────────────────────────────────────
const SECTION_COLORS = {
  'A': '1E3A5F', 'B': '2E86AB', 'C': '27AE60',
  'D': 'E67E22', 'E': '8E44AD', 'F': '16A085', 'G': 'C0392B',
};

// グレー質問行 + 白回答行のQAブロック
function qaBlock(question, lines = 2, opts = {}) {
  const important = opts.important || false;
  const hint      = opts.hint || '';
  const why       = opts.why  || '';
  const qColor    = important ? 'FEF2F2' : C.gray100;
  const qBorderColor = important ? C.red : C.gray200;
  const minHeight = lines * 400;
  const rows = [];

  // 質問行
  rows.push(new TableRow({ children: [new TableCell({
    borders: {
      top:   { style: BorderStyle.SINGLE, size: 2, color: qBorderColor },
      bottom: BD_NONE.bottom,
      left:  { style: BorderStyle.SINGLE, size: important ? 8 : 2, color: important ? C.red : C.gray400 },
      right: { style: BorderStyle.SINGLE, size: 2, color: qBorderColor },
    },
    shading: { fill: qColor, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 60, left: 160, right: 160 },
    children: [
      p([
        ...(important ? [r('★ ', { bold: true, color: C.red, size: 18 })] : []),
        r(question, { bold: true, size: 18, color: C.black }),
      ], { after: why || hint ? 40 : 0 }),
      ...(why  ? [p(r(`💡 なぜ聞くか：${why}`,  { size: 15, color: C.accent }),  { after: hint ? 20 : 0 })] : []),
      ...(hint ? [p(r(`📝 記入例：${hint}`,      { size: 15, color: C.gray400 }), { after: 0 })] : []),
    ],
  })]})
  );

  // 回答行（最小高さあり）
  rows.push(new TableRow({
    height: { value: minHeight, rule: 'atLeast' },
    children: [new TableCell({
      borders: {
        top:    BD_NONE.top,
        bottom: { style: BorderStyle.SINGLE, size: 2, color: C.gray200 },
        left:   { style: BorderStyle.SINGLE, size: important ? 8 : 2, color: important ? C.red : C.gray400 },
        right:  { style: BorderStyle.SINGLE, size: 2, color: C.gray200 },
      },
      margins: { top: 80, bottom: 80, left: 160, right: 160 },
      children: [p(r('', { size: 18 }), { after: 0 })],
    })],
  }));

  return rows;
}

// セクション見出し
function sectionHeader(label, title, color) {
  return [
    sp(160),
    new Paragraph({
      spacing: { before: 0, after: 120 },
      border: { left: { style: BorderStyle.SINGLE, size: 24, color } },
      children: [
        r(`  ${label}`, { bold: true, color, size: 22 }),
        r(`  ${title}`, { bold: true, color: C.black, size: 22 }),
      ],
    }),
  ];
}

// KPIサマリーボックス（表紙用）
function kpiSummaryBox(analysisResult) {
  const { kpi, issues, posts, account } = analysisResult;
  const issueCount = { CRITICAL: 0, HIGH: 0, MED: 0, LOW: 0 };
  issues.forEach(i => issueCount[i.severity]++);

  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [2340, 2340, 2340, 2340],
    rows: [
      new TableRow({ children: [
        dataCard('総投稿数', String(kpi.total), '件', C.primary, 2340),
        dataCard('いいね合計', kpi.likes.toLocaleString(), '', C.accent, 2340),
        dataCard('0いいね率', String(kpi.zeroLikeRate) + '%', '', kpi.zeroLikeRate > 70 ? C.red : C.orange, 2340),
        dataCard('検出課題数', `H:${issueCount.HIGH} M:${issueCount.MED}`, '', C.red, 2340),
      ]}),
    ],
  });
}

function dataCard(label, value, unit, color, w) {
  return new TableCell({
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color },
      bottom: BD_GRAY.bottom, left: BD_GRAY.left, right: BD_GRAY.right,
    },
    width: { size: w, type: WidthType.DXA },
    shading: { fill: C.gray100, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 160, right: 160 },
    children: [
      p(r(label, { size: 16, color: C.gray400 }), { after: 40 }),
      p([r(value, { bold: true, size: 36, color }), r(' ' + unit, { size: 18, color: C.gray700 })], { after: 0 }),
    ],
  });
}

// ─────────────────────────────────────────────
// ヒアリングシート本体
// ─────────────────────────────────────────────
function buildHearing(analysisResult) {
  const { account, kpi, issues, questions, missingInfo } = analysisResult;
  const content = [];

  // ── 表紙 ─────────────────────────────────────
  content.push(
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({ children: [new TableCell({
        borders: BD_NONE, width: { size: 9360, type: WidthType.DXA },
        shading: { fill: C.primary, type: ShadingType.CLEAR },
        margins: { top: 240, bottom: 240, left: 360, right: 360 },
        children: [
          p(r('MTG ヒアリングシート', { bold: true, color: C.white, size: 32 }), { after: 80 }),
          p(r(`${account.handle}｜${account.business_type || '業種未設定'}`, { color: 'B0C4DE', size: 19 }), { after: 0 }),
        ],
      })]})],
    }),
    sp(160),
  );

  // データサマリー（スクレイピングデータがある場合）
  if (analysisResult.hasData) {
    content.push(
      h2('スクレイピングデータ概要（事前分析）'),
      sp(60),
      kpiSummaryBox(analysisResult),
      sp(100),
    );
  }

  // 検出された課題サマリー
  if (issues.length > 0) {
    content.push(
      h2('データから検出された課題'),
      sp(60),
      new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [800, 2200, 6360],
        rows: [
          new TableRow({ children: [
            new TableCell({ borders: BD_NONE, width: { size: 800, type: WidthType.DXA },
              shading: { fill: C.primary, type: ShadingType.CLEAR }, margins: { top:80, bottom:80, left:100, right:100 },
              children: [p(r('重要度', { bold:true, color:C.white, size:16 }), { align:AlignmentType.CENTER, after:0 })] }),
            new TableCell({ borders: BD_NONE, width: { size: 2200, type: WidthType.DXA },
              shading: { fill: C.primary, type: ShadingType.CLEAR }, margins: { top:80, bottom:80, left:100, right:100 },
              children: [p(r('課題', { bold:true, color:C.white, size:16 }), { after:0 })] }),
            new TableCell({ borders: BD_NONE, width: { size: 6360, type: WidthType.DXA },
              shading: { fill: C.primary, type: ShadingType.CLEAR }, margins: { top:80, bottom:80, left:100, right:100 },
              children: [p(r('詳細', { bold:true, color:C.white, size:16 }), { after:0 })] }),
          ]}),
          ...issues.slice(0, 8).map(issue => {
            const sc = { CRITICAL:'FDDCDC', HIGH:'FDDCDC', MED:'FFF3CD', LOW:'D4EDDA' };
            const tc = { CRITICAL:C.red, HIGH:C.red, MED:C.orange, LOW:C.green };
            return new TableRow({ children: [
              new TableCell({ borders: BD_GRAY, width:{size:800,type:WidthType.DXA},
                shading:{fill:sc[issue.severity],type:ShadingType.CLEAR},
                margins:{top:60,bottom:60,left:80,right:80},
                children:[p(r(issue.severity,{bold:true,size:15,color:tc[issue.severity]}),{align:AlignmentType.CENTER,after:0})] }),
              new TableCell({ borders: BD_GRAY, width:{size:2200,type:WidthType.DXA},
                margins:{top:60,bottom:60,left:120,right:120},
                children:[p(r(issue.title,{size:16,bold:true}),{after:0,line:300})] }),
              new TableCell({ borders: BD_GRAY, width:{size:6360,type:WidthType.DXA},
                margins:{top:60,bottom:60,left:120,right:120},
                children:[p(r(issue.detail,{size:16,color:C.gray700}),{after:0,line:300})] }),
            ]});
          }),
        ],
      }),
      sp(80),
    );
  }

  content.push(new Paragraph({ children: [new PageBreak()] }));

  // ── セクション別ヒアリング質問 ────────────────
  const sections = ['A', 'B', 'C', 'D', 'E', 'F', 'G'];
  const sectionTitles = {
    'A': 'ビジネス基本情報',
    'B': 'Threadsの目標設定',
    'C': 'ターゲット顧客',
    'D': '実績・強み・NG情報',
    'E': '投稿スタイル・キャラクター',
    'F': '運用体制・ワークフロー',
    'G': 'データから見えた課題への質問',
  };

  sections.forEach(sec => {
    const secQuestions = questions.filter(q => q.section.startsWith(sec));
    if (secQuestions.length === 0) return;

    const color = SECTION_COLORS[sec] || C.primary;
    content.push(...sectionHeader(sec, sectionTitles[sec], color));

    // テーブルでまとめる
    const tableRows = [];
    secQuestions.forEach(q => {
      const important = q.priority === 'HIGH' || q.priority === 'CRITICAL';
      tableRows.push(...qaBlock(q.question, important ? 3 : 2, {
        important,
        why:  q.why,
        hint: q.hint,
      }));
    });

    content.push(
      new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
        rows: tableRows,
      }),
      sp(40),
    );
  });

  // ── フリーメモ欄 ──────────────────────────────
  content.push(
    sp(120),
    h2('その他・メモ（自由記入）'),
    sp(60),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({
        height: { value: 2000, rule: 'atLeast' },
        children: [new TableCell({
          borders: BD_GRAY,
          margins: { top: 80, bottom: 80, left: 160, right: 160 },
          children: [p(r('', { size: 18 }), { after: 0 })],
        })],
      })],
    }),
    sp(200),
    hr(C.gray200),
    sp(80),
    p(r(`作成日：${todayStr()}　｜　担当ディレクター：${account.director || '未設定'}`, { size: 16, color: C.gray400 }), { align: AlignmentType.CENTER, after: 0 }),
  );

  return content;
}

// ─────────────────────────────────────────────
// メイン関数
// ─────────────────────────────────────────────
function generateHearingSheet(analysisResult, outputPath) {
  const { account, questions } = analysisResult;
  const periodLabel = `MTGヒアリング ${todayStr()}`;
  const hf = makeHeaderFooter(account, periodLabel);

  const doc = new Document({
    styles: { default: { document: { run: { font: FONT, size: 20, color: C.black } } } },
    sections: [{
      properties: { ...PAGE_PROPS },
      ...hf,
      children: buildHearing(analysisResult),
    }],
  });

  return Packer.toBuffer(doc).then(buf => {
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, buf);
    console.log(`  ✅ ヒアリングシート: ${outputPath}`);
  });
}

function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}年${d.getMonth()+1}月${d.getDate()}日`;
}

module.exports = { generateHearingSheet };
