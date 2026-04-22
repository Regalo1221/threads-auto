/**
 * report-monthly.js
 * 任意のアカウント設定 + 月次データ → 月次レポートDOCX
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
// 01 エグゼクティブサマリー
// ────────────────────────────────────────
function buildSummary(account, monthData) {
  const { kpi, periodLabel } = monthData;
  return [
    // タイトル行
    new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
      rows:[new TableRow({children:[new TableCell({
        borders:BD_NONE, width:{size:9360,type:WidthType.DXA},
        shading:{fill:C.primary,type:ShadingType.CLEAR},
        margins:{top:200,bottom:200,left:360,right:360},
        children:[
          p(r(`${periodLabel} 月次運用レポート`, {bold:true,color:C.white,size:28}), {after:60}),
          p(r(`${account.handle}｜${account.business_type}｜目標：${account.goal}`, {color:'B0C4DE',size:18}), {after:0}),
        ],
      })]})],
    }),
    sp(160),
    kpiRow([
      {label:'総投稿数',      value:String(kpi.total),              unit:'件',  sub:'当月合計'},
      {label:'いいね合計',    value:kpi.likes.toLocaleString(),      unit:'',    sub:'当月合計'},
      {label:'コメント合計',  value:String(kpi.comments),            unit:'',    sub:'当月合計'},
      {label:account.kpi_label||'面談数', value:'—',                unit:'件',  sub:`目標：${account.kpi_target}件`},
    ]),
    sp(100),
    p(r('※ エンゲージメントスコア = いいね×1 + コメント×3 + リポスト×5', {size:16,color:C.gray400}), {after:0}),
    sp(80),
    new Paragraph({children:[new PageBreak()]}),
  ];
}

// ────────────────────────────────────────
// 02 週別KPI推移
// ────────────────────────────────────────
function buildWeeklyTrend(monthData) {
  const { weeklyData } = monthData;
  const cols = [1200,1200,1200,1200,1200,3360];

  return [
    h1('01', '週別エンゲージメント推移'),
    sp(80),
    new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:cols,
      rows:[
        new TableRow({children:[th('週',cols[0]),th('投稿数',cols[1]),th('いいね',cols[2]),th('コメ',cols[3]),th('スコア',cols[4]),th('所見',cols[5])]}),
        ...weeklyData.map(w => {
          const scoreColor = w.score > 1000 ? C.green : w.score > 100 ? C.orange : C.gray700;
          return new TableRow({children:[
            td(w.week, cols[0], {align:AlignmentType.CENTER, bold:true}),
            td(String(w.total), cols[1], {align:AlignmentType.CENTER}),
            td(w.likes.toLocaleString(), cols[2], {align:AlignmentType.CENTER}),
            td(String(w.comments), cols[3], {align:AlignmentType.CENTER}),
            td(w.score.toLocaleString()+'pt', cols[4], {align:AlignmentType.CENTER, bold:true, color:scoreColor}),
            td('—', cols[5], {color:C.gray400}),
          ]});
        }),
      ],
    }),
    sp(120),
    new Paragraph({children:[new PageBreak()]}),
  ];
}

// ────────────────────────────────────────
// 03 バズ投稿TOP3
// ────────────────────────────────────────
function buildBuzz(monthData) {
  const top3 = monthData.kpi.top3;
  const rankColors = [C.orange,'888888','8B5E1A'];
  const medals = ['🥇','🥈','🥉'];

  const cards = top3.map((post, idx) => {
    const bc = rankColors[idx];
    const lines = post.content.split('\n').filter(l=>l.trim()).slice(0,5);
    return [
      new Table({
        width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
        rows:[
          new TableRow({children:[new TableCell({
            borders:{top:{style:BorderStyle.SINGLE,size:6,color:bc},bottom:BD_NONE.bottom,left:{style:BorderStyle.SINGLE,size:6,color:bc},right:{style:BorderStyle.SINGLE,size:6,color:bc}},
            shading:{fill:C.gray100,type:ShadingType.CLEAR},
            margins:{top:80,bottom:80,left:200,right:200},
            children:[p([
              r(`${medals[idx]} 第${idx+1}位`, {bold:true,color:bc,size:22}),
              r(`　${post.dateStr||''}　　`, {size:18,color:C.gray700}),
              r(`❤️ ${post.likes.toLocaleString()}　💬 ${post.comments}　🔁 ${post.reposts}　🏆 ${post.score.toLocaleString()}pt`, {bold:true,color:C.primary,size:18}),
            ], {after:0})],
          })]})
          ,
          new TableRow({children:[new TableCell({
            borders:{top:BD_NONE.top,bottom:{style:BorderStyle.SINGLE,size:2,color:C.gray200},left:{style:BorderStyle.SINGLE,size:6,color:bc},right:{style:BorderStyle.SINGLE,size:6,color:bc}},
            margins:{top:60,bottom:80,left:200,right:200},
            children:lines.map((line,i)=>p(r(line,{size:18}),{after:i<lines.length-1?40:0,line:300})),
          })]})
          ,
          new TableRow({children:[new TableCell({
            borders:{top:BD_NONE.top,bottom:{style:BorderStyle.SINGLE,size:6,color:bc},left:{style:BorderStyle.SINGLE,size:6,color:bc},right:{style:BorderStyle.SINGLE,size:6,color:bc}},
            shading:{fill:'FEF2F2',type:ShadingType.CLEAR},
            margins:{top:80,bottom:100,left:200,right:200},
            children:[p([
              r('📋 所感：', {bold:true,size:17,color:C.red}),
              r('（この欄に担当ディレクターがコメントを記入）', {size:17,color:C.gray400}),
            ], {after:0})],
          })]})
        ],
      }),
      sp(100),
    ];
  });

  return [
    h1('02', 'バズ投稿 TOP 3'),
    sp(80),
    ...cards.flat(),
    new Paragraph({children:[new PageBreak()]}),
  ];
}

// ────────────────────────────────────────
// 04 コンテンツ種別分析
// ────────────────────────────────────────
function buildContentAnalysis(monthData) {
  const { kpi } = monthData;
  const types = Object.entries(kpi.byType).sort((a,b)=>b[1].count-a[1].count);
  const typeColors = {'インサイト祭り':'E2D9F3','サロン経営Tips':'FDDCDC','あるある・共感':'D1ECF1','雑談・日常':'FFF3CD'};
  const cols = [2200,800,1160,1160,1160,2880];

  return [
    h1('03', 'コンテンツ種別 パフォーマンス分析'),
    sp(80),
    new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:cols,
      rows:[
        new TableRow({children:[th('種別',cols[0]),th('件数',cols[1]),th('いいね合計',cols[2]),th('いいね平均',cols[3]),th('スコア平均',cols[4]),th('評価',cols[5])]}),
        ...types.map(([type,d])=>new TableRow({children:[
          td(type,cols[0],{fill:typeColors[type]||C.gray100}),
          td(String(d.count),cols[1],{align:AlignmentType.CENTER}),
          td(d.likes.toLocaleString(),cols[2],{align:AlignmentType.CENTER}),
          td(String(d.avgLikes),cols[3],{align:AlignmentType.CENTER,bold:d.avgLikes>=2,color:d.avgLikes>=2?C.green:C.black}),
          td(String(d.avgScore),cols[4],{align:AlignmentType.CENTER,bold:d.avgScore>=2,color:d.avgScore>=2?C.green:C.black}),
          td('—',cols[5],{color:C.gray400}),
        ]})),
      ],
    }),
    sp(120),
    new Paragraph({children:[new PageBreak()]}),
  ];
}

// ────────────────────────────────────────
// 05 課題・翌月方針
// ────────────────────────────────────────
function buildNextMonth(account, monthData) {
  const { year, month } = monthData;
  const nextMonth = month === 12 ? 1 : month + 1;
  const nextYear  = month === 12 ? year + 1 : year;

  return [
    h1('04', `課題総括・${nextYear}年${nextMonth}月 運用方針`),
    sp(80),
    h2('今月の課題'),
    sp(60),
    lbox([
      '【HIGH】課題タイトル ← ディレクターが記入',
      '【MED】課題タイトル',
      '【LOW】課題タイトル',
    ], C.red, 'FEF2F2', C.black),
    sp(120),
    h2(`${nextYear}年${nextMonth}月 アクションプラン`),
    sp(60),
    new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:[2400,6960],
      rows:[
        new TableRow({children:[th('優先度',2400),th('アクション内容',6960,AlignmentType.LEFT)]}),
        ...[['HIGH',''], ['MED',''], ['LOW','']].map(([pri, action])=>new TableRow({children:[
          td(pri, 2400, {align:AlignmentType.CENTER, bold:true, color: pri==='HIGH'?C.red:pri==='MED'?C.orange:C.green}),
          td(action||'（記入）', 6960, {color: action?C.black:C.gray400}),
        ]})),
      ],
    }),
    sp(200), hr(C.gray200), sp(80),
    new Paragraph({alignment:AlignmentType.CENTER, children:[
      r(`本レポートに関するご質問は担当ディレクターまでお気軽にご連絡ください。`, {color:C.gray400,size:18}),
    ]}),
  ];
}

// ────────────────────────────────────────
// メイン関数
// ────────────────────────────────────────
function generateMonthlyReport(account, monthData, outputPath) {
  const hf = makeHeaderFooter(account, monthData.periodLabel);
  const doc = new Document({
    styles:{default:{document:{run:{font:FONT,size:20,color:C.black}}}},
    sections:[{
      properties:{...PAGE_PROPS},
      ...hf,
      children:[
        ...buildSummary(account, monthData),
        ...buildWeeklyTrend(monthData),
        ...buildBuzz(monthData),
        ...buildContentAnalysis(monthData),
        ...buildNextMonth(account, monthData),
      ],
    }],
  });

  return Packer.toBuffer(doc).then(buf=>{
    fs.mkdirSync(path.dirname(outputPath),{recursive:true});
    fs.writeFileSync(outputPath, buf);
    console.log(`  ✅ 月次レポート: ${outputPath}`);
  });
}

module.exports = { generateMonthlyReport };
