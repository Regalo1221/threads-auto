/**
 * report-strategy.js
 * 分析結果 → 戦略書DOCX（ヒアリング前に確定できる骨格 + 要確認欄）
 */
const { Document, Packer, PageBreak, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign } = require('docx');
const fs   = require('fs');
const path = require('path');
const {
  C, FONT, BD_NONE, BD_GRAY,
  r, p, sp, hr, h1, h2, th, td,
  lbox, makeHeaderFooter, PAGE_PROPS,
} = require('./docx-base');

// 要確認マーク付きセル
function tdTodo(text, w, note) {
  return new TableCell({
    borders: {
      top:    { style: BorderStyle.SINGLE, size: 2, color: C.orange },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: C.orange },
      left:   { style: BorderStyle.SINGLE, size: 6, color: C.orange },
      right:  { style: BorderStyle.SINGLE, size: 2, color: C.orange },
    },
    width: w ? { size: w, type: WidthType.DXA } : undefined,
    shading: { fill: 'FFFBF0', type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    children: [
      p(r(text || '【ヒアリング後に記入】', { size: 17, color: text ? C.black : C.orange }), { after: note ? 30 : 0 }),
      ...(note ? [p(r(`📋 ${note}`, { size: 14, color: C.gray400 }), { after: 0 })] : []),
    ],
  });
}

// ─────────────────────────────────────────────
// 00 表紙
// ─────────────────────────────────────────────
function buildCover(analysisResult) {
  const { account, issues } = analysisResult;
  const highCount = issues.filter(i => i.severity === 'HIGH' || i.severity === 'CRITICAL').length;

  return [
    new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
      rows:[new TableRow({children:[new TableCell({
        borders:BD_NONE, width:{size:9360,type:WidthType.DXA},
        shading:{fill:C.primary,type:ShadingType.CLEAR},
        margins:{top:240,bottom:240,left:360,right:360},
        children:[
          p(r('アカウント運用 戦略書', {bold:true,color:C.white,size:32}), {after:80}),
          p(r(`${account.handle}｜${account.business_type||'—'}`, {color:'B0C4DE',size:19}), {after:0}),
        ],
      })]})],
    }),
    sp(160),
    new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:[4680,4680],
      rows:[new TableRow({children:[
        new TableCell({borders:BD_NONE, width:{size:4680,type:WidthType.DXA}, children:[
          p(r('作成日', {size:17,color:C.gray400}), {after:40}),
          p(r(todayStr(), {bold:true,size:20}), {after:0}),
        ]}),
        new TableCell({borders:BD_NONE, width:{size:4680,type:WidthType.DXA}, children:[
          p(r('ステータス', {size:17,color:C.gray400}), {after:40}),
          p(r(highCount > 0 ? `要解決課題 ${highCount}件 ／ ヒアリング待ち項目あり` : 'ヒアリング済み確定版', {bold:true,size:20,color:highCount>0?C.orange:C.green}), {after:0}),
        ]}),
      ]})],
    }),
    sp(160),
    lbox([
      '【凡例】　オレンジ枠 = ヒアリング後に記入必要な項目',
      '　　　　　通常枠 = データ分析から確定済みの方針',
    ], C.orange, 'FFFBF0'),
    new Paragraph({children:[new PageBreak()]}),
  ];
}

// ─────────────────────────────────────────────
// 01 問題の本質診断
// ─────────────────────────────────────────────
function buildDiagnosis(analysisResult) {
  const { account, kpi, issues } = analysisResult;
  const content = [
    h1('01', '現状診断'),
    sp(80),
    h2('データから見えている問題'),
    sp(60),
  ];

  if (issues.length === 0) {
    content.push(p(r('現時点では大きな課題は検出されていません。', {color:C.gray400})));
  } else {
    const cols = [1000, 2400, 5960];
    content.push(
      new Table({
        width:{size:9360,type:WidthType.DXA}, columnWidths:cols,
        rows:[
          new TableRow({children:[th('重要度',cols[0]),th('カテゴリ',cols[1]),th('課題内容',cols[2],AlignmentType.LEFT)]}),
          ...issues.map(issue => {
            const sc={CRITICAL:'FDDCDC',HIGH:'FDDCDC',MED:'FFF3CD',LOW:'D4EDDA'};
            const tc={CRITICAL:C.red,HIGH:C.red,MED:C.orange,LOW:C.green};
            return new TableRow({children:[
              td(issue.severity, cols[0], {fill:sc[issue.severity],bold:true,color:tc[issue.severity],align:AlignmentType.CENTER}),
              td(issue.category, cols[1], {fill:C.gray100}),
              td(issue.title,    cols[2]),
            ]});
          }),
        ],
      }),
    );
  }

  content.push(
    sp(120),
    h2('本質的な問題（仮説）'),
    sp(60),
    lbox([
      account.content_pillars?.[0] !== '柱①'
        ? `コンテンツ方針は設定済みですが、エンゲージメントデータと目標の整合性を確認する必要があります。`
        : '投稿の方向性・ターゲット・ゴールの3つがアラインされていない状態です。何を投稿するかより「誰に・何のために」を先に定義する必要があります。',
    ], C.red, 'FEF2F2'),
    sp(80),
    new Paragraph({children:[new PageBreak()]}),
  );

  return content;
}

// ─────────────────────────────────────────────
// 02 アカウント再定義
// ─────────────────────────────────────────────
function buildRepositioning(analysisResult) {
  const { account } = analysisResult;
  const cols = [2800, 6560];

  return [
    h1('02', 'アカウント再定義'),
    sp(80),
    new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:cols,
      rows:[
        new TableRow({children:[th('項目',cols[0]),th('内容',cols[1],AlignmentType.LEFT)]}),
        new TableRow({children:[
          td('アカウントの発信テーマ', cols[0], {fill:C.gray100,bold:true}),
          tdTodo(null, cols[1], 'ヒアリングで確定：「○○な人が○○を達成するまでの話」の形式で'),
        ]}),
        new TableRow({children:[
          td('話しかける相手（ターゲット）', cols[0], {fill:C.gray100,bold:true}),
          tdTodo(account.target_audience && account.target_audience !== 'ターゲット層を入力' ? account.target_audience : null, cols[1], 'ヒアリングセクションCで確定'),
        ]}),
        new TableRow({children:[
          td('このアカウントが与える価値', cols[0], {fill:C.gray100,bold:true}),
          tdTodo(null, cols[1], 'ヒアリングセクションDで確定'),
        ]}),
        new TableRow({children:[
          td('最終的な誘導先', cols[0], {fill:C.gray100,bold:true}),
          tdTodo(account.goal && account.goal !== '目標を入力' ? account.goal : null, cols[1], 'ヒアリングセクションBで確定'),
        ]}),
        new TableRow({children:[
          td('KPI（何件で成功か）', cols[0], {fill:C.gray100,bold:true}),
          tdTodo(account.kpi_target && account.kpi_target > 0 ? `${account.kpi_label}：月${account.kpi_target}件以上` : null, cols[1], 'ヒアリングセクションBで確定'),
        ]}),
      ],
    }),
    sp(120),
    new Paragraph({children:[new PageBreak()]}),
  ];
}

// ─────────────────────────────────────────────
// 03 コンテンツ4本柱
// ─────────────────────────────────────────────
function buildContentPillars(analysisResult) {
  const { account, kpi } = analysisResult;
  const hasPillars = account.content_pillars && account.content_pillars[0] !== '柱①';
  const cols = [800, 2200, 1000, 1000, 4360];

  const pillars = hasPillars
    ? account.content_pillars.map((name, i) => {
        const schedules = [
          {count:'週6本', time:'昼10時中心'},
          {count:'週6本', time:'朝8時中心'},
          {count:'週3本', time:'夜20時中心'},
          {count:'週3本', time:'任意'},
        ];
        return { name, ...schedules[i] || {count:'週3本', time:'任意'}, point: '（ヒアリング後に記入）' };
      })
    : [
        {name:'【ヒアリング後に設定】ノウハウ・専門情報',  count:'週6本', time:'昼10時中心', point:'専門性・信頼性の構築'},
        {name:'【ヒアリング後に設定】あるある・共感',      count:'週6本', time:'朝8時中心',  point:'ターゲットフォロワー獲得'},
        {name:'インサイト祭り（集客拡大）',                count:'週3本', time:'夜20時中心', point:'毎回異なる文章でローテーション'},
        {name:'事例・変化の報告',                          count:'週3本', time:'任意',        point:'「本当に変わる」という信頼を作る'},
      ];

  return [
    h1('03', 'コンテンツ4本柱'),
    sp(80),
    new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:cols,
      rows:[
        new TableRow({children:[th('#',cols[0]),th('コンテンツ種別',cols[1]),th('週本数',cols[2]),th('投稿時間',cols[3]),th('ポイント',cols[4],AlignmentType.LEFT)]}),
        ...pillars.map((pillar, i) => {
          const isTodo = pillar.name.startsWith('【ヒアリング後');
          return new TableRow({children:[
            td(String(i+1), cols[0], {align:AlignmentType.CENTER,bold:true}),
            isTodo ? tdTodo(null, cols[1], '業種・強みをヒアリング後に設計') : td(pillar.name, cols[1], {fill: i===2?'E2D9F3':'F4F6F8'}),
            td(pillar.count, cols[2], {align:AlignmentType.CENTER}),
            td(pillar.time,  cols[3], {align:AlignmentType.CENTER}),
            td(pillar.point, cols[4], {color:isTodo?C.orange:C.gray700}),
          ]});
        }),
      ],
    }),
    sp(80),
    p(r('※ 週計21投稿（1日3投稿 × 7日）を基本スケジュールとする', {size:16,color:C.gray400}), {after:0}),
    sp(120),
    new Paragraph({children:[new PageBreak()]}),
  ];
}

// ─────────────────────────────────────────────
// 04 週間投稿カレンダー
// ─────────────────────────────────────────────
function buildWeeklyCalendar(analysisResult) {
  const { account } = analysisResult;
  const times = account.post_times || ['08:00', '10:00', '20:00'];
  const days = ['月','火','水','木','金','土','日'];
  const cols = [720, 2880, 2880, 2880];

  // 基本カレンダー（ヒアリング前はプレースホルダー）
  const calRows = days.map((day, i) => {
    return new TableRow({children:[
      td(day, cols[0], {align:AlignmentType.CENTER,bold:true}),
      tdTodo(null, cols[1], `${times[0]} 投稿`),
      tdTodo(null, cols[2], `${times[1]} 投稿`),
      tdTodo(null, cols[3], `${times[2]} 投稿`),
    ]});
  });

  return [
    h1('04', '週間投稿カレンダー骨格'),
    sp(80),
    new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:cols,
      rows:[
        new TableRow({children:[th('曜日',cols[0]),th(times[0],cols[1],AlignmentType.LEFT),th(times[1],cols[2],AlignmentType.LEFT),th(times[2],cols[3],AlignmentType.LEFT)]}),
        ...calRows,
      ],
    }),
    sp(80),
    lbox([
      '【オレンジ枠の項目はヒアリング後に確定します】',
      `• 朝${times[0]}：あるある・共感系（ターゲットにフォローしてもらう）`,
      `• 昼${times[1]}：ノウハウ・専門性系（信頼構築）`,
      `• 夜${times[2]}：インサイト祭り / 事例報告（集客・DM誘導）`,
    ], C.orange, 'FFFBF0'),
    sp(120),
    new Paragraph({children:[new PageBreak()]}),
  ];
}

// ─────────────────────────────────────────────
// 05 プロフィール改訂案
// ─────────────────────────────────────────────
function buildProfile(analysisResult) {
  const { account } = analysisResult;
  const cols = [2400, 6960];

  return [
    h1('05', 'プロフィール改訂方針'),
    sp(80),
    new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:cols,
      rows:[
        new TableRow({children:[th('項目',cols[0]),th('方針・内容',cols[1],AlignmentType.LEFT)]}),
        new TableRow({children:[td('顔出し', cols[0], {fill:C.gray100,bold:true}), tdTodo(null,cols[1],'ヒアリングセクションEで確定')]}),
        new TableRow({children:[td('名前・ニックネーム', cols[0], {fill:C.gray100,bold:true}), tdTodo(null,cols[1],'ヒアリングセクションEで確定')]}),
        new TableRow({children:[td('肩書き（1行目）', cols[0], {fill:C.gray100,bold:true}), tdTodo(null,cols[1],'例：「大阪の個人サロン経営者｜月商40〜50万安定の作り方を発信」')]}),
        new TableRow({children:[td('実績・強み（2〜3行）', cols[0], {fill:C.gray100,bold:true}), tdTodo(null,cols[1],'ヒアリングセクションDで確定。具体的な数字を入れる')]}),
        new TableRow({children:[td('CTA（最終行）', cols[0], {fill:C.gray100,bold:true}), tdTodo(null,cols[1],'「○○で悩んでいる方はインスタDMへ📩」の形式で')]}),
        new TableRow({children:[td('インスタURL', cols[0], {fill:C.gray100,bold:true}), tdTodo(null,cols[1],'インスタのURLを記載')]}),
      ],
    }),
    sp(120),
    new Paragraph({children:[new PageBreak()]}),
  ];
}

// ─────────────────────────────────────────────
// 06 KPI・ロードマップ
// ─────────────────────────────────────────────
function buildRoadmap(analysisResult) {
  const { account, kpi } = analysisResult;
  const cols = [1400, 2400, 5560];

  return [
    h1('06', 'KPI目標・ロードマップ'),
    sp(80),
    new Table({
      width:{size:9360,type:WidthType.DXA}, columnWidths:cols,
      rows:[
        new TableRow({children:[th('時期',cols[0]),th('目標数値',cols[1]),th('主なアクション',cols[2],AlignmentType.LEFT)]}),
        new TableRow({children:[
          td('〜1ヶ月目', cols[0], {align:AlignmentType.CENTER}),
          tdTodo(null, cols[1], 'フォロワー目標数を設定'),
          td('プロフィール改訂・インサイト祭り週3本・コンテンツ4本柱を始動', cols[2]),
        ]}),
        new TableRow({children:[
          td('〜2ヶ月目', cols[0], {align:AlignmentType.CENTER}),
          tdTodo(null, cols[1], 'DM問い合わせ目標数を設定'),
          td('ノウハウ投稿の専門性を上げ、エンゲージメント率を改善', cols[2]),
        ]}),
        new TableRow({children:[
          td('〜3ヶ月目', cols[0], {align:AlignmentType.CENTER}),
          tdTodo(account.kpi_target > 0 ? `${account.kpi_label}：月${account.kpi_target}件` : null, cols[1], 'ヒアリングで確定'),
          td('KPI（面談・成約）を達成。再現性のある運用フローが確立', cols[2]),
        ]}),
      ],
    }),
    sp(200),
    hr(C.gray200),
    sp(80),
    p(r(`作成日：${todayStr()}　｜　次回更新：ヒアリング後　担当：${account.director||'未設定'}`, {size:16,color:C.gray400}), {align:AlignmentType.CENTER,after:0}),
  ];
}

// ─────────────────────────────────────────────
// メイン関数
// ─────────────────────────────────────────────
function generateStrategyDoc(analysisResult, outputPath) {
  const { account } = analysisResult;
  const hf = makeHeaderFooter(account, `戦略書 ${todayStr()}`);
  const doc = new Document({
    styles:{default:{document:{run:{font:FONT,size:20,color:C.black}}}},
    sections:[{
      properties:{...PAGE_PROPS},
      ...hf,
      children:[
        ...buildCover(analysisResult),
        ...buildDiagnosis(analysisResult),
        ...buildRepositioning(analysisResult),
        ...buildContentPillars(analysisResult),
        ...buildWeeklyCalendar(analysisResult),
        ...buildProfile(analysisResult),
        ...buildRoadmap(analysisResult),
      ],
    }],
  });

  return Packer.toBuffer(doc).then(buf=>{
    fs.mkdirSync(path.dirname(outputPath),{recursive:true});
    fs.writeFileSync(outputPath, buf);
    console.log(`  ✅ 戦略書: ${outputPath}`);
  });
}

function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}年${d.getMonth()+1}月${d.getDate()}日`;
}

module.exports = { generateStrategyDoc };
