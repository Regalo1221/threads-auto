/**
 * docx-base.js
 * 全レポート共通の色定義・ヘルパー関数
 */
const {
  Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak,
  TabStopType, TabStopPosition,
  Header, Footer,
} = require('docx');

const FONT = 'Noto Sans JP';

const C = {
  primary:  '1E3A5F',
  accent:   '2E86AB',
  light:    'EBF4FA',
  red:      'C0392B',
  orange:   'E67E22',
  green:    '27AE60',
  gray100:  'F4F6F8',
  gray200:  'DDE1E7',
  gray400:  '9BA3AF',
  gray700:  '374151',
  white:    'FFFFFF',
  black:    '111827',
  yellow:   'FFF9E6',
  yellowBdr:'E0C000',
};

const BD_NONE = {
  top:    { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};
const BD_GRAY = {
  top:    { style: BorderStyle.SINGLE, size: 2, color: C.gray200 },
  bottom: { style: BorderStyle.SINGLE, size: 2, color: C.gray200 },
  left:   { style: BorderStyle.SINGLE, size: 2, color: C.gray200 },
  right:  { style: BorderStyle.SINGLE, size: 2, color: C.gray200 },
};

// ── テキスト ──────────────────────────────
const r = (text, o = {}) => new TextRun({
  text, font: FONT,
  bold:    o.bold    || false,
  color:   o.color   || C.black,
  size:    o.size    || 19,
  italics: o.italic  || false,
});

// ── パラグラフ ────────────────────────────
const p = (children, o = {}) => {
  const runs = typeof children === 'string' ? [r(children, o)]
             : Array.isArray(children)       ? children
             : [children];
  return new Paragraph({
    alignment: o.align || AlignmentType.LEFT,
    spacing: {
      before: o.before || 0,
      after:  o.after  !== undefined ? o.after : 60,
      line:   o.line   || 276,
    },
    children: runs,
  });
};

// ── スペーサー ────────────────────────────
const sp = (n = 80) => new Paragraph({ spacing: { before: 0, after: n }, children: [] });

// ── 水平線 ────────────────────────────────
const hr = (color = C.gray200, size = 2) => new Paragraph({
  spacing: { before: 0, after: 0 },
  border: { bottom: { style: BorderStyle.SINGLE, size, color } },
  children: [],
});

// ── セクション見出し H1 ───────────────────
const h1 = (num, title, color = C.primary) => new Paragraph({
  spacing: { before: 280, after: 120 },
  border: { left: { style: BorderStyle.SINGLE, size: 24, color } },
  children: [r(`  ${num}｜${title}`, { bold: true, color, size: 26 })],
});

// ── サブ見出し H2 ─────────────────────────
const h2 = (text, color = C.accent) => new Paragraph({
  spacing: { before: 180, after: 60 },
  children: [r(`▌ ${text}`, { bold: true, color, size: 20 })],
});

// ── テーブルヘッダーセル ──────────────────
const th = (text, w, align = AlignmentType.CENTER) => new TableCell({
  borders: BD_NONE,
  width: { size: w, type: WidthType.DXA },
  shading: { fill: C.primary, type: ShadingType.CLEAR },
  margins: { top: 80, bottom: 80, left: 140, right: 140 },
  verticalAlign: VerticalAlign.CENTER,
  children: [p(r(text, { bold: true, color: C.white, size: 17 }), { after: 0, align })],
});

// ── 通常セル ──────────────────────────────
const td = (text, w, o = {}) => new TableCell({
  borders: BD_GRAY,
  width: w ? { size: w, type: WidthType.DXA } : undefined,
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR } : undefined,
  margins: { top: 80, bottom: 80, left: 140, right: 140 },
  verticalAlign: VerticalAlign.CENTER,
  children: [new Paragraph({
    alignment: o.align || AlignmentType.LEFT,
    spacing: { before: 0, after: 0, line: 300 },
    children: typeof text === 'string'
      ? [r(text, { bold: o.bold, color: o.color || C.black, size: o.size || 18 })]
      : Array.isArray(text) ? text : [text],
  })],
});

// ── KPIカード行（4枚）───────────────────
// items: [{label, value, unit, sub, subColor}]
const kpiRow = (items) => {
  const w = Math.floor(9360 / items.length);
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: items.map(() => w),
    rows: [
      new TableRow({
        children: items.map(item => new TableCell({
          borders: {
            top:    { style: BorderStyle.SINGLE, size: 8, color: C.primary },
            bottom: BD_GRAY.bottom,
            left:   BD_GRAY.left,
            right:  BD_GRAY.right,
          },
          width: { size: w, type: WidthType.DXA },
          shading: { fill: C.gray100, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          children: [
            p(r(item.label, { size: 17, color: C.gray400 }), { after: 40 }),
            p([
              r(item.value, { bold: true, size: 48, color: C.primary }),
              r(' ' + item.unit, { size: 18, color: C.gray700 }),
            ], { after: 40 }),
            p(r(item.sub || '', { size: 16, color: item.subColor || C.gray400 }), { after: 0 }),
          ],
        })),
      }),
    ],
  });
};

// ── 強調ボックス（インサイト枠）──────────
const lbox = (lines, borderColor = C.accent, fillColor = C.light, textColor = C.black) => {
  const children = lines.map((line, i) => p(r(line, { size: 18, color: textColor }), {
    after: i < lines.length - 1 ? 60 : 0,
    line: 320,
  }));
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [new TableCell({
          borders: {
            top:    { style: BorderStyle.SINGLE, size: 4, color: borderColor },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: borderColor },
            left:   { style: BorderStyle.SINGLE, size: 20, color: borderColor },
            right:  { style: BorderStyle.SINGLE, size: 4, color: borderColor },
          },
          shading: { fill: fillColor, type: ShadingType.CLEAR },
          margins: { top: 140, bottom: 140, left: 240, right: 240 },
          children,
        })],
      }),
    ],
  });
};

// ── ページヘッダー・フッター生成 ──────────
const makeHeaderFooter = (account, periodLabel) => ({
  headers: {
    default: new Header({
      children: [
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.gray200, space: 1 } },
          spacing: { before: 0, after: 120 },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            new TextRun({ text: `Threads 運用レポート｜${account.handle}`, color: C.gray400, size: 16, font: FONT }),
            new TextRun({ text: `\t${periodLabel}`, color: C.gray400, size: 16, font: FONT }),
          ],
        }),
      ],
    }),
  },
  footers: {
    default: new Footer({
      children: [
        new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.gray200, space: 1 } },
          spacing: { before: 120, after: 0 },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            new TextRun({ text: `SNS運用代行 ｜ CONFIDENTIAL`, color: C.gray400, size: 16, font: FONT }),
            new TextRun({ text: '\tPage ', color: C.gray400, size: 16, font: FONT }),
            new TextRun({ children: [PageNumber.CURRENT], color: C.gray400, size: 16, font: FONT }),
          ],
        }),
      ],
    }),
  },
});

// ── セクションプロパティ（US Letter）────
const PAGE_PROPS = {
  page: {
    size: { width: 12240, height: 15840 },
    margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
  },
};

module.exports = {
  C, FONT, BD_NONE, BD_GRAY,
  r, p, sp, hr, h1, h2, th, td,
  kpiRow, lbox, makeHeaderFooter, PAGE_PROPS,
  // docx classes（呼び出し元で使えるように）
  Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak,
  TabStopType, TabStopPosition,
};
