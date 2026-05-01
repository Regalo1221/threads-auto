/**
 * data-loader.js
 * スクレイピングCSVを読み込み、週次・月次データに集計する
 */
const fs   = require('fs');
const path = require('path');

// ────────────────────────────────────────
// CSVの最新ファイルを取得
// ────────────────────────────────────────
function findLatestCsv(csvDir, username) {
  if (!fs.existsSync(csvDir)) return null;
  const files = fs.readdirSync(csvDir)
    .filter(f => f.startsWith(username + '_') && f.endsWith('.csv'))
    .sort()
    .reverse();
  return files.length ? path.join(csvDir, files[0]) : null;
}

// ────────────────────────────────────────
// RFC 4180 準拠のCSVパーサー
// 投稿内容に改行・カンマ・ダブルクォートが含まれても正しく処理する
// ────────────────────────────────────────
function parseCsvRows(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuote = false;

  for (let i = 0; i < text.length; i++) {
    const ch   = text[i];
    const next = text[i + 1];

    if (inQuote) {
      if (ch === '"' && next === '"') { field += '"'; i++; }  // エスケープ ""
      else if (ch === '"')            { inQuote = false; }    // クォート終了
      else                            { field += ch; }
    } else {
      if      (ch === '"')  { inQuote = true; }
      else if (ch === ',')  { row.push(field); field = ''; }
      else if (ch === '\r' && next === '\n') {
        row.push(field); field = '';
        rows.push(row); row = []; i++;
      }
      else if (ch === '\n') { row.push(field); field = ''; rows.push(row); row = []; }
      else                  { field += ch; }
    }
  }
  // 末尾の行
  if (field !== '' || row.length > 0) { row.push(field); rows.push(row); }
  return rows;
}

// ────────────────────────────────────────
// CSV読み込み → 投稿配列
// ────────────────────────────────────────
function loadCsv(csvPath) {
  if (!csvPath || !fs.existsSync(csvPath)) return [];

  // BOM除去
  const raw  = fs.readFileSync(csvPath, 'utf-8').replace(/^﻿/, '');
  const rows = parseCsvRows(raw.trim());
  if (rows.length < 2) return [];

  // ヘッダー行でカラム名→インデックスのマップを作る（列順に依存しない）
  const header = rows[0].map(h => h.trim());
  const idx = {};
  header.forEach((h, i) => { idx[h] = i; });

  const iDate     = idx['投稿日時']  !== undefined ? idx['投稿日時']  : 0;
  const iContent  = idx['投稿内容']  !== undefined ? idx['投稿内容']  : 1;
  const iLikes    = idx['いいね']    !== undefined ? idx['いいね']    :
                    idx['いいね数']  !== undefined ? idx['いいね数']  : 5;
  const iComments = idx['コメント']  !== undefined ? idx['コメント']  :
                    idx['コメント数']!== undefined ? idx['コメント数']: 6;
  const iReposts  = idx['リポスト']  !== undefined ? idx['リポスト']  :
                    idx['リポスト数']!== undefined ? idx['リポスト数']: 7;

  return rows.slice(1).map(cols => {
    const get = (i) => (cols[i] || '').trim();
    const dateStr  = get(iDate);
    const content  = get(iContent);
    const likes    = parseInt(get(iLikes),    10) || 0;
    const comments = parseInt(get(iComments), 10) || 0;
    const reposts  = parseInt(get(iReposts),  10) || 0;

    return {
      dateStr,
      date:     parseDate(dateStr),
      content,
      likes,
      comments,
      reposts,
      score: likes * 1 + comments * 3 + reposts * 5,
      type: categorize(content),
    };
  }).filter(p => p.date);
}

// ────────────────────────────────────────
// 日付パース（YYYY/MM/DD HH:mm などに対応）
// ────────────────────────────────────────
function parseDate(str) {
  if (!str) return null;
  const m = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (!m) return null;
  return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
}

// ────────────────────────────────────────
// 週番号（ISO週）
// ────────────────────────────────────────
function isoWeek(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

// ────────────────────────────────────────
// コンテンツ種別の自動分類（キーワードベース）
// ────────────────────────────────────────
const CATEGORIES = [
  { type: 'インサイト祭り', keywords: ['インサイト祭り', 'インサイト参加'] },
  { type: 'サロン経営Tips', keywords: ['サロン', '経営', 'リピート', '売上', '客単価', '集客', '予約', '施術', '開業'] },
  { type: 'あるある・共感', keywords: ['あります？', 'ありません？', 'ありますか？', 'いません？', 'いますか？', 'あるある', 'わかる', '同じ'] },
  { type: '雑談・日常', keywords: [] }, // デフォルト
];

function categorize(content) {
  for (const cat of CATEGORIES) {
    if (cat.keywords.length === 0) continue;
    if (cat.keywords.some(kw => content.includes(kw))) return cat.type;
  }
  return '雑談・日常';
}

// ────────────────────────────────────────
// 週範囲フィルタ
// ────────────────────────────────────────
function filterByDateRange(posts, from, to) {
  return posts.filter(p => p.date && p.date >= from && p.date <= to);
}

// ────────────────────────────────────────
// KPI集計
// ────────────────────────────────────────
function computeKPIs(posts) {
  const total    = posts.length;
  const likes    = posts.reduce((s, p) => s + p.likes,    0);
  const comments = posts.reduce((s, p) => s + p.comments, 0);
  const reposts  = posts.reduce((s, p) => s + p.reposts,  0);
  const score    = posts.reduce((s, p) => s + p.score,    0);
  const zeroLikeRate = total ? Math.round(posts.filter(p => p.likes === 0).length / total * 100) : 0;

  // TOP3
  const top3 = [...posts].sort((a, b) => b.score - a.score).slice(0, 3);

  // 種別集計
  const byType = {};
  posts.forEach(post => {
    if (!byType[post.type]) byType[post.type] = { count: 0, likes: 0, score: 0 };
    byType[post.type].count++;
    byType[post.type].likes += post.likes;
    byType[post.type].score += post.score;
  });
  Object.values(byType).forEach(t => {
    t.avgLikes = t.count ? +(t.likes / t.count).toFixed(1) : 0;
    t.avgScore = t.count ? +(t.score / t.count).toFixed(1) : 0;
  });

  return { total, likes, comments, reposts, score, zeroLikeRate, top3, byType };
}

// ────────────────────────────────────────
// 週次データ取得（weekOffset=-1で先週）
// ────────────────────────────────────────
function getWeekData(account, weekOffset = 0) {
  const csvPath = findLatestCsv(account.csv_dir, account.scrape_username);
  const all     = loadCsv(csvPath);

  const today  = new Date();
  const monday = new Date(today);
  monday.setDate(today.getDate() - today.getDay() - 6 + weekOffset * 7);
  monday.setHours(0, 0, 0, 0);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  sunday.setHours(23, 59, 59, 999);

  const thisPosts  = filterByDateRange(all, monday, sunday);
  const prevMonday = new Date(monday); prevMonday.setDate(monday.getDate() - 7);
  const prevSunday = new Date(sunday); prevSunday.setDate(sunday.getDate() - 7);
  const prevPosts  = filterByDateRange(all, prevMonday, prevSunday);

  const fmt = d => `${d.getMonth() + 1}/${d.getDate()}`;
  const periodLabel = `${monday.getFullYear()}年${fmt(monday)}（月）〜${fmt(sunday)}（日）`;

  return {
    csvPath,
    monday, sunday,
    periodLabel,
    posts:    thisPosts,
    prevPosts: prevPosts,
    kpi:      computeKPIs(thisPosts),
    prevKpi:  computeKPIs(prevPosts),
  };
}

// ────────────────────────────────────────
// 月次データ取得
// ────────────────────────────────────────
function getMonthData(account, year, month) {
  const csvPath = findLatestCsv(account.csv_dir, account.scrape_username);
  const all     = loadCsv(csvPath);

  const from  = new Date(year, month - 1, 1);
  const to    = new Date(year, month, 0, 23, 59, 59);
  const posts = filterByDateRange(all, from, to);

  // 週別集計
  const byWeek = {};
  posts.forEach(post => {
    const w = `W${String(isoWeek(post.date)).padStart(2, '0')}`;
    if (!byWeek[w]) byWeek[w] = [];
    byWeek[w].push(post);
  });
  const weeklyData = Object.entries(byWeek)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([week, ps]) => ({ week, ...computeKPIs(ps) }));

  return {
    csvPath,
    year, month,
    periodLabel: `${year}年${month}月`,
    posts,
    kpi:        computeKPIs(posts),
    weeklyData,
  };
}

module.exports = { findLatestCsv, loadCsv, getWeekData, getMonthData, computeKPIs, categorize };
