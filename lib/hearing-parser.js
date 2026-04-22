/**
 * hearing-parser.js
 * input/[id]/hearing.md を読み込んで accounts.json 形式のオブジェクトに変換する
 */

function parseHearingInput(markdown) {
  const result = {};
  const lines = markdown.split('\n');

  let inPillars = false;
  const pillars = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // コメント行スキップ
    if (line.trim().startsWith('<!--') || line.trim().startsWith('#')) continue;

    // content_pillars の配下アイテム（インデント付き "  - "）
    if (inPillars) {
      const itemMatch = line.match(/^\s{2}-\s+(.+)/);
      if (itemMatch) {
        const val = itemMatch[1].trim();
        if (val && !val.startsWith('（例：')) pillars.push(val);
        continue;
      } else {
        // インデントが終わったら pillars モード終了
        if (line.match(/^-\s+\w/)) inPillars = false;
      }
    }

    // "- key: value" 形式のキーバリュー行
    const kvMatch = line.match(/^-\s+(\w+):\s*(.*)$/);
    if (!kvMatch) continue;

    const key   = kvMatch[1].trim();
    const value = kvMatch[2].trim();

    if (key === 'content_pillars') {
      inPillars = true;
      continue;
    }

    // 例文・空欄はスキップ
    if (!value || value.startsWith('（例：')) continue;

    // 数値変換
    if (key === 'weekly_post_count' || key === 'kpi_target_num') {
      const num = parseInt(value, 10);
      if (!isNaN(num)) result[key] = num;
    } else {
      result[key] = value;
    }
  }

  if (pillars.length > 0) result.content_pillars = pillars;

  return result;
}

/**
 * urls.txt を読み込んでURL配列を返す
 */
function parseUrlsTxt(text) {
  return text
    .split('\n')
    .map(l => l.trim())
    .filter(l => l && !l.startsWith('#') && l.startsWith('http'));
}

module.exports = { parseHearingInput, parseUrlsTxt };
