/**
 * generate_reports.js
 * 全アカウント（またはひとつ）のレポートを一括生成する
 *
 * 使い方:
 *   # 全アカウントの週次レポートを生成
 *   node generate_reports.js --type weekly
 *
 *   # 全アカウントの月次レポートを生成（今月）
 *   node generate_reports.js --type monthly
 *
 *   # 特定アカウントのみ
 *   node generate_reports.js --type weekly --account salon_ayano_diary
 *
 *   # 月を指定（YYYY-MM）
 *   node generate_reports.js --type monthly --month 2026-04
 */

const path = require('path');
const fs   = require('fs');
const { getWeekData, getMonthData } = require('./lib/data-loader');
const { generateWeeklyReport }   = require('./lib/report-weekly');
const { generateMonthlyReport }  = require('./lib/report-monthly');
const { analyzeAccount }         = require('./lib/analyzer');
const { generateHearingSheet }   = require('./lib/report-hearing');
const { generateStrategyDoc }    = require('./lib/report-strategy');

// ── 引数パース ───────────────────────────
const args = process.argv.slice(2);
const get  = (flag) => { const i = args.indexOf(flag); return i > -1 ? args[i+1] : null; };

const TYPE     = get('--type')    || 'weekly';
const ACCOUNT  = get('--account') || 'all';
const MONTH    = get('--month');   // "2026-04"
const OUTPUT   = get('--out')     || path.join(__dirname, 'output');

// ── accounts.json 読み込み ───────────────
const accounts = JSON.parse(fs.readFileSync(path.join(__dirname, 'accounts.json'), 'utf-8')).accounts;
const targets  = ACCOUNT === 'all' ? accounts : accounts.filter(a => a.id === ACCOUNT);

if (targets.length === 0) {
  console.error(`❌ アカウントが見つかりません: ${ACCOUNT}`);
  console.error(`利用可能: ${accounts.map(a=>a.id).join(', ')}`);
  process.exit(1);
}

// ── 月次の年月パース ─────────────────────
let targetYear, targetMonth;
if (MONTH) {
  [targetYear, targetMonth] = MONTH.split('-').map(Number);
} else {
  const now = new Date();
  targetYear  = now.getFullYear();
  targetMonth = now.getMonth() + 1;
}

// ── メイン処理 ───────────────────────────
async function main() {
  console.log(`\n🚀 レポート生成開始 [${TYPE}] ${targets.length}アカウント\n`);

  for (const account of targets) {
    console.log(`📊 ${account.handle} (${account.id})`);

    try {
      const accountOut = path.join(OUTPUT, account.id);

      if (TYPE === 'hearing') {
        const analysis  = analyzeAccount(account);
        const outputPath = path.join(accountOut, 'hearing', `hearing_${dateTag()}.docx`);
        await generateHearingSheet(analysis, outputPath);

      } else if (TYPE === 'strategy') {
        const analysis  = analyzeAccount(account);
        const outputPath = path.join(accountOut, 'strategy', `strategy_${dateTag()}.docx`);
        await generateStrategyDoc(analysis, outputPath);

      } else if (TYPE === 'weekly') {
        const weekData   = getWeekData(account, -1); // 先週
        const wTag       = weekData.monday.getFullYear() + String(weekData.monday.getMonth()+1).padStart(2,'0') + String(weekData.monday.getDate()).padStart(2,'0');
        const outputPath = path.join(accountOut, 'weekly', 'weekly_report_' + wTag + '.docx');
        await generateWeeklyReport(account, weekData, outputPath);

      } else if (TYPE === 'monthly') {
        const monthData  = getMonthData(account, targetYear, targetMonth);
        const outputPath = path.join(accountOut, 'monthly', `monthly_report_${targetYear}${String(targetMonth).padStart(2,'0')}.docx`);
        await generateMonthlyReport(account, monthData, outputPath);

      } else {
        console.error(`❌ 不明なtype: ${TYPE}. weekly / monthly / hearing / strategy を指定してください。`);
        process.exit(1);
      }

    } catch (err) {
      console.error(`  ❌ エラー (${account.id}): ${err.message}`);
    }
  }

  console.log(`\n✅ 完了。出力先: ${OUTPUT}\n`);
}

function dateTag() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
}

main();
