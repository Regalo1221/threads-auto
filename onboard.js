/**
 * onboard.js
 * 新規アカウントのオンボーディングを一括実行する
 *
 * 事前準備:
 *   1. accounts.json にアカウントを追加（id, handle, scrape_username は必須）
 *   2. input/[id]/hearing.md を作成（templates/hearing-input.md をコピーして記入）
 *   3. input/[id]/urls.txt を作成（任意：HP・SNSのURL）
 *
 * 使い方:
 *   node onboard.js --account salon_ayano_diary
 *
 * 実行内容:
 *   1. hearing.md をパース → accounts.json を更新
 *   2. Threads スクレイピング（初回全件取得）
 *   3. データ分析（課題検出・質問自動生成）
 *   4. ヒアリングシート DOCX 生成
 *   5. 戦略書ひな形 DOCX 生成
 *   → output/[id]/ に一式出力
 */

const path    = require('path');
const fs      = require('fs');
const { execSync } = require('child_process');
const { parseHearingInput, parseUrlsTxt } = require('./lib/hearing-parser');
const { analyzeAccount }       = require('./lib/analyzer');
const { generateHearingSheet } = require('./lib/report-hearing');
const { generateStrategyDoc }  = require('./lib/report-strategy');

// ── 引数パース ────────────────────────────────
const args = process.argv.slice(2);
const get  = (flag) => { const i = args.indexOf(flag); return i > -1 ? args[i+1] : null; };

const ACCOUNT_ID   = get('--account');
const SKIP_SCRAPE  = args.includes('--skip-scrape');
const SCRAPE_DIR   = get('--scrape-dir') || '/Users/yuya/poc_threads_scraping_ver3';
const OUTPUT_ROOT  = get('--out') || path.join(__dirname, 'output');
const ROOT         = __dirname;

if (!ACCOUNT_ID) {
  console.error('❌ --account [id] が必要です');
  console.error('例: node onboard.js --account salon_ayano_diary');
  process.exit(1);
}

// ── メイン処理 ────────────────────────────────
async function main() {
  console.log(`\n${'═'.repeat(50)}`);
  console.log(`🚀 オンボーディング開始: ${ACCOUNT_ID}`);
  console.log(`${'═'.repeat(50)}\n`);

  // ── Step 0: accounts.json 読み込み ──────────
  const accountsPath = path.join(ROOT, 'accounts.json');
  const accountsData = JSON.parse(fs.readFileSync(accountsPath, 'utf-8'));
  const idx = accountsData.accounts.findIndex(a => a.id === ACCOUNT_ID);

  if (idx === -1) {
    console.error(`❌ accounts.json に "${ACCOUNT_ID}" が見つかりません`);
    console.error(`登録済みID: ${accountsData.accounts.map(a => a.id).join(', ')}`);
    process.exit(1);
  }

  let account = { ...accountsData.accounts[idx] };

  // ── Step 1: hearing.md パース → accounts.json 更新 ──
  const inputDir   = path.join(ROOT, 'input', ACCOUNT_ID);
  const hearingPath = path.join(inputDir, 'hearing.md');
  const urlsPath    = path.join(inputDir, 'urls.txt');

  console.log('📋 [Step 1] ヒアリングシート読み込み');

  if (fs.existsSync(hearingPath)) {
    const md     = fs.readFileSync(hearingPath, 'utf-8');
    const parsed = parseHearingInput(md);
    const keys   = Object.keys(parsed);

    if (keys.length > 0) {
      account = { ...account, ...parsed };
      accountsData.accounts[idx] = account;
      fs.writeFileSync(accountsPath, JSON.stringify(accountsData, null, 2), 'utf-8');
      console.log(`  ✅ accounts.json を更新（${keys.length}フィールド: ${keys.join(', ')}）`);
    } else {
      console.log('  ℹ️  記入済みフィールドが見つかりませんでした（テンプレートのまま？）');
    }
  } else {
    console.log(`  ⚠️  ${hearingPath} が見つかりません`);
    console.log(`     templates/hearing-input.md をコピーして記入してください`);
  }

  // クライアントURL読み込み（参考情報としてログ出力）
  if (fs.existsSync(urlsPath)) {
    const urls = parseUrlsTxt(fs.readFileSync(urlsPath, 'utf-8'));
    if (urls.length > 0) {
      account.client_urls = urls;
      accountsData.accounts[idx].client_urls = urls;
      fs.writeFileSync(accountsPath, JSON.stringify(accountsData, null, 2), 'utf-8');
      console.log(`  ✅ クライアントURL登録: ${urls.length}件`);
      urls.forEach(u => console.log(`     - ${u}`));
    }
  }

  // ── Step 2: スクレイピング ───────────────────
  console.log('\n🕷️  [Step 2] スクレイピング');

  if (SKIP_SCRAPE) {
    console.log('  ℹ️  --skip-scrape 指定のためスキップ');
  } else {
    const username = account.scrape_username || account.handle.replace('@', '');
    console.log(`  対象ユーザー: ${username}`);

    try {
      // SingletonLock 削除（前回クラッシュ対策）
      execSync(
        `find "${SCRAPE_DIR}" -name "SingletonLock" -delete 2>/dev/null || true`,
        { stdio: 'pipe' }
      );

      execSync(
        `cd "${SCRAPE_DIR}" && make run-scraping USERS=${username} MAX=100`,
        { stdio: 'inherit', timeout: 600000 }
      );
      console.log('  ✅ スクレイピング完了');
    } catch (err) {
      console.error(`  ⚠️  スクレイピングエラー: ${err.message}`);
      console.log('  → データなしで分析を続行します');
    }
  }

  // ── Step 3: データ分析 ───────────────────────
  console.log('\n🔍 [Step 3] データ分析');
  const analysis = analyzeAccount(account);

  console.log(`  投稿データ: ${analysis.posts.length}件`);
  if (analysis.hasData) {
    console.log(`  いいね合計: ${analysis.kpi.likes}`);
    console.log(`  0いいね率:  ${analysis.kpi.zeroLikeRate}%`);
  }
  console.log(`  検出課題:   ${analysis.issues.length}件`);
  analysis.issues.forEach(i => console.log(`    [${i.severity}] ${i.title}`));
  console.log(`  生成質問:   ${analysis.questions.length}件`);

  // 未入力情報の警告
  if (analysis.missingInfo.length > 0) {
    console.log(`\n  ⚠️  未入力フィールド (accounts.json か hearing.md に記入推奨):`);
    analysis.missingInfo.forEach(m => console.log(`    - ${m}`));
  }

  // ── Step 4: ヒアリングシート生成 ─────────────
  console.log('\n📄 [Step 4] ヒアリングシート生成');
  const tag          = dateTag();
  const hearingOut   = path.join(OUTPUT_ROOT, ACCOUNT_ID, 'hearing',  `hearing_${tag}.docx`);
  await generateHearingSheet(analysis, hearingOut);

  // ── Step 5: 戦略書生成 ───────────────────────
  console.log('\n📊 [Step 5] 戦略書生成');
  const strategyOut  = path.join(OUTPUT_ROOT, ACCOUNT_ID, 'strategy', `strategy_${tag}.docx`);
  await generateStrategyDoc(analysis, strategyOut);

  // ── 完了サマリー ─────────────────────────────
  console.log(`\n${'═'.repeat(50)}`);
  console.log('✅ オンボーディング完了！');
  console.log(`${'═'.repeat(50)}`);
  console.log(`\n📁 出力先: output/${ACCOUNT_ID}/`);
  console.log(`   ヒアリングシート: hearing/hearing_${tag}.docx`);
  console.log(`   戦略書ひな形:     strategy/strategy_${tag}.docx`);
  console.log('\n📌 次のステップ:');
  console.log('   1. ヒアリングシートをMTGで使用');
  console.log('   2. hearing.md にMTG回答を追記');
  console.log(`   3. node onboard.js --account ${ACCOUNT_ID} --skip-scrape  （再生成）`);
  console.log('');
}

function dateTag() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
}

main().catch(err => {
  console.error('\n❌ エラー:', err.message);
  process.exit(1);
});
