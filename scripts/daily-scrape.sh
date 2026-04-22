#!/bin/bash
# daily-scrape.sh
# 全アクティブアカウントを毎日スクレイピングする
# cron設定例: 0 3 * * * /Users/yuya/Desktop/threads-auto/scripts/daily-scrape.sh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$SCRIPT_DIR/.."
ACCOUNTS_FILE="$ROOT_DIR/accounts.json"
SCRAPER_DIR="/Users/yuya/poc_threads_scraping_ver3"
LOG_DIR="$ROOT_DIR/logs/scraping"
DATE_TAG=$(date +%Y%m%d_%H%M%S)
LOG_FILE="$LOG_DIR/scrape_$DATE_TAG.log"
MAX_POSTS=${MAX_POSTS:-100}
INTERVAL_SEC=${INTERVAL_SEC:-15}  # アカウント間インターバル（レート制限対策）

mkdir -p "$LOG_DIR"

log() { echo "$1" | tee -a "$LOG_FILE"; }

log "================================================================"
log "日次スクレイピング開始: $(date '+%Y-%m-%d %H:%M:%S')"
log "================================================================"

# accounts.json からアクティブアカウントを取得
ACCOUNTS_JSON=$(node -e "
const fs = require('fs');
const data = JSON.parse(fs.readFileSync('$ACCOUNTS_FILE','utf-8'));
const active = data.accounts.filter(a => a.status === 'active');
console.log(JSON.stringify(active.map(a => ({id:a.id, username:a.scrape_username}))));
")

ACCOUNT_COUNT=$(echo "$ACCOUNTS_JSON" | node -e "process.stdin.resume();let d='';process.stdin.on('data',c=>d+=c);process.stdin.on('end',()=>console.log(JSON.parse(d).length));")
log "対象アカウント数: $ACCOUNT_COUNT"
log ""

# SingletonLock をクリア
LOCK_FILE="$SCRAPER_DIR/browser_session/SingletonLock"
if [ -f "$LOCK_FILE" ]; then
    rm -f "$LOCK_FILE"
    log "🔓 SingletonLock をクリアしました"
fi

cd "$SCRAPER_DIR"

SUCCESS=0
FAILED=0
FAILED_LIST=""

# 各アカウントをスクレイピング
echo "$ACCOUNTS_JSON" | node -e "
process.stdin.resume();
let d='';
process.stdin.on('data',c=>d+=c);
process.stdin.on('end',()=>{
  JSON.parse(d).forEach(a=>console.log(a.id + ' ' + a.username));
});
" | while read -r ACCOUNT_ID USERNAME; do
    log "--- [$ACCOUNT_ID] @$USERNAME スクレイピング中 ---"

    # SingletonLock の再チェック
    [ -f "$LOCK_FILE" ] && rm -f "$LOCK_FILE"

    if make run-scraping USERS="$USERNAME" MAX="$MAX_POSTS" >> "$LOG_FILE" 2>&1; then
        log "  ✅ 完了"
        SUCCESS=$((SUCCESS+1))
    else
        log "  ❌ エラー（詳細はログを確認）"
        FAILED=$((FAILED+1))
        FAILED_LIST="$FAILED_LIST $ACCOUNT_ID"
    fi

    log ""
    sleep "$INTERVAL_SEC"
done

log "================================================================"
log "スクレイピング完了: $(date '+%Y-%m-%d %H:%M:%S')"
log "成功: $SUCCESS / 失敗: $FAILED"
[ -n "$FAILED_LIST" ] && log "失敗アカウント: $FAILED_LIST"
log "================================================================"

# レポート自動生成（オプション）
if [ "${AUTO_REPORT:-0}" = "1" ]; then
    log ""
    log "📊 レポート自動生成を開始..."
    cd "$ROOT_DIR"
    NODE_PATH=/opt/homebrew/lib/node_modules node generate_reports.js --type weekly 2>&1 | tee -a "$LOG_FILE"
fi

exit 0
