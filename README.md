# Threads 運用代行 自動化システム

Threadsアカウントのスクレイピング・分析・レポート生成・ヒアリングシート・戦略書を自動生成するシステムです。

---

## 全体フロー

```
① accounts.json にアカウントを追加（チームメンバーの作業）
      ↓
② 毎日 daily-scrape.sh が自動実行（投稿データをCSVに蓄積）
      ↓
③ generate_reports.js でドキュメントを生成
      ↓
  ┌──────────────────────────────────┐
  │ hearing   → ヒアリングシートDOCX │
  │ strategy  → 戦略書DOCX           │
  │ weekly    → 週次レポートDOCX     │
  │ monthly   → 月次レポートDOCX     │
  └──────────────────────────────────┘
```

---

## セットアップ（初回のみ）

### 1. Node.js 依存パッケージのインストール

```bash
npm install -g docx
```

### 2. スクレイパーのセットアップ

スクレイパーは別ディレクトリ `poc_threads_scraping_ver3/` にあります。

```bash
cd /path/to/poc_threads_scraping_ver3
make setup
```

---

## アカウントの追加方法（チームメンバー向け）

### 手順

**1. `accounts.json` を編集する**

リポジトリルートの `accounts.json` を開き、`accounts` 配列に新しいエントリを追加します。

```json
{
  "id": "account_xxx",
  "handle": "@アカウント名",
  "owner_name": "オーナー名",
  "business_name": "サービス名・店舗名",
  "director": "担当ディレクター名",
  "operator": "担当運用者名",
  "platform": "threads",
  "location": "都市名",
  "business_type": "業種",
  "goal": "目標",
  "kpi_label": "KPI指標名",
  "kpi_target": 5,
  "target_audience": "ターゲット層",
  "content_pillars": ["柱①", "柱②", "柱③", "柱④"],
  "weekly_post_count": 21,
  "post_times": ["08:00", "10:00", "20:00"],
  "followers_start": 0,
  "started_at": "2026-04-22",
  "status": "active",
  "csv_dir": "/Users/yuya/poc_threads_scraping_ver3/output",
  "scrape_username": "threads_username_without_at"
}
```

> **`id` はシステム全体でユニークにしてください（出力フォルダ名になります）。**

**2. 動作確認**

```bash
node -e "const a=require('./accounts.json').accounts; console.log(a.map(x=>x.id+' '+x.handle).join('\n'));"
```

**3. ヒアリングシートを生成する**

```bash
NODE_PATH=/opt/homebrew/lib/node_modules node generate_reports.js --type hearing --account [アカウントID]
```

生成先: `output/[アカウントID]/hearing/hearing_YYYYMMDD.docx`

**4. PRを作成してレビューを依頼する**

```bash
git add accounts.json
git commit -m "add account: @アカウント名"
git push origin [ブランチ名]
```

---

## ドキュメント生成コマンド

```bash
# ヒアリングシート（課題ベースの質問を自動生成）
NODE_PATH=/opt/homebrew/lib/node_modules node generate_reports.js --type hearing

# 戦略書
NODE_PATH=/opt/homebrew/lib/node_modules node generate_reports.js --type strategy

# 週次レポート
NODE_PATH=/opt/homebrew/lib/node_modules node generate_reports.js --type weekly

# 月次レポート
NODE_PATH=/opt/homebrew/lib/node_modules node generate_reports.js --type monthly --month 2026-04

# 特定アカウントのみ
NODE_PATH=/opt/homebrew/lib/node_modules node generate_reports.js --type weekly --account salon_ayano_diary
```

---

## スクレイピング

```bash
# 全アカウント一括（日次）
bash scripts/daily-scrape.sh

# 特定アカウントのみ
cd /path/to/poc_threads_scraping_ver3
make run-scraping USERS="username" MAX=100

# cron登録（毎日午前3時）
crontab -e
# → 追加: 0 3 * * * /Users/yuya/Desktop/threads-auto/scripts/daily-scrape.sh
```

---

## ファイル構成

```
threads-auto/
├── accounts.json              ← ★ チームが編集するファイル
├── generate_reports.js        ← 全ドキュメント生成エントリポイント
├── scripts/
│   └── daily-scrape.sh        ← 全アカウント一括スクレイピング
└── lib/
    ├── docx-base.js           ← 共通デザイン部品
    ├── data-loader.js         ← CSVデータ読み込み・集計
    ├── analyzer.js            ← 課題検出・質問自動生成
    ├── report-hearing.js      ← ヒアリングシート生成
    ├── report-strategy.js     ← 戦略書生成
    ├── report-weekly.js       ← 週次レポート生成
    └── report-monthly.js      ← 月次レポート生成
```

---

## 月次業務カレンダー

| タイミング | 作業 | 担当 |
|---|---|---|
| 毎日 03:00（自動） | 全アカウントスクレイピング | cron |
| 毎週月曜 | 週次レポート生成・共有 | ディレクター |
| 毎月1日 | 月次レポート生成・クライアントMTG | ディレクター |
| MTG前 | ヒアリングシート・戦略書生成 | ディレクター |
| 新規アカウント追加時 | accounts.json 編集 → PR | チームメンバー |

---

## よくあるエラー

| エラー | 対処 |
|---|---|
| `Cannot find module 'docx'` | `npm install -g docx` を実行 |
| `Failed to create SingletonLock` | `rm -f browser_session/SingletonLock` |
| `アカウントが見つかりません` | accounts.jsonのidを確認 |
| CSVが空・0投稿 | `make run-scraping USERS=xxx MAX=100` を実行 |
