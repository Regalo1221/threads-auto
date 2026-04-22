# Threads 運用代行 自動化システム

Threadsアカウントのスクレイピング・分析・レポート生成・ヒアリングシート・戦略書を自動生成するシステムです。

---

## 全体フロー

```
① accounts.json にアカウントを追加
      ↓
② input/[id]/hearing.md を記入（初回ヒアリング内容）
   input/[id]/urls.txt  を記入（HP・SNS URL）
      ↓
③ node onboard.js --account [id]
      ↓
  ┌────────────────────────────────────────────┐
  │  スクレイピング → データ分析 → 課題検出    │
  │  → ヒアリングシートDOCX（追加質問付き）   │
  │  → 戦略書ひな形DOCX                        │
  └────────────────────────────────────────────┘
      ↓
④ 毎日 daily-scrape.sh が自動実行（数値を継続蓄積）
      ↓
⑤ 週次・月次レポートを generate_reports.js で生成
      ↓
  ┌──────────────────────────────────┐
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

## 新規アカウントのオンボーディング

### 概要

アカウントIDと初回ヒアリング情報を用意するだけで、スクレイピング・分析・ヒアリングシート・戦略書を自動生成します。

### 手順

**Step 1. `accounts.json` に最低限の情報を追加する**

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
> 最低限必要なフィールド: `id`, `handle`, `scrape_username`, `csv_dir`
> 残りは hearing.md から自動補完されます。

**Step 2. ヒアリングシートテンプレートを記入する**

```bash
# テンプレートをコピー
cp templates/hearing-input.md input/[id]/hearing.md

# HP・SNS URLを記入（任意）
cp templates/urls.txt input/[id]/urls.txt
```

`input/[id]/hearing.md` を開いて、初回MTGで聞いた内容を記入します。

**Step 3. オンボーディングを実行する**

```bash
node onboard.js --account [id]
```

これだけで以下が自動実行されます：
- hearing.md の内容を accounts.json に反映
- Threads スクレイピング（全件取得）
- データ分析・課題検出
- ヒアリングシート DOCX 生成（データから追加質問を自動生成）
- 戦略書ひな形 DOCX 生成

出力先: `output/[id]/`

**Step 4. PRを作成してレビューを依頼する**

```bash
git add accounts.json input/[id]/
git commit -m "add account: @アカウント名"
git push origin [ブランチ名]
```

---

## コマンドリファレンス

### オンボーディング（新規アカウント）

```bash
# 通常実行（スクレイピング → 分析 → DOCX生成）
node onboard.js --account [id]

# スクレイピングをスキップ（再生成のみ）
node onboard.js --account [id] --skip-scrape
```

### 継続レポート生成

```bash
# ヒアリングシート
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
