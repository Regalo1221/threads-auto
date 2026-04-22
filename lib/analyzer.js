/**
 * analyzer.js
 * スクレイピングデータ + アカウント設定から
 * 「課題」「不足情報」「ヒアリング質問」を自動生成する
 */
const { loadCsv, findLatestCsv, computeKPIs } = require('./data-loader');

// ─────────────────────────────────────────────
// メイン分析関数
// ─────────────────────────────────────────────
function analyzeAccount(account) {
  const csvPath = findLatestCsv(account.csv_dir, account.scrape_username);
  const posts   = csvPath ? loadCsv(csvPath) : [];

  const result = {
    account,
    csvPath,
    posts,
    hasData:   posts.length > 0,
    kpi:       computeKPIs(posts),
    issues:    [],      // {severity, category, title, detail, evidence}
    questions: [],      // {section, priority, question, why, hint}
    missingInfo: [],    // accounts.json に入っていない設定値
  };

  // データなしの場合は即返却
  if (!result.hasData) {
    result.issues.push({
      severity: 'CRITICAL',
      category: 'データ',
      title:    'スクレイピングデータなし',
      detail:   'CSVファイルが存在しないか、投稿が0件です。スクレイピングを実行してください。',
      evidence: null,
    });
    result.questions.push({
      section: 'データ取得',
      priority: 'HIGH',
      question: 'Threadsのアカウント名（@以降）を正確に教えてください。',
      why: 'スクレイピングが失敗している可能性があります。',
      hint: '',
    });
    return result;
  }

  // ── 分析モジュールを順番に実行 ────────────────
  checkPostingConsistency(result);
  checkEngagementQuality(result);
  checkContentMix(result);
  checkGoalAlignment(result);
  checkMissingAccountInfo(result);
  buildUniversalQuestions(result);
  buildDataDrivenQuestions(result);

  // 優先度でソート
  const order = { CRITICAL:0, HIGH:1, MED:2, LOW:3 };
  result.issues.sort((a,b) => order[a.severity] - order[b.severity]);
  result.questions.sort((a,b) => order[a.priority] - order[b.priority]);

  return result;
}

// ─────────────────────────────────────────────
// 投稿一貫性チェック
// ─────────────────────────────────────────────
function checkPostingConsistency(result) {
  const { posts, kpi, account } = result;
  if (posts.length === 0) return;

  // 週別投稿数
  const byWeek = {};
  posts.forEach(p => {
    if (!p.date) return;
    const w = weekKey(p.date);
    byWeek[w] = (byWeek[w] || 0) + 1;
  });
  const weekCounts = Object.values(byWeek);
  const avgWeekPosts = weekCounts.reduce((s,v)=>s+v,0) / (weekCounts.length || 1);
  const target = account.weekly_post_count || 21;

  // 投稿数が少ない
  if (avgWeekPosts < target * 0.7) {
    result.issues.push({
      severity: 'HIGH',
      category: '投稿頻度',
      title:    `平均週${avgWeekPosts.toFixed(1)}件（目標${target}件の${Math.round(avgWeekPosts/target*100)}%）`,
      detail:   '投稿頻度が目標を大きく下回っています。アルゴリズムリーチの低下が懸念されます。',
      evidence: `最多週: ${Math.max(...weekCounts)}件 / 最少週: ${Math.min(...weekCounts)}件`,
    });
  }

  // 週によってバラつきが大きい
  if (weekCounts.length >= 3) {
    const max = Math.max(...weekCounts);
    const min = Math.min(...weekCounts);
    if (max > min * 3) {
      result.issues.push({
        severity: 'MED',
        category: '投稿頻度',
        title:    `週ごとの投稿数にバラつきあり（${min}〜${max}件）`,
        detail:   '投稿スケジュールが不安定です。特定の週に集中し、他の週が薄い状態になっています。',
        evidence: `週別投稿数: ${JSON.stringify(byWeek)}`,
      });
    }
  }
}

// ─────────────────────────────────────────────
// エンゲージメント品質チェック
// ─────────────────────────────────────────────
function checkEngagementQuality(result) {
  const { kpi, posts } = result;

  // 0いいね率が高い
  if (kpi.zeroLikeRate > 70) {
    result.issues.push({
      severity: 'HIGH',
      category: 'エンゲージメント',
      title:    `0いいね投稿率 ${kpi.zeroLikeRate}%`,
      detail:   '過半数の投稿がほぼ反応を得られていません。コンテンツの方向性かターゲット設定に問題がある可能性があります。',
      evidence: `全${kpi.total}投稿中、いいね0件が${Math.round(kpi.total * kpi.zeroLikeRate / 100)}件`,
    });
  }

  // コメントが極端に少ない
  if (kpi.total > 10 && kpi.comments < kpi.total * 0.05) {
    result.issues.push({
      severity: 'MED',
      category: 'エンゲージメント',
      title:    `コメント誘発の仕掛けがほぼない（平均${(kpi.comments/kpi.total).toFixed(2)}件/投稿）`,
      detail:   'コメントはスコア×3の重みを持つ重要指標です。問いかけ型の投稿が不足しています。',
      evidence: `総コメント数: ${kpi.comments}件 / 総投稿数: ${kpi.total}件`,
    });
  }

  // バズ依存（上位2件が全体の80%以上を占める）
  if (kpi.top3.length >= 2) {
    const top2Score = kpi.top3[0].score + kpi.top3[1].score;
    if (kpi.score > 0 && top2Score / kpi.score > 0.8) {
      result.issues.push({
        severity: 'HIGH',
        category: 'エンゲージメント',
        title:    `バズ依存構造（上位2投稿がスコアの${Math.round(top2Score/kpi.score*100)}%を占める）`,
        detail:   '一部の投稿に数値が集中し、他はほぼ反応なし。安定した運用ができていない状態です。',
        evidence: `TOP2スコア計: ${top2Score}pt / 全体: ${kpi.score}pt`,
      });
    }
  }
}

// ─────────────────────────────────────────────
// コンテンツミックスチェック
// ─────────────────────────────────────────────
function checkContentMix(result) {
  const { kpi, posts } = result;
  const types = kpi.byType;
  const total = kpi.total;
  if (total === 0) return;

  // インサイト祭りが少ない（目標集客施策）
  const insightRate = (types['インサイト祭り']?.count || 0) / total;
  if (insightRate < 0.1) {
    result.issues.push({
      severity: 'HIGH',
      category: 'コンテンツ',
      title:    `集客投稿（インサイト祭り）がほぼない（全体の${Math.round(insightRate*100)}%）`,
      detail:   'フォロワー獲得の核となるインサイト祭り投稿が少なすぎます。目標の週3本（全体の14%以上）を確保してください。',
      evidence: `インサイト祭り: ${types['インサイト祭り']?.count || 0}件 / 全体: ${total}件`,
    });
  }

  // 雑談比率が高すぎる
  const casualRate = (types['雑談・日常']?.count || 0) / total;
  if (casualRate > 0.5) {
    result.issues.push({
      severity: 'MED',
      category: 'コンテンツ',
      title:    `雑談・日常投稿が過半数（${Math.round(casualRate*100)}%）を占めている`,
      detail:   'バズは取りやすいが、ビジネスゴールに直結しないコンテンツが主体になっています。雑談でバズっても面談・集客にはつながりません。',
      evidence: `雑談: ${types['雑談・日常']?.count || 0}件 / 全体: ${total}件`,
    });
  }

  // サロン経営Tipsが少ない（専門性の欠如）
  const tipsRate = (types['サロン経営Tips']?.count || 0) / total;
  if (tipsRate < 0.15 && total > 10) {
    result.issues.push({
      severity: 'MED',
      category: 'コンテンツ',
      title:    `専門性・ノウハウ投稿が少ない（全体の${Math.round(tipsRate*100)}%）`,
      detail:   '「この人の話には価値がある」という信頼構築に必要な専門性コンテンツが不足しています。',
      evidence: `Tips投稿: ${types['サロン経営Tips']?.count || 0}件 / 全体: ${total}件`,
    });
  }
}

// ─────────────────────────────────────────────
// ゴール整合性チェック
// ─────────────────────────────────────────────
function checkGoalAlignment(result) {
  const { account, kpi } = result;

  // KPI目標が未設定
  if (!account.kpi_target || account.kpi_target === 0) {
    result.issues.push({
      severity: 'HIGH',
      category: '目標設定',
      title:    'KPI目標数値が未設定',
      detail:   '何件のDM・面談・成約を目標にするか定義されていません。測定できない目標は達成できません。',
      evidence: `accounts.json: kpi_target = ${account.kpi_target || '未設定'}`,
    });
  }

  // ターゲット未設定
  if (!account.target_audience || account.target_audience === 'ターゲット層を入力') {
    result.issues.push({
      severity: 'CRITICAL',
      category: '目標設定',
      title:    'ターゲット顧客が未定義',
      detail:   '「誰に向けて発信するか」が定義されていません。全ての投稿が誰にも刺さらない状態です。',
      evidence: `accounts.json: target_audience = 未設定`,
    });
  }

  // コンテンツ4本柱が未設定
  if (!account.content_pillars || account.content_pillars[0] === '柱①') {
    result.issues.push({
      severity: 'HIGH',
      category: '戦略',
      title:    'コンテンツ4本柱が未設計',
      detail:   '投稿の方向性が定まっていないため、何を投稿するかが毎回の判断に委ねられています。',
      evidence: `accounts.json: content_pillars = 未設定`,
    });
  }
}

// ─────────────────────────────────────────────
// accounts.json の不足情報チェック
// ─────────────────────────────────────────────
function checkMissingAccountInfo(result) {
  const { account } = result;
  const required = [
    { key: 'goal',            label: '目標・ゴール' },
    { key: 'target_audience', label: 'ターゲット層' },
    { key: 'business_type',   label: '業種・サービス内容' },
    { key: 'kpi_label',       label: 'KPI指標名' },
    { key: 'kpi_target',      label: 'KPI目標数値' },
    { key: 'location',        label: '所在地' },
  ];
  required.forEach(({ key, label }) => {
    const val = account[key];
    if (!val || val === '業種を入力' || val === '目標を入力' || val === 'ターゲット層を入力' || val === 0) {
      result.missingInfo.push({ key, label });
    }
  });
}

// ─────────────────────────────────────────────
// ユニバーサルヒアリング質問（全アカウント共通）
// ─────────────────────────────────────────────
function buildUniversalQuestions(result) {
  const qs = result.questions;

  // A. ビジネス基本情報
  qs.push(...[
    { section:'A. ビジネス基本情報', priority:'HIGH', question:'提供しているサービス・商品の内容を詳しく教えてください。', why:'投稿コンテンツの専門性を設計するために必要です。', hint:'例：フェイシャルエステ、小顔矯正、まつエクなど' },
    { section:'A. ビジネス基本情報', priority:'HIGH', question:'客単価と月商（またはその目標）を教えてください。', why:'KPI設計と費用対効果の算出に必要です。', hint:'例：1回 8,000円、月商 30〜40万円' },
    { section:'A. ビジネス基本情報', priority:'MED',  question:'今の顧客層（年齢・性別・職業など）を教えてください。', why:'コンテンツの言葉遣いとターゲット設定に必要です。', hint:'' },
    { section:'A. ビジネス基本情報', priority:'MED',  question:'競合・同業と比べた際の一番の強みは何ですか？', why:'「この人だから」という理由を投稿で訴求するために必要です。', hint:'例：価格、技術力、アフターフォロー、アクセス' },
  ]);

  // B. Threads目標
  qs.push(...[
    { section:'B. Threadsの目標設定', priority:'HIGH', question:'Threadsで最終的に何を達成したいですか？', why:'全ての投稿方針の根幹になります。', hint:'例：DM相談月10件、面談月3件、商品販売月5件' },
    { section:'B. Threadsの目標設定', priority:'HIGH', question:'フォロワーに対してどんなアクションをしてほしいですか？（理想の動線）', why:'CTAの設計に必要です。', hint:'例：インスタDM → 面談 → 成約' },
    { section:'B. Threadsの目標設定', priority:'MED',  question:'「このアカウントが成功した」と判断できる具体的な数字は何ですか？', why:'KPI目標の設定に使います。', hint:'例：フォロワー1,000名、月DM5件' },
  ]);

  // C. ターゲット顧客
  qs.push(...[
    { section:'C. ターゲット顧客', priority:'HIGH', question:'一番来てほしい・繋がりたい人を1人具体的に描いてみてください。', why:'ターゲットペルソナの設計に必要です。', hint:'例：30代、個人サロン経営、月商20万、リピート率に悩んでいる' },
    { section:'C. ターゲット顧客', priority:'HIGH', question:'ターゲットが一番困っていること・悩んでいることは何ですか？', why:'共感コンテンツと問題提起型投稿の設計に使います。', hint:'' },
    { section:'C. ターゲット顧客', priority:'MED',  question:'ターゲットが「これを知りたかった」と思う情報はどんなものですか？', why:'価値提供型コンテンツの設計に使います。', hint:'' },
  ]);

  // D. 実績・強み
  qs.push(...[
    { section:'D. 実績・強み', priority:'HIGH', question:'今まででいちばん印象に残っている成功体験・成果を教えてください（数字があれば）。', why:'「実績と信頼性」の投稿ネタになります。', hint:'例：リピート率を3ヶ月で30%→70%に改善' },
    { section:'D. 実績・強み', priority:'MED',  question:'お客様から言われることが多い褒め言葉・感謝の声はどんなものですか？', why:'「お客様の声」型コンテンツに使います。', hint:'' },
    { section:'D. 実績・強み', priority:'MED',  question:'Threadsでは話したくない・出したくない情報はありますか？', why:'コンテンツのNG範囲を事前に確定するために必要です。', hint:'例：価格、個人情報、特定の競合への言及' },
  ]);

  // E. 投稿スタイル
  qs.push(...[
    { section:'E. 投稿スタイル・キャラクター', priority:'HIGH', question:'顔出し・本名公開の可否を教えてください。', why:'アカウントのキャラクター設計に影響します。', hint:'例：顔出しOK / イラストアイコンのみ / 完全匿名' },
    { section:'E. 投稿スタイル・キャラクター', priority:'MED',  question:'話し言葉・口調のクセや地域性はありますか？（関西弁など）', why:'投稿の「声」に人間らしさを持たせるために使います。', hint:'例：関西弁OK、丁寧語希望、語尾に🙂をよく使う' },
    { section:'E. 投稿スタイル・キャラクター', priority:'MED',  question:'「この人の投稿、好き」と思っているアカウントがあれば教えてください。', why:'目指すトーンと世界観の参考にします。', hint:'' },
  ]);

  // F. 運用体制
  qs.push(...[
    { section:'F. 運用体制', priority:'HIGH', question:'投稿は誰が書きますか？（本人 / 代行 / 一緒に作る）', why:'ワークフローと投稿ストックの作り方に影響します。', hint:'' },
    { section:'F. 運用体制', priority:'MED',  question:'コメント・DMへの返信対応はできますか？どれくらいの頻度で確認できますか？', why:'エンゲージメント運用の設計に必要です。', hint:'例：1日1回確認できる / 週2〜3回' },
    { section:'F. 運用体制', priority:'LOW',  question:'投稿前の確認・承認フローはどうしますか？', why:'運用ワークフローの設計に必要です。', hint:'例：前日に確認、修正あれば当日朝までに連絡' },
  ]);
}

// ─────────────────────────────────────────────
// データドリブンな追加質問（課題から自動生成）
// ─────────────────────────────────────────────
function buildDataDrivenQuestions(result) {
  const { issues, kpi, account } = result;
  const qs = result.questions;

  issues.forEach(issue => {
    switch(issue.category) {
      case 'エンゲージメント':
        if (issue.title.includes('バズ依存')) {
          qs.push({
            section: 'G. データから見えた課題への質問',
            priority: 'HIGH',
            question: `データ上、上位2投稿がエンゲージメントの${Math.round((kpi.top3[0]?.score+kpi.top3[1]?.score||0)/kpi.score*100)}%を占めています。普段どんな意図で投稿内容を決めていますか？`,
            why: 'バズ頼みの構造を変えるために、投稿選定の現状を理解する必要があります。',
            hint: '特に「バズった投稿」と「バズらなかった投稿」の違いをどう認識しているか',
          });
        }
        if (issue.title.includes('コメント')) {
          qs.push({
            section: 'G. データから見えた課題への質問',
            priority: 'HIGH',
            question: '投稿の最後に「コメントで教えてください」のような問いかけをしていますか？していない場合、その理由を教えてください。',
            why: 'コメント数がほぼ0件のため、誘発の仕掛けが機能していない可能性があります。',
            hint: '',
          });
        }
        break;
      case 'コンテンツ':
        if (issue.title.includes('インサイト祭り')) {
          qs.push({
            section: 'G. データから見えた課題への質問',
            priority: 'HIGH',
            question: '現在、フォロワーを能動的に増やすための投稿（インサイト祭り参加・自己紹介投稿など）はしていますか？',
            why: '集客施策投稿がほぼない状態です。フォロワー獲得の手段を確認する必要があります。',
            hint: '',
          });
        }
        if (issue.title.includes('雑談')) {
          qs.push({
            section: 'G. データから見えた課題への質問',
            priority: 'MED',
            question: '日常・雑談系の投稿が多い理由を教えてください。「ビジネスの話は難しそう」「何を書けばいいかわからない」など、投稿する上での不安はありますか？',
            why: '雑談が多いのは「専門的な話題が書けない」という課題の反映の可能性があります。',
            hint: '',
          });
        }
        break;
      case '目標設定':
        if (issue.title.includes('KPI')) {
          qs.push({
            section: 'G. データから見えた課題への質問',
            priority: 'HIGH',
            question: '今のThreads運用で「月に何件○○できれば成功」という具体的な目標数値を持っていますか？',
            why: 'KPI目標が未設定のため、成果を測定する基準がない状態です。',
            hint: '例：月DM5件、面談3件、商品販売2件',
          });
        }
        if (issue.title.includes('ターゲット')) {
          qs.push({
            section: 'G. データから見えた課題への質問',
            priority: 'CRITICAL',
            question: '今のThreadsのフォロワーに「この人に来てほしい」と思える人はいますか？フォロワーの質はどう見えていますか？',
            why: 'ターゲット顧客が未定義のため、誰に向けた発信かが定まっていません。',
            hint: '',
          });
        }
        break;
    }
  });

  // TOP投稿がビジネスと無関係の場合
  if (kpi.top3.length > 0) {
    const topNonBusiness = kpi.top3.filter(p => p.type === '雑談・日常');
    if (topNonBusiness.length === kpi.top3.length) {
      qs.push({
        section: 'G. データから見えた課題への質問',
        priority: 'HIGH',
        question: `現在のバズ投稿TOP${kpi.top3.length}がすべて日常・雑談系です。「バズること」と「ビジネスの目標」のどちらを優先すべきと考えていますか？`,
        why: 'バズ投稿がゴールとミスアラインしている構造を本人が認識しているか確認が必要です。',
        hint: 'TOP投稿例：' + kpi.top3.map(p => `「${p.content.slice(0,20)}…」(${p.likes}いいね)`).join(' / '),
      });
    }
  }
}

// ─────────────────────────────────────────────
// ユーティリティ
// ─────────────────────────────────────────────
function weekKey(date) {
  const d   = new Date(date);
  const mon = new Date(d.setDate(d.getDate() - (d.getDay() || 7) + 1));
  return `${mon.getFullYear()}-${String(mon.getMonth()+1).padStart(2,'0')}-${String(mon.getDate()).padStart(2,'0')}`;
}

// 全アカウント一括分析
function analyzeAll(accounts) {
  return accounts.map(account => analyzeAccount(account));
}

module.exports = { analyzeAccount, analyzeAll };
