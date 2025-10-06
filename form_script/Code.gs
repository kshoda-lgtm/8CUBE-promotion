/**
 * 統合ナレッジ共有基盤 - Form to Markdown (v4.0)
 * PowerPoint自動抽出 & 手動入力の両方に対応
 *
 * 【セットアップ手順】
 * 1. Google Formを作成（統合Google_Form設計.md参照）
 * 2. このスクリプトをフォームのApps Scriptエディタにコピー
 * 3. スクリプトプロパティを設定:
 *    - GEMINI_API_KEY: Gemini APIキー
 *    - OUTPUT_FOLDER_ID: Markdown出力先フォルダID
 * 4. トリガーを設定: onFormSubmit関数をフォーム送信時に実行
 */

// ========================================
// 設定
// ========================================

const CONFIG = {
  // スクリプトプロパティから取得
  get GEMINI_API_KEY() {
    return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  },
  get OUTPUT_FOLDER_ID() {
    return PropertiesService.getScriptProperties().getProperty('OUTPUT_FOLDER_ID');
  },

  // Gemini API設定
  GEMINI_ENDPOINT: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent',

  // フォーム質問のインデックス（0始まり）
  FORM_INDEXES: {
    INPUT_METHOD: 0,           // データ入力方法
    // PowerPoint自動抽出
    POWERPOINT_LINK: 1,        // PowerPointリンク
    POWERPOINT_SUPPLEMENT: 2,  // 補足情報
    // 手動入力
    MANUAL_CLIENT_NAME: 3,     // クライアント名
    MANUAL_EVENT_TYPE: 4,      // イベント種別
    MANUAL_EVENT_DATE: 5,      // 実施時期
    MANUAL_EVENT_DESC: 6,      // イベント内容
    MANUAL_VENUE: 7,           // 会場
    MANUAL_TARGET_COUNT: 8,    // ターゲット人数
    MANUAL_UNIT_PRICE: 9,      // 単価
    MANUAL_TOTAL_COST: 10,     // 総費用
    MANUAL_ORDER_QTY: 11,      // 発注数量
    MANUAL_DEADLINE: 12,       // 納期
    MANUAL_PARTNERS: 13,       // 協力会社
    MANUAL_NOVELTY: 14,        // ノベルティ
    MANUAL_KEYWORDS: 15,       // キーワード
    EMAIL: 16                  // 通知先メール
  }
};

// ========================================
// メイン処理
// ========================================

/**
 * フォーム送信時のトリガー関数
 */
function onFormSubmit(e) {
  try {
    Logger.log('📝 Form submitted - Starting processing...');

    // フォーム回答を取得
    const responses = e.response.getItemResponses();
    const inputMethod = responses[CONFIG.FORM_INDEXES.INPUT_METHOD].getResponse();
    const email = responses[CONFIG.FORM_INDEXES.EMAIL].getResponse();

    Logger.log(`Input method: ${inputMethod}`);

    let analysisResult;
    let fileName = 'Unknown';

    if (inputMethod === "PowerPointから自動抽出") {
      // PowerPoint自動抽出処理
      analysisResult = processPowerPoint(responses);
      fileName = analysisResult.file_name || 'PowerPoint';
    } else {
      // 手動入力処理
      analysisResult = processManualInput(responses);
      fileName = analysisResult.client_name || 'Manual';
    }

    // Markdown生成
    const markdown = generateMarkdown(analysisResult, inputMethod);

    // Google Driveに保存
    const file = saveMarkdownToDrive(markdown, analysisResult.client_name || fileName);

    // メール通知
    sendNotificationEmail(email, file, analysisResult);

    Logger.log('✅ Processing completed successfully');

  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    Logger.log(error.stack);

    // エラー時もメール通知
    try {
      const email = e.response.getItemResponses()[CONFIG.FORM_INDEXES.EMAIL].getResponse();
      sendErrorEmail(email, error.message);
    } catch (e2) {
      Logger.log('Failed to send error email');
    }
  }
}

// ========================================
// PowerPoint自動抽出処理
// ========================================

/**
 * PowerPointファイルを処理
 */
function processPowerPoint(responses) {
  Logger.log('🔍 Processing PowerPoint...');

  // PowerPointリンクから File ID を抽出
  const driveLink = responses[CONFIG.FORM_INDEXES.POWERPOINT_LINK].getResponse();
  const fileId = extractFileIdFromLink(driveLink);

  if (!fileId) {
    throw new Error('Invalid Google Drive link. Please check the URL.');
  }

  // ファイル情報を取得
  const file = DriveApp.getFileById(fileId);
  const fileName = file.getName();
  const mimeType = file.getMimeType();

  Logger.log(`File: ${fileName} (${mimeType})`);

  // PowerPointまたはGoogle Slidesかチェック
  if (mimeType !== MimeType.GOOGLE_SLIDES &&
      mimeType !== 'application/vnd.openxmlformats-officedocument.presentationml.presentation' &&
      mimeType !== 'application/vnd.ms-powerpoint') {
    throw new Error('File is not a PowerPoint or Google Slides presentation.');
  }

  // Google Slidesに変換（必要な場合）
  let presentationId = fileId;
  if (mimeType !== MimeType.GOOGLE_SLIDES) {
    Logger.log('Converting PowerPoint to Google Slides...');
    // PowerPointファイルの場合は、既にDriveにアップロードされているものを
    // Slides APIで開くことはできないため、ユーザーに事前変換を依頼
    throw new Error('PowerPoint形式(.pptx)は未対応です。Google Slidesに変換してからリンクを貼り付けてください。');
  }

  // テキスト抽出
  const slideTexts = extractTextFromSlides(presentationId);

  if (slideTexts.length === 0) {
    throw new Error('No text found in the presentation.');
  }

  Logger.log(`Extracted ${slideTexts.length} text blocks`);

  // 補足情報
  const supplement = responses[CONFIG.FORM_INDEXES.POWERPOINT_SUPPLEMENT].getResponse() || '';

  // Gemini APIで解析
  const analysisResult = analyzeWithGemini(slideTexts, fileName, supplement);
  analysisResult.file_name = fileName;
  analysisResult.source = 'powerpoint';

  return analysisResult;
}

/**
 * Google DriveリンクからFile IDを抽出
 */
function extractFileIdFromLink(link) {
  // https://drive.google.com/file/d/FILE_ID/view?usp=sharing
  const match = link.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match) {
    return match[1];
  }

  // https://drive.google.com/open?id=FILE_ID
  const match2 = link.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match2) {
    return match2[1];
  }

  return null;
}

/**
 * Google SlidesからテキストをすべてText抽出
 */
function extractTextFromSlides(presentationId) {
  const presentation = SlidesApp.openById(presentationId);
  const slides = presentation.getSlides();
  const allTexts = [];

  slides.forEach(function(slide, index) {
    Logger.log(`Processing slide ${index + 1}/${slides.length}`);

    // 図形からテキストを抽出
    const shapes = slide.getShapes();
    shapes.forEach(function(shape) {
      try {
        const text = shape.getText().asString().trim();
        if (text) {
          allTexts.push(text);
        }
      } catch (e) {
        // テキストがない図形はスキップ
      }
    });

    // テーブルからテキストを抽出
    const tables = slide.getTables();
    tables.forEach(function(table) {
      const numRows = table.getNumRows();
      const numCols = table.getNumColumns();

      for (let r = 0; r < numRows; r++) {
        const rowTexts = [];
        for (let c = 0; c < numCols; c++) {
          const cell = table.getCell(r, c);
          const cellText = cell.getText().asString().trim();
          if (cellText) {
            rowTexts.push(cellText);
          }
        }
        if (rowTexts.length > 0) {
          allTexts.push(rowTexts.join(' | '));
        }
      }
    });
  });

  return allTexts;
}

/**
 * Gemini APIでテキストを解析
 */
function analyzeWithGemini(slideTexts, fileName, supplement) {
  Logger.log('🤖 Calling Gemini API...');

  const apiKey = CONFIG.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is not set in Script Properties.');
  }

  // プロンプト作成
  const combinedText = slideTexts.join('\n\n');
  const prompt = `あなたはプロモーション事業のデータ分析AIです。
以下のPowerPointスライドのテキストから、構造化データを抽出してください。

【ファイル名】
${fileName}

【抽出項目】
1. client_name: クライアント名（【XX様】などから企業名を抽出。「様」「株式会社」「有限会社」は除く）
2. event_date: 実施時期（YYYY/MM/DD形式で。複数ある場合は最も重要なもの）
3. event_type: イベント種別（提案書/運営マニュアル/進行台本/企画書/キャンペーン/イベント/展示会/セミナーなど）
4. event_description: イベント内容の概要（1-2文で）
5. unit_price: 単価（円、数値のみ。複数ある場合は代表的なもの）
6. total_cost: 総費用（円、数値のみ）
7. order_quantity: 発注数量（数値のみ）
8. target_count: ターゲット人数（「先着XX名」などから）
9. deadline: 納期（「XX営業日」「YYYY年MM月」など、元の表現を保持）
10. partner_companies: 協力会社名のリスト（最大5社）
11. novelty_items: ノベルティ/景品の具体的な名称リスト（最大5個）
12. venue: 会場名
13. keywords: 重要なキーワードリスト（最大10個）

${supplement ? `【補足情報】\n${supplement}\n` : ''}

【スライドテキスト】
${combinedText.substring(0, 5000)}

【出力形式】
以下のJSON形式で出力してください。値が不明な場合はnullを設定してください。
{
  "client_name": "クライアント名",
  "event_date": "2024/01/01",
  "event_type": "種別",
  "event_description": "概要",
  "unit_price": 500,
  "total_cost": 300000,
  "order_quantity": 1000,
  "target_count": 500,
  "deadline": "14営業日",
  "partner_companies": ["会社1", "会社2"],
  "novelty_items": ["景品1", "景品2"],
  "venue": "会場名",
  "keywords": ["キーワード1", "キーワード2"]
}

重要: 必ずJSON形式のみを出力してください。説明文は不要です。`;

  const payload = {
    "contents": [{
      "parts": [{
        "text": prompt
      }]
    }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(CONFIG.GEMINI_ENDPOINT + '?key=' + apiKey, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error(`Gemini API error: ${response.getContentText()}`);
  }

  const json = JSON.parse(response.getContentText());

  // レスポンスからJSONを抽出
  const text = json.candidates[0].content.parts[0].text;
  Logger.log(`Gemini response: ${text.substring(0, 200)}...`);

  // JSONブロックを抽出
  const jsonMatch = text.match(/```json\s*([\s\S]*?)\s*```/);
  const jsonStr = jsonMatch ? jsonMatch[1] : text;

  try {
    const result = JSON.parse(jsonStr);
    result.confidence_score = calculateConfidence(result);
    Logger.log(`Confidence score: ${result.confidence_score}%`);
    return result;
  } catch (e) {
    Logger.log(`Failed to parse JSON: ${jsonStr}`);
    throw new Error('Gemini API returned invalid JSON');
  }
}

/**
 * 信頼度スコアを計算
 */
function calculateConfidence(data) {
  let score = 0;

  if (data.client_name) score += 15;
  if (data.event_date) score += 15;
  if (data.event_type) score += 10;
  if (data.event_description) score += 10;
  if (data.unit_price) score += 10;
  if (data.total_cost) score += 10;
  if (data.order_quantity) score += 5;
  if (data.deadline) score += 5;
  if (data.partner_companies && data.partner_companies.length > 0) score += 10;
  if (data.novelty_items && data.novelty_items.length > 0) score += 5;
  if (data.keywords && data.keywords.length > 0) score += 5;

  return Math.min(score, 100);
}

// ========================================
// 手動入力処理
// ========================================

/**
 * 手動入力フォームから取得を処理
 */
function processManualInput(responses) {
  Logger.log('✍️ Processing manual input...');

  const result = {
    client_name: responses[CONFIG.FORM_INDEXES.MANUAL_CLIENT_NAME].getResponse(),
    event_type: responses[CONFIG.FORM_INDEXES.MANUAL_EVENT_TYPE].getResponse(),
    event_date: getOptionalResponse(responses, CONFIG.FORM_INDEXES.MANUAL_EVENT_DATE),
    event_description: getOptionalResponse(responses, CONFIG.FORM_INDEXES.MANUAL_EVENT_DESC),
    venue: getOptionalResponse(responses, CONFIG.FORM_INDEXES.MANUAL_VENUE),
    target_count: parseNumber(getOptionalResponse(responses, CONFIG.FORM_INDEXES.MANUAL_TARGET_COUNT)),
    unit_price: parseNumber(getOptionalResponse(responses, CONFIG.FORM_INDEXES.MANUAL_UNIT_PRICE)),
    total_cost: parseNumber(getOptionalResponse(responses, CONFIG.FORM_INDEXES.MANUAL_TOTAL_COST)),
    order_quantity: parseNumber(getOptionalResponse(responses, CONFIG.FORM_INDEXES.MANUAL_ORDER_QTY)),
    deadline: getOptionalResponse(responses, CONFIG.FORM_INDEXES.MANUAL_DEADLINE),
    partner_companies: parseList(getOptionalResponse(responses, CONFIG.FORM_INDEXES.MANUAL_PARTNERS)),
    novelty_items: parseList(getOptionalResponse(responses, CONFIG.FORM_INDEXES.MANUAL_NOVELTY)),
    keywords: parseList(getOptionalResponse(responses, CONFIG.FORM_INDEXES.MANUAL_KEYWORDS)),
    confidence_score: 100, // 手動入力は100%
    source: 'manual'
  };

  return result;
}

/**
 * オプション回答を取得（存在しない場合はnull）
 */
function getOptionalResponse(responses, index) {
  try {
    const response = responses[index].getResponse();
    return response ? response.trim() : null;
  } catch (e) {
    return null;
  }
}

/**
 * 数値をパース
 */
function parseNumber(str) {
  if (!str) return null;
  const num = parseInt(str.replace(/,/g, ''), 10);
  return isNaN(num) ? null : num;
}

/**
 * カンマ区切り文字列をリストに変換
 */
function parseList(str) {
  if (!str) return [];
  return str.split(',').map(s => s.trim()).filter(s => s.length > 0);
}

// ========================================
// Markdown生成
// ========================================

/**
 * Markdownファイルを生成
 */
function generateMarkdown(data, inputMethod) {
  Logger.log('📄 Generating Markdown...');

  const md = [];

  // ヘッダー
  const title = data.client_name ? `【${data.client_name}様】` : '';
  const eventType = data.event_type || 'プロモーション案件';
  md.push(`# ${title}${eventType}\n`);
  md.push(`**処理日時**: ${new Date().toISOString()}`);
  md.push(`**データソース**: ${inputMethod}`);
  md.push(`**信頼度スコア**: ${data.confidence_score || 0}%\n`);
  md.push('---\n');

  // 基本情報
  md.push('## 📋 基本情報\n');
  if (data.client_name) md.push(`- **クライアント名**: ${data.client_name}`);
  if (data.event_date) md.push(`- **実施時期**: ${data.event_date}`);
  if (data.event_type) md.push(`- **イベント種別**: ${data.event_type}`);
  if (data.venue) md.push(`- **会場**: ${data.venue}`);
  if (data.target_count) md.push(`- **ターゲット人数**: ${data.target_count}名`);
  md.push('');

  // イベント内容
  if (data.event_description) {
    md.push('## 📝 イベント内容\n');
    md.push(data.event_description + '\n');
  }

  // 価格情報
  if (data.unit_price || data.total_cost || data.order_quantity) {
    md.push('## 💰 価格情報\n');
    if (data.unit_price) md.push(`- **単価**: ¥${data.unit_price.toLocaleString()}`);
    if (data.total_cost) md.push(`- **総費用**: ¥${data.total_cost.toLocaleString()}`);
    if (data.order_quantity) md.push(`- **発注数量**: ${data.order_quantity.toLocaleString()}個`);
    md.push('');
  }

  // 納期
  if (data.deadline) {
    md.push('## ⏰ 納期\n');
    md.push(`- **納期**: ${data.deadline}\n`);
  }

  // 協力会社
  if (data.partner_companies && data.partner_companies.length > 0) {
    md.push('## 🤝 協力会社\n');
    data.partner_companies.forEach(company => md.push(`- ${company}`));
    md.push('');
  }

  // ノベルティ
  if (data.novelty_items && data.novelty_items.length > 0) {
    md.push('## 🎁 ノベルティ/景品\n');
    data.novelty_items.forEach(item => md.push(`- ${item}`));
    md.push('');
  }

  // キーワード
  if (data.keywords && data.keywords.length > 0) {
    md.push('## 🏷️ タグ・キーワード\n');
    const tags = data.keywords.map(kw => `\`#${kw}\``).join(' ');
    md.push(tags + '\n');
  }

  // フッター
  md.push('---');
  md.push(`\n*Generated by NotebookLM Knowledge System v4.0 - ${new Date().toISOString()}*`);

  return md.join('\n');
}

// ========================================
// Google Drive保存
// ========================================

/**
 * Google DriveにMarkdownファイルを保存
 */
function saveMarkdownToDrive(markdown, clientName) {
  Logger.log('💾 Saving to Google Drive...');

  const folderId = CONFIG.OUTPUT_FOLDER_ID;
  if (!folderId) {
    throw new Error('OUTPUT_FOLDER_ID is not set in Script Properties.');
  }

  const folder = DriveApp.getFolderById(folderId);
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmmss');
  const fileName = `${timestamp}_${clientName || 'Unknown'}.md`;

  const file = folder.createFile(fileName, markdown, MimeType.PLAIN_TEXT);

  Logger.log(`File saved: ${fileName}`);
  return file;
}

// ========================================
// メール通知
// ========================================

/**
 * 処理完了メールを送信
 */
function sendNotificationEmail(email, file, data) {
  Logger.log(`📧 Sending email to ${email}...`);

  const subject = '【ナレッジ共有基盤】Markdown生成完了';
  const body = `
Markdownファイルの生成が完了しました。

━━━━━━━━━━━━━━━━━━━━━━━━
📄 生成されたファイル
━━━━━━━━━━━━━━━━━━━━━━━━
ファイル名: ${file.getName()}
クライアント名: ${data.client_name || 'N/A'}
イベント種別: ${data.event_type || 'N/A'}
信頼度スコア: ${data.confidence_score || 0}%

ダウンロード: ${file.getUrl()}

━━━━━━━━━━━━━━━━━━━━━━━━
📝 次のステップ
━━━━━━━━━━━━━━━━━━━━━━━━
1. 上記リンクからMarkdownファイルをダウンロード
2. NotebookLMにアクセス: https://notebooklm.google.com/
3. Markdownファイルをアップロード（ドラッグ&ドロップ）

これで完了です！

━━━━━━━━━━━━━━━━━━━━━━━━
💡 NotebookLMの活用例
━━━━━━━━━━━━━━━━━━━━━━━━
- 「${data.client_name || 'XX'}様に似た案件はありますか？」
- 「過去の${data.event_type || 'イベント'}の単価を教えて」
- 「${data.keywords && data.keywords[0] ? data.keywords[0] : ''}の案件を教えて」

---
株式会社エイトキューブ プロモーション事業部
ナレッジ共有基盤システム v4.0
`;

  MailApp.sendEmail(email, subject, body);
  Logger.log('✅ Email sent');
}

/**
 * エラーメールを送信
 */
function sendErrorEmail(email, errorMessage) {
  const subject = '【ナレッジ共有基盤】エラーが発生しました';
  const body = `
Markdownファイルの生成中にエラーが発生しました。

エラー内容:
${errorMessage}

お手数ですが、以下を確認してください:
- PowerPointリンクが正しいか
- PowerPointファイルがGoogle Slidesに変換されているか
- 必須項目がすべて入力されているか

それでも解決しない場合は、システム管理者にお問い合わせください。

---
株式会社エイトキューブ プロモーション事業部
ナレッジ共有基盤システム v4.0
`;

  MailApp.sendEmail(email, subject, body);
}

// ========================================
// テスト関数
// ========================================

/**
 * Gemini API接続テスト
 */
function testGeminiAPI() {
  const testTexts = [
    '【広研様】洛北阪急スクエアイベント案',
    '実施時期: 2024年10月',
    '単価: ¥1,000',
    '総費用: ¥500,000'
  ];

  const result = analyzeWithGemini(testTexts, 'test.pptx', '');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * Markdown生成テスト
 */
function testMarkdownGeneration() {
  const testData = {
    client_name: '広研',
    event_type: '提案書',
    event_date: '2024/10/02',
    event_description: '洛北阪急スクエアでのプロモーションイベント',
    unit_price: 1000,
    total_cost: 500000,
    keywords: ['競馬', 'ファミリー向け'],
    confidence_score: 85
  };

  const markdown = generateMarkdown(testData, 'PowerPointから自動抽出');
  Logger.log(markdown);
}
