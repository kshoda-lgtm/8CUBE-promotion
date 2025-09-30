/**
 * ナレッジ共有基盤システム - Google Apps Script（拡張版 v2.0）
 *
 * 新機能：
 * - OCR処理による画像テキスト抽出
 * - AI情報抽出エンジン
 * - チャットボット機能
 * - 過去資料一括処理
 */

// ===== 設定値 =====
const CONFIG = {
  // 本番DBスプレッドシートのID
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',

  // シート名
  SHEET_NAME: 'ナレッジDB',
  HISTORY_SHEET: '会話履歴',
  OCR_LOG_SHEET: 'OCR処理ログ',

  // Google Cloud Vision API Key（OCR用）
  VISION_API_KEY: 'YOUR_API_KEY_HERE',

  // OpenAI API Key（高度なAI処理用）
  OPENAI_API_KEY: 'YOUR_OPENAI_KEY_HERE',

  // 処理対象フォルダID（過去資料格納用）
  SOURCE_FOLDER_ID: 'YOUR_FOLDER_ID_HERE',

  // 処理済みフォルダID
  PROCESSED_FOLDER_ID: 'YOUR_PROCESSED_FOLDER_ID_HERE',

  // バッチ処理設定
  BATCH_SIZE: 10, // 一度に処理するファイル数
  OCR_TIMEOUT: 30000, // OCRタイムアウト（ミリ秒）
};

// ===== OCR処理関数 =====

/**
 * 過去資料を一括でOCR処理してDBに登録
 * トリガーで定期実行することを推奨
 */
function batchProcessHistoricalDocuments() {
  console.log('=== 過去資料一括処理開始 ===');

  try {
    const sourceFolder = DriveApp.getFolderById(CONFIG.SOURCE_FOLDER_ID);
    const processedFolder = DriveApp.getFolderById(CONFIG.PROCESSED_FOLDER_ID);

    // PowerPointファイルを検索
    const files = sourceFolder.searchFiles(
      'mimeType = "application/vnd.ms-powerpoint" or ' +
      'mimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation" or ' +
      'mimeType = "application/vnd.google-apps.presentation"'
    );

    let processedCount = 0;
    const startTime = new Date();

    while (files.hasNext() && processedCount < CONFIG.BATCH_SIZE) {
      const file = files.next();

      try {
        console.log(`処理中: ${file.getName()}`);

        // ファイルを処理
        const result = processFileWithOCR(file.getId());

        // 処理成功したらフォルダを移動
        if (result.success) {
          file.moveTo(processedFolder);
          logOCRProcess(file.getName(), 'SUCCESS', result.extractedData);
        } else {
          logOCRProcess(file.getName(), 'ERROR', null, result.error);
        }

        processedCount++;

        // レート制限回避のため少し待機
        Utilities.sleep(2000);

      } catch (error) {
        console.error(`ファイル処理エラー: ${file.getName()}`, error);
        logOCRProcess(file.getName(), 'ERROR', null, error.toString());
      }
    }

    const endTime = new Date();
    const processingTime = (endTime - startTime) / 1000;

    console.log(`=== 処理完了 ===`);
    console.log(`処理ファイル数: ${processedCount}`);
    console.log(`処理時間: ${processingTime}秒`);

  } catch (error) {
    console.error('バッチ処理エラー:', error);
    throw error;
  }
}

/**
 * ファイルをOCR処理して情報を抽出
 * @param {string} fileId - ファイルID
 * @returns {Object} 処理結果
 */
function processFileWithOCR(fileId) {
  try {
    // PowerPointをGoogleスライドに変換
    const presentationId = convertToGoogleSlides(fileId);

    // テキスト抽出（通常のテキスト要素）
    let extractedText = extractTextFromSlides(presentationId);

    // OCR処理（画像内のテキスト）
    const ocrText = performOCROnSlides(presentationId);
    extractedText += '\n' + ocrText;

    // AI処理で情報を抽出
    const extractedData = extractWithAI(extractedText);

    // 信頼度スコアを計算
    extractedData.confidenceScore = calculateConfidenceScore(extractedData);

    // データベースに保存
    saveExtractedData(extractedData, fileId);

    return {
      success: true,
      extractedData: extractedData
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Google Cloud Vision APIを使用してOCR処理
 * @param {string} presentationId - GoogleスライドID
 * @returns {string} OCRで抽出したテキスト
 */
function performOCROnSlides(presentationId) {
  let ocrText = '';

  try {
    const presentation = Slides.Presentations.get(presentationId);

    presentation.slides.forEach((slide, index) => {
      // スライド内の画像要素を探す
      if (slide.pageElements) {
        slide.pageElements.forEach(element => {
          if (element.image) {
            // 画像をBase64エンコード
            const imageUrl = element.image.sourceUrl || element.image.contentUrl;
            if (imageUrl) {
              const imageText = ocrImage(imageUrl);
              ocrText += `\n[スライド${index + 1}の画像テキスト]\n${imageText}\n`;
            }
          }
        });
      }
    });

  } catch (error) {
    console.error('OCR処理エラー:', error);
  }

  return ocrText;
}

/**
 * 画像URLからOCR処理
 * @param {string} imageUrl - 画像のURL
 * @returns {string} 抽出されたテキスト
 */
function ocrImage(imageUrl) {
  try {
    const apiUrl = `https://vision.googleapis.com/v1/images:annotate?key=${CONFIG.VISION_API_KEY}`;

    // 画像を取得してBase64エンコード
    const response = UrlFetchApp.fetch(imageUrl);
    const base64Image = Utilities.base64Encode(response.getBlob().getBytes());

    const requestBody = {
      requests: [{
        image: {
          content: base64Image
        },
        features: [{
          type: 'TEXT_DETECTION',
          maxResults: 1
        }]
      }]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody)
    };

    const ocrResponse = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(ocrResponse.getContentText());

    if (result.responses && result.responses[0].fullTextAnnotation) {
      return result.responses[0].fullTextAnnotation.text;
    }

    return '';

  } catch (error) {
    console.error('画像OCRエラー:', error);
    return '';
  }
}

// ===== AI処理関数 =====

/**
 * AIを使用して高度な情報抽出
 * @param {string} text - 抽出対象テキスト
 * @returns {Object} 抽出された情報
 */
function extractWithAI(text) {
  const extractedInfo = {
    staffName: '',
    clientName: '',
    period: '',
    eventType: '',
    prizeCategory: '',
    prizeName: '',
    unitPrice: null,
    quantity: null,
    moq: null,
    leadTime: '',
    vendor: '',
    venueName: '',
    venueCost: null,
    successFactors: [],
    risks: [],
    tags: [],
    summary: '',
    confidenceScore: 0
  };

  try {
    // OpenAI APIを使用した高度な抽出
    if (CONFIG.OPENAI_API_KEY && CONFIG.OPENAI_API_KEY !== 'YOUR_OPENAI_KEY_HERE') {
      const aiExtracted = extractWithOpenAI(text);
      Object.assign(extractedInfo, aiExtracted);
    } else {
      // フォールバック：改良版パターンマッチング
      Object.assign(extractedInfo, enhancedPatternExtraction(text));
    }

    // カテゴリの自動分類
    extractedInfo.tags = generateSmartTags(text, extractedInfo);

    // サマリーの生成
    extractedInfo.summary = generateSummary(extractedInfo);

  } catch (error) {
    console.error('AI抽出エラー:', error);
  }

  return extractedInfo;
}

/**
 * OpenAI APIを使用した情報抽出
 * @param {string} text - 入力テキスト
 * @returns {Object} 抽出結果
 */
function extractWithOpenAI(text) {
  try {
    const apiUrl = 'https://api.openai.com/v1/chat/completions';

    const prompt = `
以下のテキストから景品・イベント関連の情報を抽出してJSON形式で返してください：

抽出項目:
- clientName: クライアント名
- period: 実施時期
- eventType: イベント種別
- prizeName: 景品名
- unitPrice: 単価（数値のみ）
- quantity: 数量（数値のみ）
- vendor: 協力会社名
- venueName: 会場名

テキスト:
${text.substring(0, 3000)}
`;

    const requestBody = {
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'system',
          content: 'あなたは情報抽出の専門家です。与えられたテキストから正確に情報を抽出してください。'
        },
        {
          role: 'user',
          content: prompt
        }
      ],
      temperature: 0.3,
      max_tokens: 1000
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${CONFIG.OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(requestBody)
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());

    if (result.choices && result.choices[0].message.content) {
      const extracted = JSON.parse(result.choices[0].message.content);
      return extracted;
    }

  } catch (error) {
    console.error('OpenAI API エラー:', error);
  }

  return {};
}

/**
 * 改良版パターンマッチング抽出
 * @param {string} text - 入力テキスト
 * @returns {Object} 抽出結果
 */
function enhancedPatternExtraction(text) {
  const info = {};

  // クライアント名の抽出（より柔軟に）
  const clientPatterns = [
    /(?:クライアント|顧客|お客様|発注元)[：:]\s*([^\s\n]+)/,
    /([株式会社][^\s]+会社)/,
    /([^\s]+株式会社)/
  ];

  for (const pattern of clientPatterns) {
    const match = text.match(pattern);
    if (match) {
      info.clientName = match[1];
      break;
    }
  }

  // 価格の抽出（自然言語対応）
  const pricePatterns = [
    /(?:単価|価格|金額)[：:]?\s*(?:約|およそ)?([￥¥]?)([\d,]+)円?/,
    /([￥¥])([\d,]+)円?\/個/,
    /@([￥¥]?)([\d,]+)円?/,
    /(?:ワンコイン|500円程度)/
  ];

  for (const pattern of pricePatterns) {
    const match = text.match(pattern);
    if (match) {
      if (match[0].includes('ワンコイン')) {
        info.unitPrice = 500;
      } else {
        info.unitPrice = parseInt(match[match.length - 1].replace(/,/g, ''));
      }
      break;
    }
  }

  // 納期の抽出（自然言語対応）
  const leadTimePatterns = [
    /(?:納期|リードタイム)[：:]?\s*(?:約|およそ)?([\d]+)\s*(日|週間|ヶ月|営業日)/,
    /(?:最短|通常)([\d]+)(日|週間)(?:で|程度)/,
    /(?:即納|即日)/
  ];

  for (const pattern of leadTimePatterns) {
    const match = text.match(pattern);
    if (match) {
      if (match[0].includes('即')) {
        info.leadTime = '即日';
      } else {
        info.leadTime = match[1] + match[2];
      }
      break;
    }
  }

  return info;
}

// ===== チャットボット関数 =====

/**
 * チャットボットのメインハンドラー
 * @param {string} question - ユーザーの質問
 * @param {string} userId - ユーザーID
 * @returns {string} AIの回答
 */
function handleChatbotQuery(question, userId) {
  console.log(`質問受信: ${question} (ユーザー: ${userId})`);

  try {
    // 質問の意図を解析
    const intent = analyzeIntent(question);

    let response = '';

    switch (intent.type) {
      case 'SEARCH':
        response = searchKnowledge(intent.keywords);
        break;

      case 'ESTIMATE':
        response = generateEstimate(intent.parameters);
        break;

      case 'TREND':
        response = analyzeTrends(intent.period);
        break;

      case 'RECOMMENDATION':
        response = recommendItems(intent.criteria);
        break;

      default:
        response = generalSearch(question);
    }

    // 会話履歴を保存
    saveConversationHistory(userId, question, response);

    return response;

  } catch (error) {
    console.error('チャットボットエラー:', error);
    return 'エラーが発生しました。もう一度お試しください。';
  }
}

/**
 * 質問の意図を解析
 * @param {string} question - 質問文
 * @returns {Object} 意図解析結果
 */
function analyzeIntent(question) {
  const intent = {
    type: 'GENERAL',
    keywords: [],
    parameters: {}
  };

  // 検索クエリの判定
  if (question.includes('教えて') || question.includes('検索') || question.includes('探して')) {
    intent.type = 'SEARCH';

    // キーワード抽出
    const keywords = extractKeywords(question);
    intent.keywords = keywords;
  }

  // 見積もり依頼の判定
  if (question.includes('見積') || question.includes('予算') || question.includes('いくら')) {
    intent.type = 'ESTIMATE';

    // 数量抽出
    const quantityMatch = question.match(/(\d+)[個枚]/);
    if (quantityMatch) {
      intent.parameters.quantity = parseInt(quantityMatch[1]);
    }

    // 予算抽出
    const budgetMatch = question.match(/(\d+)万円/);
    if (budgetMatch) {
      intent.parameters.budget = parseInt(budgetMatch[1]) * 10000;
    }
  }

  // トレンド分析の判定
  if (question.includes('トレンド') || question.includes('人気') || question.includes('流行')) {
    intent.type = 'TREND';

    // 期間抽出
    if (question.includes('今月')) {
      intent.period = 'THIS_MONTH';
    } else if (question.includes('今年')) {
      intent.period = 'THIS_YEAR';
    } else {
      intent.period = 'RECENT';
    }
  }

  // レコメンドの判定
  if (question.includes('おすすめ') || question.includes('提案')) {
    intent.type = 'RECOMMENDATION';
  }

  return intent;
}

/**
 * ナレッジDBを検索
 * @param {Array} keywords - 検索キーワード
 * @returns {string} 検索結果
 */
function searchKnowledge(keywords) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    const results = [];

    // ヘッダー行をスキップして検索
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let score = 0;

      // キーワードマッチングスコア計算
      keywords.forEach(keyword => {
        row.forEach(cell => {
          if (String(cell).toLowerCase().includes(keyword.toLowerCase())) {
            score += 1;
          }
        });
      });

      if (score > 0) {
        results.push({
          score: score,
          data: row
        });
      }
    }

    // スコア順にソート
    results.sort((a, b) => b.score - a.score);

    // 上位3件を整形して返す
    if (results.length === 0) {
      return '該当する情報が見つかりませんでした。別のキーワードでお試しください。';
    }

    let response = `${keywords.join('、')}に関する情報が${results.length}件見つかりました：\n\n`;

    results.slice(0, 3).forEach((result, index) => {
      const row = result.data;
      response += `【結果${index + 1}】\n`;
      response += `クライアント: ${row[2]}\n`;
      response += `景品: ${row[6]}\n`;
      response += `単価: ${row[7]}円\n`;
      response += `協力会社: ${row[11]}\n`;
      response += `---\n`;
    });

    return response;

  } catch (error) {
    console.error('検索エラー:', error);
    return '検索中にエラーが発生しました。';
  }
}

/**
 * トレンド分析
 * @param {string} period - 分析期間
 * @returns {string} トレンド分析結果
 */
function analyzeTrends(period) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    // 期間でフィルタリング
    const now = new Date();
    const filteredData = data.slice(1).filter(row => {
      const timestamp = new Date(row[0]);

      switch(period) {
        case 'THIS_MONTH':
          return timestamp.getMonth() === now.getMonth() &&
                 timestamp.getFullYear() === now.getFullYear();
        case 'THIS_YEAR':
          return timestamp.getFullYear() === now.getFullYear();
        default:
          // 直近3ヶ月
          const threeMonthsAgo = new Date();
          threeMonthsAgo.setMonth(now.getMonth() - 3);
          return timestamp > threeMonthsAgo;
      }
    });

    // 景品カテゴリ別集計
    const categoryCount = {};
    filteredData.forEach(row => {
      const category = row[5]; // 景品カテゴリ
      categoryCount[category] = (categoryCount[category] || 0) + 1;
    });

    // ランキング作成
    const ranking = Object.entries(categoryCount)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);

    let response = '📊 トレンド分析結果\n\n';
    response += '【人気景品カテゴリ TOP5】\n';

    ranking.forEach((item, index) => {
      const emoji = ['🥇', '🥈', '🥉', '4️⃣', '5️⃣'][index];
      response += `${emoji} ${item[0]} (${item[1]}件)\n`;
    });

    return response;

  } catch (error) {
    console.error('トレンド分析エラー:', error);
    return 'トレンド分析中にエラーが発生しました。';
  }
}

// ===== ユーティリティ関数 =====

/**
 * 信頼度スコアを計算
 * @param {Object} data - 抽出されたデータ
 * @returns {number} 信頼度スコア（0-100）
 */
function calculateConfidenceScore(data) {
  let score = 0;
  let fields = 0;

  // 各フィールドの充実度をチェック
  const importantFields = ['clientName', 'prizeName', 'unitPrice', 'quantity', 'vendor'];

  importantFields.forEach(field => {
    if (data[field]) {
      score += 20;
    }
    fields++;
  });

  // 数値の妥当性チェック
  if (data.unitPrice && data.unitPrice > 0 && data.unitPrice < 1000000) {
    score += 10;
  }

  if (data.quantity && data.quantity > 0 && data.quantity < 1000000) {
    score += 10;
  }

  return Math.min(score, 100);
}

/**
 * スマートタグの生成
 * @param {string} text - 元テキスト
 * @param {Object} extractedInfo - 抽出された情報
 * @returns {Array} タグの配列
 */
function generateSmartTags(text, extractedInfo) {
  const tags = [];

  // 季節タグ
  if (text.match(/春|桜|新年度|新生活/)) tags.push('春季');
  if (text.match(/夏|海|プール|花火/)) tags.push('夏季');
  if (text.match(/秋|紅葉|ハロウィン/)) tags.push('秋季');
  if (text.match(/冬|クリスマス|年末|正月/)) tags.push('冬季');

  // 価格帯タグ
  if (extractedInfo.unitPrice) {
    if (extractedInfo.unitPrice < 100) tags.push('低価格帯');
    else if (extractedInfo.unitPrice < 500) tags.push('中価格帯');
    else if (extractedInfo.unitPrice < 1000) tags.push('高価格帯');
    else tags.push('プレミアム');
  }

  // ターゲットタグ
  if (text.match(/女性|レディース|女子/)) tags.push('女性向け');
  if (text.match(/男性|メンズ|男子/)) tags.push('男性向け');
  if (text.match(/子供|キッズ|ファミリー/)) tags.push('ファミリー向け');
  if (text.match(/シニア|高齢/)) tags.push('シニア向け');

  // 用途タグ
  if (text.match(/販促|プロモーション|キャンペーン/)) tags.push('販促');
  if (text.match(/記念品|周年/)) tags.push('記念品');
  if (text.match(/ノベルティ/)) tags.push('ノベルティ');

  return [...new Set(tags)]; // 重複を除去
}

/**
 * OCR処理ログを記録
 * @param {string} fileName - ファイル名
 * @param {string} status - 処理ステータス
 * @param {Object} data - 抽出データ
 * @param {string} error - エラーメッセージ
 */
function logOCRProcess(fileName, status, data = null, error = null) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let logSheet = spreadsheet.getSheetByName(CONFIG.OCR_LOG_SHEET);

    if (!logSheet) {
      logSheet = spreadsheet.insertSheet(CONFIG.OCR_LOG_SHEET);
      // ヘッダー設定
      logSheet.getRange(1, 1, 1, 5).setValues([
        ['処理日時', 'ファイル名', 'ステータス', '抽出データ', 'エラー']
      ]);
    }

    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
      new Date(),
      fileName,
      status,
      data ? JSON.stringify(data) : '',
      error || ''
    ]]);

  } catch (err) {
    console.error('ログ記録エラー:', err);
  }
}

/**
 * 会話履歴を保存
 * @param {string} userId - ユーザーID
 * @param {string} question - 質問
 * @param {string} answer - 回答
 */
function saveConversationHistory(userId, question, answer) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let historySheet = spreadsheet.getSheetByName(CONFIG.HISTORY_SHEET);

    if (!historySheet) {
      historySheet = spreadsheet.insertSheet(CONFIG.HISTORY_SHEET);
      // ヘッダー設定
      historySheet.getRange(1, 1, 1, 5).setValues([
        ['日時', 'ユーザーID', '質問', '回答', '満足度']
      ]);
    }

    const lastRow = historySheet.getLastRow();
    historySheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
      new Date(),
      userId,
      question,
      answer,
      '' // 満足度は後から更新
    ]]);

  } catch (error) {
    console.error('会話履歴保存エラー:', error);
  }
}

// ===== Webhook関数（チャット連携用） =====

/**
 * Googleチャット/Slack からのWebhookを受信
 * @param {Object} e - イベントオブジェクト
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Googleチャットの場合
    if (data.type === 'MESSAGE') {
      const message = data.message.text;
      const userId = data.user.email;

      const response = handleChatbotQuery(message, userId);

      return ContentService
        .createTextOutput(JSON.stringify({
          text: response
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Slackの場合
    if (data.event && data.event.type === 'message') {
      const message = data.event.text;
      const userId = data.event.user;

      const response = handleChatbotQuery(message, userId);

      // Slack APIで返信
      postToSlack(data.event.channel, response);
    }

  } catch (error) {
    console.error('Webhookエラー:', error);

    return ContentService
      .createTextOutput(JSON.stringify({
        error: 'エラーが発生しました'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===== 初期セットアップ =====

/**
 * システムの完全セットアップ
 */
function setupSystemV2() {
  console.log('=== ナレッジ共有基盤 v2.0 セットアップ開始 ===');

  try {
    // 1. スプレッドシートの作成/確認
    let spreadsheet;
    if (CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
      spreadsheet = SpreadsheetApp.create('ナレッジ共有DB_v2');
      console.log('新しいスプレッドシートを作成しました');
      console.log('ID:', spreadsheet.getId());
    } else {
      spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    }

    // 2. メインシートの作成
    let mainSheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    if (!mainSheet) {
      mainSheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
      // 拡張ヘッダー設定
      const headers = [
        '登録日時', '担当者名', 'クライアント名', '実施時期', 'イベント種別',
        '景品カテゴリ', '具体的な景品名', '単価', '発注数量', 'MOQ',
        '納期', '協力会社名', '協力会社評価', '会場名', '会場費用',
        '成功要因', '失敗・反省点', '企画書URL', 'タグ', '入力方式',
        'OCR処理', '信頼度スコア', '元ファイル名', '処理日時', 'カテゴリタグ',
        '類似案件ID', '特記事項', '画像URL'
      ];
      mainSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      mainSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    // 3. 処理用フォルダの作成
    const rootFolder = DriveApp.getRootFolder();
    let sourceFolder, processedFolder;

    try {
      sourceFolder = DriveApp.getFolderById(CONFIG.SOURCE_FOLDER_ID);
    } catch (e) {
      sourceFolder = rootFolder.createFolder('ナレッジDB_未処理');
      console.log('未処理フォルダを作成:', sourceFolder.getId());
    }

    try {
      processedFolder = DriveApp.getFolderById(CONFIG.PROCESSED_FOLDER_ID);
    } catch (e) {
      processedFolder = rootFolder.createFolder('ナレッジDB_処理済み');
      console.log('処理済みフォルダを作成:', processedFolder.getId());
    }

    // 4. バッチ処理トリガーの設定
    ScriptApp.newTrigger('batchProcessHistoricalDocuments')
      .timeBased()
      .everyHours(1)
      .create();

    console.log('バッチ処理トリガーを設定しました（1時間ごと）');

    // 5. Webhook URLの生成
    const scriptUrl = ScriptApp.getService().getUrl();
    console.log('Webhook URL:', scriptUrl);

    console.log('\n=== セットアップ完了 ===');
    console.log('次のステップ:');
    console.log('1. CONFIG内のAPIキーを設定してください');
    console.log('2. 過去資料を「ナレッジDB_未処理」フォルダにアップロード');
    console.log('3. Googleチャット/SlackにWebhook URLを設定');

  } catch (error) {
    console.error('セットアップエラー:', error);
    throw error;
  }
}