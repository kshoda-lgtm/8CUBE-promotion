/**
 * JSON処理機能（無料版）
 * Pythonで事前処理されたJSONファイルを読み込み、スプレッドシートに自動登録
 */

// ===== JSON取り込み処理システム =====

/**
 * メイン関数：JSONファイルからデータを取り込み
 */
function processJSONFiles() {
  console.log('🚀 JSON取り込みシステム開始');

  try {
    // スプレッドシート準備
    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = getOrCreateMainSheet(spreadsheet);

    // Google Drive上のJSONファイルを検索
    const jsonFiles = findJSONFiles();

    if (jsonFiles.length === 0) {
      console.log('⚠️ JSONファイルが見つかりません');
      return;
    }

    console.log(`📁 ${jsonFiles.length}個のJSONファイルを発見`);

    // 各JSONファイルを処理
    let processedCount = 0;
    jsonFiles.forEach((file, index) => {
      try {
        const success = processJSONFile(file, sheet, index + 1);
        if (success) processedCount++;
      } catch (error) {
        console.error(`❌ ファイル処理エラー [${file.getName()}]:`, error);
      }
    });

    console.log(`✅ JSON取り込み完了: ${processedCount}/${jsonFiles.length}件`);
    console.log(`📊 スプレッドシート: ${spreadsheet.getUrl()}`);

  } catch (error) {
    console.error('❌ システムエラー:', error);
  }
}

/**
 * Google Drive上のJSONファイルを検索
 */
function findJSONFiles() {
  try {
    // 特定フォルダまたは全体からJSONファイルを検索
    const files = DriveApp.searchFiles(
      'name contains ".json" and mimeType = "application/json"'
    );

    const jsonFiles = [];
    while (files.hasNext()) {
      const file = files.next();
      // PowerPoint由来のJSONファイルのみ対象
      if (file.getName().includes('powerpoint') ||
          file.getName().includes('pptx') ||
          file.getName().includes('presentation')) {
        jsonFiles.push(file);
      }
    }

    return jsonFiles;
  } catch (error) {
    console.error('❌ JSONファイル検索エラー:', error);
    return [];
  }
}

/**
 * 個別JSONファイルの処理
 */
function processJSONFile(file, sheet, fileNumber) {
  try {
    console.log(`📄 処理中 [${fileNumber}]: ${file.getName()}`);

    // JSONデータを読み込み
    const jsonContent = file.getBlob().getDataAsString('UTF-8');
    const data = JSON.parse(jsonContent);

    // エラーチェック
    if (data.error) {
      console.log(`⚠️ ファイル[${fileNumber}]にエラー情報: ${data.error}`);
      return false;
    }

    // データ抽出・整理
    const extractedInfo = extractInfoFromJSON(data);

    // スプレッドシートに追加
    const rowData = formatDataForSpreadsheet(extractedInfo, file.getName());
    appendToSpreadsheet(sheet, rowData);

    console.log(`✅ ファイル[${fileNumber}]処理完了`);
    return true;

  } catch (error) {
    console.error(`❌ JSONファイル処理エラー [${file.getName()}]:`, error);
    return false;
  }
}

/**
 * JSONデータから情報を抽出
 */
function extractInfoFromJSON(data) {
  const summary = data.summary || {};
  const fileInfo = data.file_info || {};

  // 価格情報の処理
  const prices = summary.all_prices || [];
  const avgPrice = prices.length > 0 ?
    Math.round(prices.reduce((a, b) => a + b, 0) / prices.length) : null;
  const minPrice = prices.length > 0 ? Math.min(...prices) : null;
  const maxPrice = prices.length > 0 ? Math.max(...prices) : null;

  // 数量情報の処理
  const quantities = summary.all_quantities || [];
  const totalQuantity = quantities.length > 0 ?
    quantities.reduce((a, b) => a + b, 0) : null;

  // 会社情報
  const companies = summary.all_companies || [];
  const mainCompany = companies.length > 0 ? companies[0] : '';

  // キーワード情報
  const keywords = summary.all_keywords || [];
  const tags = keywords.join(', ');

  // イベント種別の推定
  const eventType = estimateEventType(keywords, data.slides);

  // 実施時期の推定
  const eventDate = estimateEventDate(data.slides);

  return {
    fileName: fileInfo.file_name || '',
    slideCount: fileInfo.slide_count || 0,
    processedAt: fileInfo.processed_at || new Date().toISOString(),
    eventType: eventType,
    eventDate: eventDate,
    mainCompany: mainCompany,
    allCompanies: companies.join(', '),
    avgPrice: avgPrice,
    minPrice: minPrice,
    maxPrice: maxPrice,
    totalQuantity: totalQuantity,
    tags: tags,
    slideTexts: extractAllTexts(data.slides),
    confidenceScore: calculateConfidenceScore(data)
  };
}

/**
 * イベント種別を推定
 */
function estimateEventType(keywords, slides) {
  const typeKeywords = {
    'キャンペーン': ['キャンペーン', 'プレゼント', '景品'],
    '展示会': ['展示会', 'ブース', '出展'],
    'セミナー': ['セミナー', '講座', '研修'],
    'プロモーション': ['プロモーション', '宣伝', 'PR']
  };

  for (const [type, words] of Object.entries(typeKeywords)) {
    if (words.some(word => keywords.includes(word))) {
      return type;
    }
  }

  return '不明';
}

/**
 * 実施時期を推定
 */
function estimateEventDate(slides) {
  if (!slides || slides.length === 0) return '';

  // 全スライドのテキストから日付パターンを検索
  const datePattern = /(\d{4})[年\/\-](\d{1,2})[月\/\-](\d{1,2})?/g;

  for (const slide of slides) {
    const texts = slide.raw_texts || [];
    for (const text of texts) {
      const match = datePattern.exec(text);
      if (match) {
        const year = match[1];
        const month = match[2].padStart(2, '0');
        const day = match[3] ? match[3].padStart(2, '0') : '01';
        return `${year}/${month}/${day}`;
      }
    }
  }

  return '';
}

/**
 * 全テキストを抽出
 */
function extractAllTexts(slides) {
  if (!slides || slides.length === 0) return '';

  const allTexts = [];
  slides.forEach(slide => {
    if (slide.raw_texts) {
      allTexts.push(...slide.raw_texts);
    }
  });

  return allTexts.join(' ').substring(0, 1000); // 1000文字制限
}

/**
 * 信頼度スコアを計算
 */
function calculateConfidenceScore(data) {
  let score = 0;

  // ファイル情報があれば +20
  if (data.file_info) score += 20;

  // 価格情報があれば +30
  if (data.summary && data.summary.all_prices && data.summary.all_prices.length > 0) {
    score += 30;
  }

  // 会社情報があれば +25
  if (data.summary && data.summary.all_companies && data.summary.all_companies.length > 0) {
    score += 25;
  }

  // キーワードがあれば +15
  if (data.summary && data.summary.all_keywords && data.summary.all_keywords.length > 0) {
    score += 15;
  }

  // スライド数に応じて +10
  if (data.file_info && data.file_info.slide_count > 5) {
    score += 10;
  }

  return Math.min(score, 100); // 最大100%
}

/**
 * スプレッドシート用にデータをフォーマット
 */
function formatDataForSpreadsheet(info, fileName) {
  return [
    new Date(), // A: 登録日時
    '', // B: 担当者名（手動入力）
    '', // C: クライアント名（後で手動入力）
    info.eventDate, // D: 実施時期
    info.eventType, // E: イベント種別
    '', // F: 景品カテゴリ（後で分類）
    '', // G: 具体的な景品名（後で入力）
    info.avgPrice, // H: 単価（平均）
    info.totalQuantity, // I: 発注数量
    '', // J: MOQ（後で入力）
    '', // K: 納期（後で入力）
    info.mainCompany, // L: 協力会社名
    '', // M: 協力会社評価（後で評価）
    '', // N: 会場名（後で入力）
    '', // O: 会場費用（後で入力）
    '', // P: 成功要因（後で入力）
    '', // Q: 失敗・反省点（後で入力）
    '', // R: 企画書URL（後で入力）
    info.tags, // S: タグ
    info.confidenceScore, // T: 信頼度スコア
    fileName, // U: 元ファイル名
    info.slideTexts, // V: 抽出テキスト（参考用）
    info.allCompanies // W: 全会社名
  ];
}

/**
 * スプレッドシートまたはシートを取得・作成
 */
function getOrCreateSpreadsheet() {
  const spreadsheetId = '1LVjGOulUFlrsq1TOwZR3hEXO_c4DSDPM0UkQ9doBKMw';
  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    console.log('📊 新しいスプレッドシートを作成');
    return SpreadsheetApp.create('ナレッジ共有基盤DB（JSON版）');
  }
}

function getOrCreateMainSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName('メインDB');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('メインDB');
    // ヘッダー行を追加
    const headers = [
      '登録日時', '担当者名', 'クライアント名', '実施時期', 'イベント種別',
      '景品カテゴリ', '具体的な景品名', '単価', '発注数量', 'MOQ', '納期',
      '協力会社名', '協力会社評価', '会場名', '会場費用', '成功要因',
      '失敗・反省点', '企画書URL', 'タグ', '信頼度スコア',
      '元ファイル名', '抽出テキスト', '全会社名'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // ヘッダー行をフォーマット
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sheet;
}

/**
 * スプレッドシートにデータを追加
 */
function appendToSpreadsheet(sheet, rowData) {
  try {
    sheet.appendRow(rowData);
    console.log('📊 データをスプレッドシートに追加');
  } catch (error) {
    console.error('❌ スプレッドシート書き込みエラー:', error);
    throw error;
  }
}

/**
 * テスト用：サンプルJSONで動作確認
 */
function testJSONProcessing() {
  console.log('🧪 JSON処理テスト開始');

  // サンプルJSONデータ
  const sampleData = {
    file_info: {
      file_name: 'test_presentation.pptx',
      slide_count: 10,
      processed_at: new Date().toISOString()
    },
    summary: {
      all_prices: [500, 750, 1000],
      all_quantities: [100, 200],
      all_companies: ['株式会社テスト印刷', 'サンプル製作所'],
      all_keywords: ['ノベルティ', 'キャンペーン', 'エコ']
    },
    slides: [
      {
        slide_number: 1,
        raw_texts: ['2024年夏のキャンペーン企画', 'エコバッグ配布企画'],
        analyzed_info: {
          prices: [500],
          event_types: ['キャンペーン']
        }
      }
    ]
  };

  try {
    const info = extractInfoFromJSON(sampleData);
    console.log('✅ データ抽出成功:', info);

    const rowData = formatDataForSpreadsheet(info, 'test.json');
    console.log('✅ フォーマット成功:', rowData.length, '列');

  } catch (error) {
    console.error('❌ テストエラー:', error);
  }
}