/**
 * 完全版セットアップ用スクリプト
 * スプレッドシート自動作成＋PowerPoint抽出テスト
 */

// ===== 自動セットアップ関数 =====

/**
 * メイン設定関数 - すべてを自動で設定
 */
function setupComplete() {
  console.log('🚀 ナレッジ共有基盤 完全セットアップ開始');

  try {
    // 1. スプレッドシート自動作成
    const spreadsheetId = createKnowledgeDatabase();
    console.log('✅ スプレッドシート作成完了:', spreadsheetId);

    // 2. テスト用スプレッドシート設定
    CONFIG.SPREADSHEET_ID = spreadsheetId;

    // 3. PowerPoint抽出テスト実行
    const testResults = runExtractionTest();
    console.log('✅ 抽出テスト完了');

    // 4. 結果をスプレッドシートに保存
    saveTestResults(testResults);

    console.log('\n🎉 セットアップ完了！');
    console.log('📊 スプレッドシート:', `https://docs.google.com/spreadsheets/d/${spreadsheetId}`);

    return {
      spreadsheetId: spreadsheetId,
      testResults: testResults
    };

  } catch (error) {
    console.error('❌ セットアップエラー:', error);
    throw error;
  }
}

/**
 * ナレッジデータベース用スプレッドシートを自動作成
 */
function createKnowledgeDatabase() {
  try {
    // 新しいスプレッドシート作成
    const spreadsheet = SpreadsheetApp.create('📊 ナレッジ共有データベース（テスト版）');
    const spreadsheetId = spreadsheet.getId();

    // メインシートの設定
    const mainSheet = spreadsheet.getActiveSheet();
    mainSheet.setName('ナレッジDB');

    // ヘッダー行を作成
    const headers = [
      '登録日時', '担当者名', 'クライアント名', '実施時期', 'イベント種別',
      '景品カテゴリ', '具体的な景品名', '単価', '発注数量', 'MOQ',
      '納期', '協力会社名', '協力会社評価', '会場名', '会場費用',
      '成功要因', '失敗・反省点', '企画書URL', 'タグ', '信頼度スコア'
    ];

    const headerRange = mainSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');

    // 列幅を自動調整
    headers.forEach((header, index) => {
      mainSheet.setColumnWidth(index + 1, 120);
    });

    // テスト結果シートを作成
    const testSheet = spreadsheet.insertSheet('抽出テスト結果');
    const testHeaders = [
      'ファイル番号', 'ファイルURL', '処理状況', 'クライアント名', '景品名',
      '単価', '数量', '協力会社', '信頼度スコア', '抽出時刻'
    ];

    const testHeaderRange = testSheet.getRange(1, 1, 1, testHeaders.length);
    testHeaderRange.setValues([testHeaders]);
    testHeaderRange.setFontWeight('bold');
    testHeaderRange.setBackground('#34a853');
    testHeaderRange.setFontColor('white');

    console.log('📊 スプレッドシート作成完了:', spreadsheetId);
    return spreadsheetId;

  } catch (error) {
    console.error('❌ スプレッドシート作成エラー:', error);
    throw error;
  }
}

// ===== 設定値（自動更新される） =====
let CONFIG = {
  SPREADSHEET_ID: '', // 自動で設定される

  // テスト対象ファイル
  TEST_FILES: [
    'https://docs.google.com/presentation/d/1MlVP3kEd6MQtyo2w8ANOZiDTFUAeHYj3/edit?usp=drive_link&ouid=106090073074900580884&rtpof=true&sd=true',
    'https://docs.google.com/presentation/d/12OF7jJHE_WgEk_mQaCe06Cl4ojNhcAqo/edit?usp=drive_link&ouid=106090073074900580884&rtpof=true&sd=true',
    'https://docs.google.com/presentation/d/1opasIPp6zOpkLwI3gAQC-Br6dbdyDf6c/edit?usp=drive_link&ouid=106090073074900580884&rtpof=true&sd=true',
    'https://docs.google.com/presentation/d/1UrTXzw3pSMDp4aAxyubTnAruMIzuajVl/edit?usp=drive_link&ouid=106090073074900580884&rtpof=true&sd=true',
    'https://docs.google.com/presentation/d/1fw2bB5SUQ_xx37H3jYeiyfDhtaqLXgOD/edit?usp=drive_link&ouid=106090073074900580884&rtpof=true&sd=true'
  ]
};

// ===== PowerPoint抽出テスト =====

/**
 * 全PowerPointファイルの抽出テストを実行
 */
function runExtractionTest() {
  console.log('🔬 PowerPoint抽出テスト開始');

  const results = [];

  CONFIG.TEST_FILES.forEach((url, index) => {
    try {
      console.log(`\n📁 ファイル${index + 1}を処理中...`);

      const result = extractFromSingleFile(url, index + 1);
      results.push(result);

      // レート制限回避
      Utilities.sleep(1000);

    } catch (error) {
      console.error(`❌ ファイル${index + 1}でエラー:`, error);
      results.push({
        fileNumber: index + 1,
        fileUrl: url,
        success: false,
        error: error.toString(),
        timestamp: new Date()
      });
    }
  });

  return results;
}

/**
 * 単一ファイルから情報を抽出
 */
function extractFromSingleFile(url, fileNumber) {
  try {
    // プレゼンテーションIDを抽出
    const presentationId = extractPresentationId(url);
    if (!presentationId) {
      throw new Error('プレゼンテーションIDが取得できません');
    }

    // テキスト抽出
    const extractedText = extractTextFromPresentation(presentationId);

    // 情報解析
    const extractedInfo = analyzeExtractedText(extractedText);

    console.log(`✅ ファイル${fileNumber}: ${extractedInfo.prizeName || '景品名不明'} - 信頼度${extractedInfo.confidence}%`);

    return {
      fileNumber: fileNumber,
      fileUrl: url,
      success: true,
      data: extractedInfo,
      sourceTextLength: extractedText.length,
      timestamp: new Date()
    };

  } catch (error) {
    console.error(`❌ ファイル${fileNumber}抽出エラー:`, error);
    throw error;
  }
}

/**
 * URLからプレゼンテーションIDを抽出
 */
function extractPresentationId(url) {
  const match = url.match(/\/presentation\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

/**
 * プレゼンテーションからテキストを抽出
 */
function extractTextFromPresentation(presentationId) {
  try {
    const presentation = Slides.Presentations.get(presentationId);
    let allText = '';

    presentation.slides.forEach((slide, index) => {
      allText += `\n=== スライド ${index + 1} ===\n`;

      if (slide.pageElements) {
        slide.pageElements.forEach(element => {
          // テキストボックス・図形
          if (element.shape && element.shape.text) {
            const textElements = element.shape.text.textElements;
            if (textElements) {
              textElements.forEach(textElement => {
                if (textElement.textRun && textElement.textRun.content) {
                  allText += textElement.textRun.content;
                }
              });
            }
          }

          // テーブル
          if (element.table && element.table.tableRows) {
            allText += '\n[テーブル]\n';
            element.table.tableRows.forEach(row => {
              if (row.tableCells) {
                row.tableCells.forEach(cell => {
                  if (cell.text && cell.text.textElements) {
                    cell.text.textElements.forEach(textElement => {
                      if (textElement.textRun && textElement.textRun.content) {
                        allText += textElement.textRun.content + ' | ';
                      }
                    });
                  }
                });
                allText += '\n';
              }
            });
          }
        });
      }
    });

    return allText;

  } catch (error) {
    console.error('テキスト抽出エラー:', error);
    throw error;
  }
}

/**
 * 抽出テキストを解析
 */
function analyzeExtractedText(text) {
  const info = {
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
    tags: [],
    confidence: 0
  };

  // クライアント名の抽出
  const clientPatterns = [
    /(?:クライアント|顧客|お客様)[：:]\s*([^\s\n]+)/,
    /([株式会社][^\s\n]+)/,
    /([^\s\n]+株式会社)/,
    /([^\s\n]+様)/
  ];

  for (const pattern of clientPatterns) {
    const match = text.match(pattern);
    if (match) {
      info.clientName = match[1].replace('様', '');
      break;
    }
  }

  // 景品名の抽出（拡張版）
  const prizePatterns = [
    // 一般的な景品
    /(エコバッグ|タンブラー|ボールペン|マグカップ|タオル|キーホルダー|ステッカー)/,
    /(カレンダー|Tシャツ|パーカー|キャップ|トートバッグ|USB|モバイルバッテリー)/,
    /(スマホスタンド|団扇|うちわ|クリアファイル|ノート|メモ帳|ペン|マスク)/,
    /(除菌|ハンドクリーム|ティッシュ|ウェットティッシュ|手帳|カレンダー)/,
    // 具体的な商品名
    /([^\s]+バッグ|[^\s]+ペン|[^\s]+タンブラー|[^\s]+マグ)/
  ];

  for (const pattern of prizePatterns) {
    const match = text.match(pattern);
    if (match) {
      info.prizeName = match[1];
      break;
    }
  }

  // 価格の抽出（自然言語対応）
  const pricePatterns = [
    /(?:単価|価格|金額)[：:\s]*([¥￥]?)([\d,]+)円?/,
    /([¥￥])([\d,]+)円?(?:\/個|\/枚|\/本)?/,
    /@\s*([¥￥]?)([\d,]+)円?/,
    /約\s*([¥￥]?)([\d,]+)円?/,
    /([\d,]+)円程度/,
    /ワンコイン|500円程度/
  ];

  for (const pattern of pricePatterns) {
    const match = text.match(pattern);
    if (match) {
      if (match[0].includes('ワンコイン')) {
        info.unitPrice = 500;
      } else {
        const numbers = match.filter(m => m && /^\d/.test(m.replace(/,/g, '')));
        if (numbers.length > 0) {
          info.unitPrice = parseInt(numbers[numbers.length - 1].replace(/,/g, ''));
        }
      }
      break;
    }
  }

  // 数量の抽出
  const quantityPatterns = [
    /(?:数量|ロット|個数)[：:\s]*([\d,]+)\s*(?:個|枚|本|セット)?/,
    /([\d,]+)\s*(?:個|枚|本|セット)(?:配布|製作)/,
    /合計\s*([\d,]+)(?:個|枚|本)/
  ];

  for (const pattern of quantityPatterns) {
    const match = text.match(pattern);
    if (match) {
      info.quantity = parseInt(match[1].replace(/,/g, ''));
      break;
    }
  }

  // 協力会社の抽出
  const vendorPatterns = [
    /(?:制作会社|協力会社|発注先|印刷会社)[：:\s]*([^\s\n]+(?:株式会社|有限会社|印刷|製作所)[^\s\n]*)/,
    /([^\s\n]+(?:株式会社|有限会社|印刷|製作所)[^\s\n]*)/
  ];

  for (const pattern of vendorPatterns) {
    const match = text.match(pattern);
    if (match) {
      info.vendor = match[1];
      break;
    }
  }

  // 実施時期の抽出
  const periodPatterns = [
    /(\d{4}年\d{1,2}月)/,
    /(\d{4}年Q[1-4])/,
    /(春|夏|秋|冬)(?:季|期)?/,
    /(\d{1,2}月)/
  ];

  for (const pattern of periodPatterns) {
    const match = text.match(pattern);
    if (match) {
      info.period = match[1];
      break;
    }
  }

  // タグ生成
  info.tags = generateAutoTags(text, info);

  // 信頼度計算
  info.confidence = calculateConfidenceScore(info);

  return info;
}

/**
 * 自動タグ生成
 */
function generateAutoTags(text, info) {
  const tags = [];

  // 季節タグ
  if (text.match(/春|桜|新年度/)) tags.push('春季');
  if (text.match(/夏|海|プール|暑中/)) tags.push('夏季');
  if (text.match(/秋|紅葉|ハロウィン/)) tags.push('秋季');
  if (text.match(/冬|クリスマス|年末|正月/)) tags.push('冬季');

  // 価格帯タグ
  if (info.unitPrice) {
    if (info.unitPrice < 100) tags.push('低価格帯');
    else if (info.unitPrice < 500) tags.push('中価格帯');
    else if (info.unitPrice < 1000) tags.push('高価格帯');
    else tags.push('プレミアム');
  }

  // 属性タグ
  if (text.match(/エコ|環境|SDGs|サステナブル/)) tags.push('エコ');
  if (text.match(/高級|プレミアム|限定|特別/)) tags.push('高級');
  if (text.match(/オリジナル|カスタム|名入れ|特注/)) tags.push('オリジナル');
  if (text.match(/大量|1000個|2000個|5000個/)) tags.push('大量発注');

  return tags;
}

/**
 * 信頼度スコア計算
 */
function calculateConfidenceScore(info) {
  let score = 0;

  // 主要項目の存在チェック
  if (info.clientName) score += 25;
  if (info.prizeName) score += 25;
  if (info.unitPrice) score += 20;
  if (info.vendor) score += 15;
  if (info.quantity) score += 10;
  if (info.period) score += 5;

  return Math.min(score, 100);
}

/**
 * テスト結果をスプレッドシートに保存
 */
function saveTestResults(results) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const testSheet = spreadsheet.getSheetByName('抽出テスト結果');

    results.forEach((result, index) => {
      const row = [
        result.fileNumber,
        result.fileUrl,
        result.success ? '✅ 成功' : '❌ 失敗',
        result.success ? result.data.clientName || '未検出' : 'エラー',
        result.success ? result.data.prizeName || '未検出' : 'エラー',
        result.success ? (result.data.unitPrice ? result.data.unitPrice + '円' : '未検出') : 'エラー',
        result.success ? (result.data.quantity ? result.data.quantity + '個' : '未検出') : 'エラー',
        result.success ? result.data.vendor || '未検出' : 'エラー',
        result.success ? result.data.confidence + '%' : '0%',
        result.timestamp
      ];

      testSheet.getRange(index + 2, 1, 1, row.length).setValues([row]);
    });

    // 成功したものは本番DBにも保存
    const successResults = results.filter(r => r.success);
    if (successResults.length > 0) {
      saveToMainDatabase(successResults);
    }

    console.log(`✅ テスト結果をスプレッドシートに保存（${results.length}件）`);

  } catch (error) {
    console.error('❌ テスト結果保存エラー:', error);
  }
}

/**
 * 成功した抽出結果を本番DBに保存
 */
function saveToMainDatabase(successResults) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const mainSheet = spreadsheet.getSheetByName('ナレッジDB');

    successResults.forEach(result => {
      const data = result.data;
      const row = [
        result.timestamp,
        'システム管理者',
        data.clientName,
        data.period,
        data.eventType,
        data.prizeCategory,
        data.prizeName,
        data.unitPrice,
        data.quantity,
        data.moq,
        data.leadTime,
        data.vendor,
        '', // 協力会社評価
        data.venueName,
        data.venueCost,
        '', // 成功要因
        '', // 失敗・反省点
        result.fileUrl,
        data.tags.join(', '),
        data.confidence
      ];

      const lastRow = mainSheet.getLastRow();
      mainSheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
    });

    console.log(`✅ 本番DBに${successResults.length}件のデータを保存`);

  } catch (error) {
    console.error('❌ 本番DB保存エラー:', error);
  }
}

// ===== 個別実行用関数 =====

/**
 * スプレッドシートのみを作成
 */
function createSpreadsheetOnly() {
  const spreadsheetId = createKnowledgeDatabase();
  console.log('📊 スプレッドシートURL:', `https://docs.google.com/spreadsheets/d/${spreadsheetId}`);
  return spreadsheetId;
}

/**
 * 抽出テストのみを実行（スプレッドシートIDを手動設定）
 */
function runTestOnly() {
  if (!CONFIG.SPREADSHEET_ID) {
    console.error('❌ CONFIG.SPREADSHEET_IDを設定してください');
    return;
  }

  const results = runExtractionTest();
  saveTestResults(results);
  return results;
}

/**
 * 単一ファイルテスト（デバッグ用）
 */
function testFirstFileOnly() {
  const result = extractFromSingleFile(CONFIG.TEST_FILES[0], 1);
  console.log('🔍 テスト結果:', result);
  return result;
}