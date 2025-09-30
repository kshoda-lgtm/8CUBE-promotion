/**
 * 修正版: PowerPoint変換対応
 */

// ===== 修正版：PowerPoint自動変換機能 =====

/**
 * PowerPointをGoogleスライドに変換してからテキスト抽出
 */
function convertAndExtractFromFile(url, fileNumber) {
  try {
    console.log(`📁 ファイル${fileNumber}: 変換処理開始`);

    // プレゼンテーションIDを抽出
    const presentationId = extractPresentationId(url);
    if (!presentationId) {
      throw new Error('プレゼンテーションIDが取得できません');
    }

    // まず直接アクセスを試行
    try {
      const presentation = Slides.Presentations.get(presentationId);
      console.log(`✅ ファイル${fileNumber}: 既にGoogleスライド形式`);

      const extractedText = extractTextFromPresentation(presentationId);
      const extractedInfo = analyzeExtractedText(extractedText);

      return {
        fileNumber: fileNumber,
        fileUrl: url,
        success: true,
        data: extractedInfo,
        sourceTextLength: extractedText.length,
        timestamp: new Date(),
        conversionStatus: 'No conversion needed'
      };

    } catch (apiError) {
      console.log(`⚠️ ファイル${fileNumber}: PowerPoint形式のため変換が必要`);

      // PowerPointの場合は変換処理
      return convertPowerPointToSlides(url, fileNumber, presentationId);
    }

  } catch (error) {
    console.error(`❌ ファイル${fileNumber}処理エラー:`, error);
    return {
      fileNumber: fileNumber,
      fileUrl: url,
      success: false,
      error: error.toString(),
      timestamp: new Date(),
      conversionStatus: 'Failed'
    };
  }
}

/**
 * PowerPointをGoogleスライドに変換
 */
function convertPowerPointToSlides(originalUrl, fileNumber, originalId) {
  try {
    console.log(`🔄 ファイル${fileNumber}: PowerPoint→Googleスライド変換中...`);

    // Drive APIを使用してファイル情報を取得
    const file = Drive.Files.get(originalId);

    // PowerPointファイルをダウンロード
    const blob = DriveApp.getFileById(originalId).getBlob();

    // Googleスライド形式で新しいファイルを作成
    const convertedFile = Drive.Files.insert({
      title: `[変換済み] ${file.title}`,
      mimeType: 'application/vnd.google-apps.presentation'
    }, blob);

    console.log(`✅ ファイル${fileNumber}: 変換完了 (新ID: ${convertedFile.id})`);

    // 変換されたファイルからテキスト抽出
    const extractedText = extractTextFromPresentation(convertedFile.id);
    const extractedInfo = analyzeExtractedText(extractedText);

    // 元のURLも保持
    extractedInfo.originalUrl = originalUrl;
    extractedInfo.convertedUrl = `https://docs.google.com/presentation/d/${convertedFile.id}/edit`;

    return {
      fileNumber: fileNumber,
      fileUrl: originalUrl,
      convertedUrl: extractedInfo.convertedUrl,
      success: true,
      data: extractedInfo,
      sourceTextLength: extractedText.length,
      timestamp: new Date(),
      conversionStatus: 'Converted successfully'
    };

  } catch (conversionError) {
    console.error(`❌ ファイル${fileNumber}変換エラー:`, conversionError);

    // 手動変換の案内
    return {
      fileNumber: fileNumber,
      fileUrl: originalUrl,
      success: false,
      error: '自動変換に失敗しました。手動でGoogleスライドに変換してください。',
      conversionStatus: 'Manual conversion required',
      timestamp: new Date()
    };
  }
}

/**
 * 修正版：全ファイルテスト
 */
function runExtractionTestFixed() {
  console.log('🔬 PowerPoint抽出テスト開始（修正版）');

  const results = [];

  CONFIG.TEST_FILES.forEach((url, index) => {
    try {
      console.log(`\n📁 ファイル${index + 1}を処理中...`);

      // 修正版の変換対応関数を使用
      const result = convertAndExtractFromFile(url, index + 1);
      results.push(result);

      // レート制限回避
      Utilities.sleep(2000);

    } catch (error) {
      console.error(`❌ ファイル${index + 1}でエラー:`, error);
      results.push({
        fileNumber: index + 1,
        fileUrl: url,
        success: false,
        error: error.toString(),
        timestamp: new Date(),
        conversionStatus: 'Error'
      });
    }
  });

  return results;
}

/**
 * 修正版セットアップ：変換対応版
 */
function setupCompleteFixed() {
  console.log('🚀 ナレッジ共有基盤 完全セットアップ開始（修正版）');

  try {
    // 既存のスプレッドシートを使用
    CONFIG.SPREADSHEET_ID = '1LVjGOulUFlrsq1TOwZR3hEXO_c4DSDPM0UkQ9doBKMw';

    // 修正版のPowerPoint抽出テスト実行
    const testResults = runExtractionTestFixed();
    console.log('✅ 抽出テスト完了（修正版）');

    // 結果をスプレッドシートに保存
    saveTestResultsFixed(testResults);

    console.log('\n🎉 修正版セットアップ完了！');
    console.log('📊 スプレッドシート:', `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}`);

    // 結果サマリーを表示
    displayTestSummary(testResults);

    return {
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      testResults: testResults
    };

  } catch (error) {
    console.error('❌ 修正版セットアップエラー:', error);
    throw error;
  }
}

/**
 * 修正版：テスト結果保存（変換ステータス付き）
 */
function saveTestResultsFixed(results) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // 修正版テスト結果シートを作成
    let testSheet = spreadsheet.getSheetByName('修正版テスト結果');
    if (!testSheet) {
      testSheet = spreadsheet.insertSheet('修正版テスト結果');

      const testHeaders = [
        'ファイル番号', 'ファイルURL', '処理状況', '変換ステータス',
        'クライアント名', '景品名', '単価', '数量', '協力会社',
        '信頼度スコア', '変換後URL', '抽出時刻'
      ];

      const testHeaderRange = testSheet.getRange(1, 1, 1, testHeaders.length);
      testHeaderRange.setValues([testHeaders]);
      testHeaderRange.setFontWeight('bold');
      testHeaderRange.setBackground('#ff9900');
      testHeaderRange.setFontColor('white');
    }

    results.forEach((result, index) => {
      const row = [
        result.fileNumber,
        result.fileUrl,
        result.success ? '✅ 成功' : '❌ 失敗',
        result.conversionStatus || 'Unknown',
        result.success ? result.data.clientName || '未検出' : 'エラー',
        result.success ? result.data.prizeName || '未検出' : 'エラー',
        result.success ? (result.data.unitPrice ? result.data.unitPrice + '円' : '未検出') : 'エラー',
        result.success ? (result.data.quantity ? result.data.quantity + '個' : '未検出') : 'エラー',
        result.success ? result.data.vendor || '未検出' : 'エラー',
        result.success ? result.data.confidence + '%' : '0%',
        result.convertedUrl || '',
        result.timestamp
      ];

      testSheet.getRange(index + 2, 1, 1, row.length).setValues([row]);
    });

    console.log(`✅ 修正版テスト結果をスプレッドシートに保存（${results.length}件）`);

  } catch (error) {
    console.error('❌ 修正版テスト結果保存エラー:', error);
  }
}

/**
 * テスト結果サマリー表示
 */
function displayTestSummary(results) {
  console.log('\n📊 ===== テスト結果サマリー =====');

  const totalFiles = results.length;
  const successCount = results.filter(r => r.success).length;
  const conversionCount = results.filter(r => r.conversionStatus === 'Converted successfully').length;
  const manualCount = results.filter(r => r.conversionStatus === 'Manual conversion required').length;

  console.log(`📁 総ファイル数: ${totalFiles}`);
  console.log(`✅ 処理成功: ${successCount}/${totalFiles}`);
  console.log(`🔄 自動変換成功: ${conversionCount}`);
  console.log(`⚠️ 手動変換が必要: ${manualCount}`);

  if (manualCount > 0) {
    console.log('\n📝 手動変換が必要なファイル:');
    results.filter(r => r.conversionStatus === 'Manual conversion required')
           .forEach(r => console.log(`   ファイル${r.fileNumber}: ${r.fileUrl}`));
  }
}

// 元のCONFIG設定
let CONFIG = {
  SPREADSHEET_ID: '1LVjGOulUFlrsq1TOwZR3hEXO_c4DSDPM0UkQ9doBKMw',
  TEST_FILES: [
    'https://docs.google.com/presentation/d/1MlVP3kEd6MQtyo2w8ANOZiDTFUAeHYj3/edit?usp=drive_link&ouid=106090073074900580884&rtpof=true&sd=true',
    'https://docs.google.com/presentation/d/12OF7jJHE_WgEk_mQaCe06Cl4ojNhcAqo/edit?usp=drive_link&ouid=106090073074900580884&rtpof=true&sd=true',
    'https://docs.google.com/presentation/d/1opasIPp6zOpkLwI3gAQC-Br6dbdyDf6c/edit?usp=drive_link&ouid=106090073074900580884&rtpof=true&sd=true',
    'https://docs.google.com/presentation/d/1UrTXzw3pSMDp4aAxyubTnAruMIzuajVl/edit?usp=drive_link&ouid=106090073074900580884&rtpof=true&sd=true',
    'https://docs.google.com/presentation/d/1fw2bB5SUQ_xx37H3jYeiyfDhtaqLXgOD/edit?usp=drive_link&ouid=106090073074900580884&rtpof=true&sd=true'
  ]
};

// 元の関数も保持
function extractPresentationId(url) {
  const match = url.match(/\/presentation\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

function extractTextFromPresentation(presentationId) {
  try {
    const presentation = Slides.Presentations.get(presentationId);
    let allText = '';

    presentation.slides.forEach((slide, index) => {
      allText += `\n=== スライド ${index + 1} ===\n`;

      if (slide.pageElements) {
        slide.pageElements.forEach(element => {
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

  // 景品名の抽出
  const prizePatterns = [
    /(エコバッグ|タンブラー|ボールペン|マグカップ|タオル|キーホルダー|ステッカー)/,
    /(カレンダー|Tシャツ|パーカー|キャップ|トートバッグ|USB|モバイルバッテリー)/,
    /(スマホスタンド|団扇|うちわ|クリアファイル|ノート|メモ帳|ペン|マスク)/
  ];

  for (const pattern of prizePatterns) {
    const match = text.match(pattern);
    if (match) {
      info.prizeName = match[1];
      break;
    }
  }

  // 価格の抽出
  const pricePatterns = [
    /(?:単価|価格|金額)[：:\s]*([¥￥]?)([\d,]+)円?/,
    /([¥￥])([\d,]+)円?(?:\/個|\/枚|\/本)?/,
    /@\s*([¥￥]?)([\d,]+)円?/
  ];

  for (const pattern of pricePatterns) {
    const match = text.match(pattern);
    if (match) {
      const numbers = match.filter(m => m && /^\d/.test(m.replace(/,/g, '')));
      if (numbers.length > 0) {
        info.unitPrice = parseInt(numbers[numbers.length - 1].replace(/,/g, ''));
        break;
      }
    }
  }

  // 数量の抽出
  const quantityMatch = text.match(/(?:数量|ロット|個数)[：:\s]*([\d,]+)\s*(?:個|枚|本|セット)?/);
  if (quantityMatch) {
    info.quantity = parseInt(quantityMatch[1].replace(/,/g, ''));
  }

  // 協力会社の抽出
  const vendorMatch = text.match(/([^\s\n]+(?:株式会社|有限会社|印刷|製作所)[^\s\n]*)/);
  if (vendorMatch) {
    info.vendor = vendorMatch[1];
  }

  // 信頼度計算
  let score = 0;
  if (info.clientName) score += 25;
  if (info.prizeName) score += 25;
  if (info.unitPrice) score += 20;
  if (info.vendor) score += 15;
  if (info.quantity) score += 10;
  info.confidence = Math.min(score, 100);

  return info;
}