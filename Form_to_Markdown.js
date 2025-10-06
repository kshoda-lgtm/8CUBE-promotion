/**
 * Google Form回答をNotebookLM用Markdownに自動変換
 * フォーム送信時にトリガーで実行
 */

// ===== 設定 =====
const CONFIG = {
  // Markdownファイルの保存先フォルダID（Google Drive）
  OUTPUT_FOLDER_ID: 'YOUR_FOLDER_ID_HERE', // 要変更

  // NotebookLM用フォルダID（オプション）
  NOTEBOOKLM_FOLDER_ID: 'YOUR_NOTEBOOKLM_FOLDER_ID_HERE' // 要変更
};

/**
 * フォーム送信時のトリガー関数
 * Google Formの「スクリプトエディタ」で設定
 */
function onFormSubmit(e) {
  try {
    console.log('📨 新しいフォーム回答を受信');

    // フォーム回答を取得
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();

    // データを抽出
    const data = extractFormData(itemResponses, formResponse);

    // Markdown生成
    const markdown = generateMarkdownFromForm(data);

    // Markdownファイルを保存
    const fileName = generateFileName(data);
    saveMarkdownToGoogleDrive(fileName, markdown);

    console.log('✅ Markdown生成完了: ' + fileName);

  } catch (error) {
    console.error('❌ エラー:', error);
    // メール通知（オプション）
    // MailApp.sendEmail('your-email@example.com', 'Form処理エラー', error.toString());
  }
}

/**
 * フォームデータを抽出
 */
function extractFormData(itemResponses, formResponse) {
  const data = {
    timestamp: formResponse.getTimestamp(),
    respondentEmail: formResponse.getRespondentEmail()
  };

  // 各質問の回答を取得
  itemResponses.forEach(itemResponse => {
    const title = itemResponse.getItem().getTitle();
    const response = itemResponse.getResponse();

    // 質問タイトルをキーにマッピング
    switch(title) {
      case '案件名':
        data.projectName = response;
        break;
      case 'クライアント名':
        data.clientName = response;
        break;
      case '担当者名':
        data.personInCharge = response;
        break;
      case '実施時期':
        data.eventDate = response;
        break;
      case 'イベント種別':
        data.eventType = response;
        break;
      case 'イベント内容（概要）':
        data.eventDescription = response;
        break;
      case '会場名':
        data.venue = response;
        break;
      case 'ターゲット人数':
        data.targetCount = response;
        break;
      case '単価（円）':
        data.unitPrice = response;
        break;
      case '総費用（円）':
        data.totalCost = response;
        break;
      case '発注数量':
        data.orderQuantity = response;
        break;
      case '協力会社名':
        data.partnerCompanies = response;
        break;
      case '協力会社の評価':
        data.partnerEvaluation = response;
        break;
      case 'ノベルティ/景品の種類':
        data.noveltyItems = response;
        break;
      case '納期':
        data.deadline = response;
        break;
      case '成功要因・うまくいった点':
        data.successFactors = response;
        break;
      case '失敗・反省点':
        data.failurePoints = response;
        break;
      case '企画書・資料のURL':
        data.documentUrl = response;
        break;
      case 'タグ・キーワード':
        data.tags = response;
        break;
      case '備考・補足情報':
        data.notes = response;
        break;
    }
  });

  return data;
}

/**
 * NotebookLM用Markdownを生成
 */
function generateMarkdownFromForm(data) {
  const lines = [];

  // タイトル
  const title = data.projectName || `【${data.clientName || '不明'}様】案件`;
  lines.push(`# ${title}\n`);

  // メタ情報
  lines.push(`**登録日時**: ${formatDate(data.timestamp)}`);
  lines.push(`**担当者**: ${data.personInCharge || data.respondentEmail || '不明'}`);
  lines.push(`**データソース**: Google Form（手入力）\n`);
  lines.push('---\n');

  // 基本情報
  lines.push('## 📋 基本情報\n');

  if (data.clientName) {
    lines.push(`- **クライアント名**: ${data.clientName}`);
  }

  if (data.eventDate) {
    lines.push(`- **実施時期**: ${data.eventDate}`);
  }

  if (data.eventType) {
    lines.push(`- **イベント種別**: ${data.eventType}`);
  }

  if (data.venue) {
    lines.push(`- **会場**: ${data.venue}`);
  }

  if (data.targetCount) {
    lines.push(`- **ターゲット人数**: ${data.targetCount}`);
  }

  lines.push('');

  // イベント内容
  if (data.eventDescription) {
    lines.push('## 📝 イベント内容\n');
    lines.push(`${data.eventDescription}\n`);
  }

  // 価格情報
  if (data.unitPrice || data.totalCost || data.orderQuantity) {
    lines.push('## 💰 価格情報\n');

    if (data.unitPrice) {
      lines.push(`- **単価**: ¥${Number(data.unitPrice).toLocaleString()}`);
    }

    if (data.totalCost) {
      lines.push(`- **総費用**: ¥${Number(data.totalCost).toLocaleString()}`);
    }

    if (data.orderQuantity) {
      lines.push(`- **発注数量**: ${Number(data.orderQuantity).toLocaleString()}個`);
    }

    lines.push('');
  }

  // 納期
  if (data.deadline) {
    lines.push('## ⏰ 納期\n');
    lines.push(`- **納期**: ${data.deadline}\n`);
  }

  // 協力会社
  if (data.partnerCompanies) {
    lines.push('## 🤝 協力会社\n');
    const companies = data.partnerCompanies.split('\n').filter(c => c.trim());
    companies.forEach(company => {
      lines.push(`- ${company.trim()}`);
    });
    lines.push('');

    if (data.partnerEvaluation) {
      lines.push(`**評価**: ${data.partnerEvaluation}\n`);
    }
  }

  // ノベルティ/景品
  if (data.noveltyItems) {
    lines.push('## 🎁 ノベルティ/景品\n');
    const items = data.noveltyItems.split('\n').filter(i => i.trim());
    items.forEach(item => {
      lines.push(`- ${item.trim()}`);
    });
    lines.push('');
  }

  // 成功要因
  if (data.successFactors) {
    lines.push('## ✅ 成功要因\n');
    lines.push(`${data.successFactors}\n`);
  }

  // 反省点
  if (data.failurePoints) {
    lines.push('## ⚠️ 反省点\n');
    lines.push(`${data.failurePoints}\n`);
  }

  // タグ・キーワード
  if (data.tags) {
    lines.push('## 🏷️ タグ・キーワード\n');
    const tags = data.tags.split(',').map(t => t.trim()).filter(t => t);
    const tagString = tags.map(t => `\`#${t}\``).join(' ');
    lines.push(`${tagString}\n`);
  }

  // 参考資料
  if (data.documentUrl) {
    lines.push('## 📎 参考資料\n');
    lines.push(`- [企画書・資料リンク](${data.documentUrl})\n`);
  }

  // 備考
  if (data.notes) {
    lines.push('## 📌 備考\n');
    lines.push(`${data.notes}\n`);
  }

  // フッター
  lines.push('---\n');
  lines.push(`*登録者: ${data.personInCharge || data.respondentEmail} | 登録日: ${formatDate(data.timestamp)}*`);

  return lines.join('\n');
}

/**
 * ファイル名を生成
 */
function generateFileName(data) {
  const date = formatDateShort(data.timestamp);
  const client = data.clientName ? `【${data.clientName}様】` : '';
  const project = data.projectName || data.eventType || 'プロモーション案件';

  // ファイル名に使えない文字を削除
  const safeName = `${date}_${client}${project}`.replace(/[\/\\:*?"<>|]/g, '_');

  return `${safeName}.md`;
}

/**
 * MarkdownをGoogle Driveに保存
 */
function saveMarkdownToGoogleDrive(fileName, markdownContent) {
  try {
    // フォルダを取得
    const folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);

    // Markdownファイルを作成
    const file = folder.createFile(fileName, markdownContent, MimeType.PLAIN_TEXT);

    console.log(`✅ 保存完了: ${file.getUrl()}`);

    // NotebookLM用フォルダにもコピー（オプション）
    if (CONFIG.NOTEBOOKLM_FOLDER_ID) {
      const notebookLMFolder = DriveApp.getFolderById(CONFIG.NOTEBOOKLM_FOLDER_ID);
      file.makeCopy(fileName, notebookLMFolder);
      console.log('📋 NotebookLM用フォルダにもコピー完了');
    }

    return file;

  } catch (error) {
    console.error('❌ ファイル保存エラー:', error);
    throw error;
  }
}

/**
 * 日付フォーマット（詳細）
 */
function formatDate(date) {
  if (!date) return '';
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

/**
 * 日付フォーマット（短縮）
 */
function formatDateShort(date) {
  if (!date) return '';
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');
}

/**
 * テスト用関数（手動実行）
 */
function testFormToMarkdown() {
  console.log('🧪 テスト実行中...');

  // サンプルデータ
  const testData = {
    timestamp: new Date(),
    respondentEmail: 'test@example.com',
    projectName: '洛北阪急スクエアイベント',
    clientName: '広研',
    personInCharge: '山田太郎',
    eventDate: '2024/10/15',
    eventType: 'イベント',
    eventDescription: '洛北阪急スクエアでのハロウィンイベント。子供向けワークショップとノベルティ配布を実施。',
    venue: '洛北阪急スクエア',
    targetCount: '先着500名',
    unitPrice: '500',
    totalCost: '300000',
    orderQuantity: '1000',
    partnerCompanies: '株式会社A印刷\nB製作所\nC企画',
    partnerEvaluation: 'A印刷: 納期厳守で高品質。また依頼したい。',
    noveltyItems: 'オリジナルエコバッグ\nクリアファイル\nボールペン',
    deadline: '14営業日',
    successFactors: 'エコバッグのデザインが好評。SNS拡散効果が高かった。',
    failurePoints: '数量が不足して追加発注が必要になった。余裕を持った発注が必要。',
    tags: 'エコ, ハロウィン, 子供向け, ワークショップ'
  };

  // Markdown生成
  const markdown = generateMarkdownFromForm(testData);
  console.log('生成されたMarkdown:\n');
  console.log(markdown);

  // ファイル保存（コメントアウトを解除して実行）
  // const fileName = generateFileName(testData);
  // saveMarkdownToGoogleDrive(fileName, markdown);

  console.log('✅ テスト完了');
}
