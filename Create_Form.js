/**
 * Google Form自動作成スクリプト
 *
 * 【使用方法】
 * 1. https://script.google.com/ にアクセス
 * 2. 「新しいプロジェクト」をクリック
 * 3. このコードをコピペ
 * 4. createKnowledgeForm() を実行
 * 5. 実行ログにForm URLが表示される
 */

function createKnowledgeForm() {
  // Formを作成
  const form = FormApp.create('ナレッジ共有基盤 - 案件情報登録');

  form.setDescription(
    'PowerPoint自動抽出または手動入力で案件情報を登録できます。\n' +
    'Gemini AIが自動で情報を抽出し、NotebookLM用のMarkdownファイルを生成します。'
  );

  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true);

  Logger.log('📝 Creating form questions...');

  // ========================================
  // 質問1: データ入力方法（必須）
  // ========================================
  const inputMethodItem = form.addMultipleChoiceItem();
  inputMethodItem.setTitle('データ入力方法を選択してください')
    .setChoices([
      inputMethodItem.createChoice('PowerPointから自動抽出'),
      inputMethodItem.createChoice('手動で情報を入力')
    ])
    .setRequired(true);

  // ========================================
  // セクション2-A: PowerPoint自動抽出
  // ========================================
  const pptSection = form.addPageBreakItem()
    .setTitle('PowerPoint自動抽出')
    .setHelpText('Google DriveにアップロードしたPowerPointのリンクを貼り付けてください');

  // 質問2: PowerPointリンク
  const pptLinkItem = form.addTextItem();
  pptLinkItem.setTitle('PowerPointファイルのGoogle Driveリンク')
    .setHelpText(
      '【リンクの取得方法】\n' +
      '1. Google DriveでPowerPointファイルを右クリック\n' +
      '2. 「共有」→「リンクをコピー」\n' +
      '3. ここに貼り付け\n\n' +
      '例: https://drive.google.com/file/d/1ABC...XYZ/view'
    )
    .setRequired(true);

  // URL検証を追加
  const pptValidation = FormApp.createTextValidation()
    .requireTextContainsPattern('https://drive.google.com/file/d/')
    .setHelpText('Google DriveのファイルリンクのURLを入力してください')
    .build();
  pptLinkItem.setValidation(pptValidation);

  // 質問3: 補足情報
  const pptSupplementItem = form.addParagraphTextItem();
  pptSupplementItem.setTitle('補足情報（任意）')
    .setHelpText('PowerPointに含まれていない情報や、特に抽出してほしい情報があれば記載してください')
    .setRequired(false);

  // PowerPointセクションから次のページへ（メールアドレス）
  const pptToEmail = form.addPageBreakItem()
    .setTitle('通知設定');

  // ========================================
  // セクション2-B: 手動入力
  // ========================================
  const manualSection = form.addPageBreakItem()
    .setTitle('手動入力')
    .setHelpText('案件情報を入力してください');

  // 質問4: クライアント名（必須）
  const clientNameItem = form.addTextItem();
  clientNameItem.setTitle('クライアント名')
    .setHelpText('「様」は不要です　例: 広研')
    .setRequired(true);

  // 質問5: イベント種別（必須）
  const eventTypeItem = form.addMultipleChoiceItem();
  eventTypeItem.setTitle('イベント種別')
    .setChoices([
      eventTypeItem.createChoice('提案書'),
      eventTypeItem.createChoice('運営マニュアル'),
      eventTypeItem.createChoice('進行台本'),
      eventTypeItem.createChoice('企画書'),
      eventTypeItem.createChoice('キャンペーン'),
      eventTypeItem.createChoice('イベント'),
      eventTypeItem.createChoice('展示会'),
      eventTypeItem.createChoice('セミナー'),
      eventTypeItem.createChoice('その他')
    ])
    .setRequired(true);

  // 質問6: 実施時期（任意）
  const eventDateItem = form.addTextItem();
  eventDateItem.setTitle('実施時期')
    .setHelpText('例: 2025年3月、2025/03/15')
    .setRequired(false);

  // 質問7: イベント内容（任意）
  const eventDescItem = form.addParagraphTextItem();
  eventDescItem.setTitle('イベント内容')
    .setHelpText('イベントの概要を1-2文で記載してください')
    .setRequired(false);

  // 質問8: 会場（任意）
  const venueItem = form.addTextItem();
  venueItem.setTitle('会場')
    .setHelpText('例: 大阪城ホール')
    .setRequired(false);

  // 質問9: ターゲット人数（任意）
  const targetCountItem = form.addTextItem();
  targetCountItem.setTitle('ターゲット人数')
    .setHelpText('数字のみ入力　例: 500')
    .setRequired(false);

  // 質問10: 単価（任意）
  const unitPriceItem = form.addTextItem();
  unitPriceItem.setTitle('単価')
    .setHelpText('円単位で数字のみ入力　例: 1000')
    .setRequired(false);

  // 質問11: 総費用（任意）
  const totalCostItem = form.addTextItem();
  totalCostItem.setTitle('総費用')
    .setHelpText('円単位で数字のみ入力　例: 500000')
    .setRequired(false);

  // 質問12: 発注数量（任意）
  const orderQtyItem = form.addTextItem();
  orderQtyItem.setTitle('発注数量')
    .setHelpText('数字のみ入力　例: 1000')
    .setRequired(false);

  // 質問13: 納期（任意）
  const deadlineItem = form.addTextItem();
  deadlineItem.setTitle('納期')
    .setHelpText('例: 14営業日、2025年3月末')
    .setRequired(false);

  // 質問14: 協力会社（任意）
  const partnersItem = form.addTextItem();
  partnersItem.setTitle('協力会社')
    .setHelpText('複数ある場合はカンマ区切り　例: A社,B社,C社')
    .setRequired(false);

  // 質問15: ノベルティ・景品（任意）
  const noveltyItem = form.addTextItem();
  noveltyItem.setTitle('ノベルティ・景品')
    .setHelpText('複数ある場合はカンマ区切り　例: エコバッグ,ボールペン,クリアファイル')
    .setRequired(false);

  // 質問16: キーワード・タグ（任意）
  const keywordsItem = form.addTextItem();
  keywordsItem.setTitle('キーワード・タグ')
    .setHelpText('検索用のキーワードをカンマ区切り　例: 競馬,ファミリー向け,夏イベント')
    .setRequired(false);

  // 手動入力セクションから次のページへ（メールアドレス）
  const manualToEmail = form.addPageBreakItem()
    .setTitle('通知設定');

  // ========================================
  // セクション3: 共通項目（メールアドレス）
  // ========================================

  // 質問17: 通知先メールアドレス（必須）
  const emailItem = form.addTextItem();
  emailItem.setTitle('通知先メールアドレス')
    .setHelpText('Markdown生成完了の通知を受け取るメールアドレスを入力してください')
    .setRequired(true);

  // メール検証を追加
  const emailValidation = FormApp.createTextValidation()
    .requireTextIsEmail()
    .setHelpText('有効なメールアドレスを入力してください')
    .build();
  emailItem.setValidation(emailValidation);

  // ========================================
  // 条件分岐の設定
  // ========================================
  Logger.log('🔀 Setting up conditional logic...');

  // 質問1の条件分岐
  inputMethodItem.setChoices([
    inputMethodItem.createChoice('PowerPointから自動抽出', pptSection),
    inputMethodItem.createChoice('手動で情報を入力', manualSection)
  ]);

  // PowerPointセクション → メールアドレス
  pptToEmail.setGoToPage(FormApp.PageNavigationType.SUBMIT);

  // 手動入力セクション → メールアドレス
  manualToEmail.setGoToPage(FormApp.PageNavigationType.SUBMIT);

  // ========================================
  // 確認メッセージ
  // ========================================
  form.setConfirmationMessage(
    '送信ありがとうございます！\n\n' +
    'Markdownファイルの生成を開始しました。\n' +
    '完了次第、メールで通知いたします。\n\n' +
    '処理には10-30秒程度かかります。'
  );

  // ========================================
  // Form URLを取得
  // ========================================
  const formUrl = form.getPublishedUrl();
  const editUrl = form.getEditUrl();

  Logger.log('\n✅ Form created successfully!');
  Logger.log('\n📋 Form URL (share this with users):');
  Logger.log(formUrl);
  Logger.log('\n⚙️ Edit URL (for you):');
  Logger.log(editUrl);
  Logger.log('\n🔗 Form ID:');
  Logger.log(form.getId());

  // スプレッドシートにリンク（回答を記録）
  const spreadsheet = SpreadsheetApp.create('ナレッジ共有基盤 - 回答記録');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  Logger.log('\n📊 Spreadsheet URL:');
  Logger.log(spreadsheet.getUrl());

  // 次のステップを表示
  Logger.log('\n📝 Next steps:');
  Logger.log('1. Open the Edit URL above');
  Logger.log('2. Go to "Script editor" (⋮ menu → Script editor)');
  Logger.log('3. Copy the content of Integrated_Form_to_Markdown.js');
  Logger.log('4. Set up Script Properties (GEMINI_API_KEY, OUTPUT_FOLDER_ID)');
  Logger.log('5. Set up trigger (onFormSubmit on form submit)');

  return {
    formUrl: formUrl,
    editUrl: editUrl,
    formId: form.getId(),
    spreadsheetUrl: spreadsheet.getUrl()
  };
}

/**
 * テスト用: 既存のFormを削除
 * ※注意: 実行すると復元できません
 */
function deleteTestForm(formId) {
  const form = FormApp.openById(formId);
  DriveApp.getFileById(formId).setTrashed(true);
  Logger.log('Form deleted');
}
