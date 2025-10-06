# 統合Google Form設計書（パターンA: 無料版）

**バージョン**: v4.0
**最終更新**: 2025-10-06

---

## 🎯 システム概要

### 職員さんの操作フロー

```
1. Google Formにアクセス
   ↓
2. 入力方法を選択
   ○ PowerPointから自動抽出
   ○ 手動で情報入力
   ↓
3. 必要な情報を入力/リンクを貼り付け
   ↓
4. 送信ボタンをクリック
   ↓
【自動処理】
   ↓
5. メールで通知を受け取る
   ↓
6. Google DriveからMarkdownをダウンロード
   ↓
7. NotebookLMに手動アップロード（3秒）
```

---

## 📝 Form質問項目設計

### セクション1: 入力方法の選択

**質問1: データ入力方法**
- タイプ: ラジオボタン（必須）
- 選択肢:
  - `PowerPointから自動抽出`
  - `手動で情報を入力`

---

### セクション2-A: PowerPoint自動抽出（条件分岐）

**※ 「PowerPointから自動抽出」を選択した場合のみ表示**

**質問2: PowerPointファイルのGoogle Driveリンク**
- タイプ: 短い回答（必須）
- 説明文:
  ```
  Google DriveにアップロードしたPowerPointファイルのリンクを貼り付けてください。

  【リンクの取得方法】
  1. Google DriveでPowerPointファイルを右クリック
  2. 「リンクを取得」をクリック
  3. 「リンクをコピー」をクリック
  4. ここに貼り付け

  例: https://drive.google.com/file/d/1ABC...XYZ/view?usp=sharing
  ```
- 検証: URLパターン `https://drive.google.com/file/d/`

**質問3: 補足情報（任意）**
- タイプ: 段落
- 説明文:
  ```
  PowerPointに含まれていない情報や、
  特に抽出してほしい情報があれば記載してください（任意）
  ```

---

### セクション2-B: 手動入力（条件分岐）

**※ 「手動で情報を入力」を選択した場合のみ表示**

**質問4: クライアント名**
- タイプ: 短い回答（必須）
- 説明: 「様」は不要です
- 例: 広研

**質問5: イベント種別**
- タイプ: プルダウン（必須）
- 選択肢:
  - 提案書
  - 運営マニュアル
  - 進行台本
  - 企画書
  - キャンペーン
  - イベント
  - 展示会
  - セミナー
  - その他

**質問6: 実施時期**
- タイプ: 短い回答（任意）
- 例: 2025年3月、2025/03/15

**質問7: イベント内容**
- タイプ: 段落（任意）
- 説明: イベントの概要を1-2文で記載

**質問8: 会場**
- タイプ: 短い回答（任意）
- 例: 大阪城ホール

**質問9: ターゲット人数**
- タイプ: 短い回答（任意）
- 例: 500

**質問10: 単価**
- タイプ: 短い回答（任意）
- 説明: 円単位で数字のみ入力
- 例: 1000

**質問11: 総費用**
- タイプ: 短い回答（任意）
- 説明: 円単位で数字のみ入力
- 例: 500000

**質問12: 発注数量**
- タイプ: 短い回答（任意）
- 例: 1000

**質問13: 納期**
- タイプ: 短い回答（任意）
- 例: 14営業日、2025年3月末

**質問14: 協力会社**
- タイプ: 短い回答（任意）
- 説明: 複数ある場合はカンマ区切り
- 例: A社,B社,C社

**質問15: ノベルティ・景品**
- タイプ: 短い回答（任意）
- 説明: 複数ある場合はカンマ区切り
- 例: エコバッグ,ボールペン,クリアファイル

**質問16: キーワード・タグ**
- タイプ: 短い回答（任意）
- 説明: 検索用のキーワードをカンマ区切り
- 例: 競馬,ファミリー向け,夏イベント

---

### セクション3: 共通項目

**質問17: 通知先メールアドレス**
- タイプ: 短い回答（必須）
- 検証: メールアドレス形式
- 説明: Markdown生成完了の通知を受け取るメールアドレス

---

## 🔧 Apps Script実装仕様

### 設定値（スクリプトプロパティ）

```javascript
// スクリプトプロパティで設定
GEMINI_API_KEY: "AIzaSy..." // Gemini APIキー
OUTPUT_FOLDER_ID: "1ABC...XYZ" // Markdown出力先フォルダID
```

### トリガー設定

- **イベント**: フォーム送信時
- **実行する関数**: `onFormSubmit`

### 処理フロー

```javascript
function onFormSubmit(e) {
  // 1. フォーム回答を取得
  var responses = e.response.getItemResponses();
  var inputMethod = responses[0].getResponse(); // PowerPoint or 手動

  if (inputMethod === "PowerPointから自動抽出") {
    // 2-A. PowerPoint処理
    var driveLink = responses[1].getResponse();
    var fileId = extractFileIdFromLink(driveLink);

    // 3. Slides APIでテキスト抽出
    var slideTexts = extractTextFromSlides(fileId);

    // 4. Gemini APIで解析
    var analysisResult = analyzeWithGemini(slideTexts, fileName);

  } else {
    // 2-B. 手動入力処理
    var analysisResult = {
      client_name: responses[3].getResponse(),
      event_type: responses[4].getResponse(),
      event_date: responses[5].getResponse(),
      // ... その他の項目
    };
  }

  // 5. Markdownファイル生成
  var markdown = generateMarkdown(analysisResult);

  // 6. Google Driveに保存
  var file = saveMarkdownToDrive(markdown, analysisResult.client_name);

  // 7. メール通知
  var email = responses[responses.length - 1].getResponse();
  sendNotificationEmail(email, file);
}
```

### 主要関数

#### 1. PowerPointテキスト抽出

```javascript
function extractTextFromSlides(fileId) {
  var presentation = SlidesApp.openById(fileId);
  var slides = presentation.getSlides();
  var allTexts = [];

  slides.forEach(function(slide) {
    var shapes = slide.getShapes();
    shapes.forEach(function(shape) {
      try {
        var text = shape.getText().asString();
        if (text.trim()) {
          allTexts.push(text.trim());
        }
      } catch (e) {
        // テキストがない図形はスキップ
      }
    });

    // テーブルも処理
    var tables = slide.getTables();
    tables.forEach(function(table) {
      // テーブルのテキストを抽出
    });
  });

  return allTexts;
}
```

#### 2. Gemini API呼び出し

```javascript
function analyzeWithGemini(slideTexts, fileName) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var endpoint = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent';

  var prompt = `あなたはプロモーション事業のデータ分析AIです。
以下のPowerPointスライドのテキストから、構造化データを抽出してください。

【ファイル名】
${fileName}

【抽出項目】
1. client_name: クライアント名
2. event_date: 実施時期
3. event_type: イベント種別
... (省略)

【スライドテキスト】
${slideTexts.join('\n\n')}

【出力形式】
JSON形式で出力してください。
`;

  var payload = {
    "contents": [{
      "parts": [{
        "text": prompt
      }]
    }]
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(endpoint + '?key=' + apiKey, options);
  var json = JSON.parse(response.getContentText());

  // レスポンスからJSONを抽出してパース
  var text = json.candidates[0].content.parts[0].text;
  var jsonMatch = text.match(/```json\s*([\s\S]*?)\s*```/);
  var jsonStr = jsonMatch ? jsonMatch[1] : text;

  return JSON.parse(jsonStr);
}
```

#### 3. Markdown生成

```javascript
function generateMarkdown(data) {
  var md = [];

  // ヘッダー
  md.push(`# 【${data.client_name || 'Unknown'}様】${data.event_type || 'プロモーション案件'}\n`);
  md.push(`**処理日時**: ${new Date().toISOString()}`);
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
    if (data.unit_price) md.push(`- **単価**: ¥${Number(data.unit_price).toLocaleString()}`);
    if (data.total_cost) md.push(`- **総費用**: ¥${Number(data.total_cost).toLocaleString()}`);
    if (data.order_quantity) md.push(`- **発注数量**: ${Number(data.order_quantity).toLocaleString()}個`);
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
    data.partner_companies.forEach(function(company) {
      md.push(`- ${company}`);
    });
    md.push('');
  }

  // ノベルティ
  if (data.novelty_items && data.novelty_items.length > 0) {
    md.push('## 🎁 ノベルティ/景品\n');
    data.novelty_items.forEach(function(item) {
      md.push(`- ${item}`);
    });
    md.push('');
  }

  // キーワード
  if (data.keywords && data.keywords.length > 0) {
    md.push('## 🏷️ タグ・キーワード\n');
    var tags = data.keywords.map(function(kw) { return '#' + kw; }).join(' ');
    md.push(tags + '\n');
  }

  // フッター
  md.push('---');
  md.push(`\n*Generated by NotebookLM Knowledge System v4.0 - ${new Date().toISOString()}*`);

  return md.join('\n');
}
```

#### 4. Google Driveに保存

```javascript
function saveMarkdownToDrive(markdown, clientName) {
  var folderId = PropertiesService.getScriptProperties().getProperty('OUTPUT_FOLDER_ID');
  var folder = DriveApp.getFolderById(folderId);

  var timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HHmmss');
  var fileName = `${timestamp}_${clientName || 'Unknown'}.md`;

  var file = folder.createFile(fileName, markdown, MimeType.PLAIN_TEXT);

  return file;
}
```

#### 5. メール通知

```javascript
function sendNotificationEmail(email, file) {
  var subject = '【ナレッジ共有基盤】Markdown生成完了';
  var body = `
Markdownファイルの生成が完了しました。

ファイル名: ${file.getName()}
ダウンロード: ${file.getUrl()}

【次のステップ】
1. 上記リンクからMarkdownファイルをダウンロード
2. NotebookLMにアクセス: https://notebooklm.google.com/
3. Markdownファイルをアップロード（ドラッグ&ドロップ）

完了です！

---
株式会社エイトキューブ プロモーション事業部
ナレッジ共有基盤システム v4.0
`;

  MailApp.sendEmail(email, subject, body);
}
```

---

## 📊 処理時間の目安

- PowerPoint自動抽出: 10-30秒
- 手動入力: 3-5秒
- メール通知: 即時

---

## 🔐 セキュリティ設定

### スクリプトプロパティ

```
GEMINI_API_KEY: クライアントのAPIキー（絶対に公開しない）
OUTPUT_FOLDER_ID: Markdown出力先フォルダID
```

### 実行権限

- **Form送信者**: 実行権限不要（トリガーが自動実行）
- **Apps Script**: 以下の権限が必要
  - Google Drive API
  - Google Slides API
  - Gmail API（メール送信用）

---

## ✅ 実装チェックリスト

- [ ] Google Formを作成
- [ ] 条件分岐を設定（PowerPoint/手動入力）
- [ ] Apps Scriptをフォームにリンク
- [ ] スクリプトプロパティを設定
- [ ] Gemini APIキーを設定
- [ ] 出力先フォルダを作成
- [ ] トリガーを設定（フォーム送信時）
- [ ] テスト実行（PowerPoint）
- [ ] テスト実行（手動入力）
- [ ] メール通知のテスト

---

## 📞 次のステップ

1. Google Formを実際に作成
2. Apps Scriptコードを実装
3. テスト実行
4. 職員さんに使い方を案内

---

**株式会社エイトキューブ プロモーション事業部**
**ナレッジ共有基盤システム v4.0**
