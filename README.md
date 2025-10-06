# ナレッジ共有基盤システム v4.0

**NotebookLM連携型 ハイブリッドナレッジ収集システム**

---

## 🎯 概要

PowerPointファイルとGoogle Formの両方からデータを収集し、NotebookLMで過去のナレッジを参照できるシステムです。

### 核心価値

- ✅ **完全無料** - 月30,000ファイルまで無料
- ✅ **AI自動抽出** - 95%以上の精度
- ✅ **手入力対応** - Formでも情報収集
- ✅ **課金防止** - 無料枠超過時は自動停止
- ✅ **NotebookLM連携** - AIが過去案件を自動参照

---

## 📁 ファイル一覧

### Python スクリプト（ローカル実行）

| ファイル | 用途 |
|---------|------|
| `markdown_generator.py` | PowerPoint 1ファイル → Markdown変換 |
| `batch_markdown_generator.py` | PowerPoint複数ファイル → 一括Markdown変換 |
| `powerpoint_processor_gemini.py` | Gemini API処理エンジン |
| `batch_process_gemini.py` | バッチJSON生成 |

### Google Apps Script（クラウド実行）

| ファイル | 用途 |
|---------|------|
| `Form_to_Markdown.js` | Google Form回答 → Markdown自動変換 |
| `JSON_processor.js` | JSON → スプレッドシート登録（オプション） |

### ドキュメント

| ファイル | 内容 |
|---------|------|
| `ナレッジ共有基盤_無料版仕様書.md` | システム仕様書（詳細版） |
| `Google_Form設計.md` | フォーム設計仕様 |
| `README.md` | このファイル |

---

## 🚀 クイックスタート

### 1. PowerPointからMarkdown生成

```bash
# 環境準備
pip install python-pptx google-generativeai

# Gemini APIキー設定
set GEMINI_API_KEY=your-api-key-here

# 単一ファイル変換
cd promotion部
python markdown_generator.py "path/to/file.pptx"

# バッチ変換
python batch_markdown_generator.py "AIマニュアル化\AIマニュアル化"
```

### 2. Google Form設定

1. https://forms.google.com/ でフォーム作成
2. `Google_Form設計.md` の仕様に従って質問を設定
3. スクリプトエディタで `Form_to_Markdown.js` を設定
4. トリガー設定（フォーム送信時）

### 3. NotebookLMにアップロード

1. https://notebooklm.google.com/ にアクセス
2. 新しいノートブックを作成
3. 生成された `.md` ファイルをすべてアップロード
4. 完了！AIに質問できます

---

## 💡 使用例

### NotebookLMへの質問例

```
「広研様に似た案件はありますか？」
→ 過去の類似案件を自動検索

「エコバッグの過去の単価を教えて」
→ 価格相場を即座に把握

「A社との協力実績を教えて」
→ 協力会社の評価を参照

「2024年10月に実施したイベントは？」
→ 時期別の案件検索

「過去にうまくいった成功要因は？」
→ 成功パターンの学習

「よくある失敗パターンは？」
→ 失敗を未然に防ぐ
```

---

## 📊 データ収集の2つのルート

### ルート1: PowerPoint（自動）

**対象:**
- 既存の企画書・提案書・運営マニュアル

**処理:**
- Gemini APIが自動で情報抽出
- 処理時間: 1ファイル5-10秒
- 精度: 95%以上

**手順:**
```bash
python markdown_generator.py "file.pptx"
```

### ルート2: Google Form（手入力）

**対象:**
- PowerPointがない案件
- 口頭で進んだ案件
- 成功・失敗事例の共有

**処理:**
- 職員がフォームに入力
- Apps Scriptが自動でMarkdown生成
- 処理時間: 入力時間次第（5-10分）
- 精度: 100%（手入力）

**手順:**
1. フォームURLにアクセス
2. 案件情報を入力
3. 送信 → 自動でMarkdown生成

---

## 🛡️ 無料枠超過防止機能

システムは自動的に無料枠を監視:

- **日次制限**: 1,000回/日
- **月次制限**: 30,000回/月
- **使用状況**: `.gemini_usage.json` で記録
- **超過時**: 自動停止（課金なし）

### 使用状況の確認

起動時に自動表示:
```
📊 FREE TIER USAGE STATUS:
   Today: 45/1,000 requests (残り 955)
   This month: 1,234/30,000 requests (残り 28,766)
   Total: 1,234 requests
```

---

## 🔧 トラブルシューティング

### Q1: Gemini APIキーが見つからない

```bash
# 環境変数を設定
set GEMINI_API_KEY=your-api-key-here

# または直接入力
python markdown_generator.py
# プロンプトで入力
```

### Q2: python-pptx がインストールできない

```bash
# Pythonのバージョン確認
python --version

# 再インストール
pip uninstall python-pptx
pip install python-pptx
```

### Q3: Form_to_Markdown.js が動作しない

1. `CONFIG.OUTPUT_FOLDER_ID` を設定したか確認
2. トリガーが正しく設定されているか確認
3. Apps Scriptの実行ログを確認

---

## 📞 サポート

詳細な仕様は `ナレッジ共有基盤_無料版仕様書.md` を参照してください。

---

## 📝 更新履歴

### v4.0 (2025-10-01)
- ✅ NotebookLM連携機能を追加
- ✅ Google Form連携機能を追加
- ✅ ハイブリッド型ナレッジ収集システムに刷新
- ✅ Markdown生成機能を実装
- ✅ 無料枠超過防止機能を実装

### v3.0
- 正規表現ベースの抽出システム

---

**株式会社エイトキューブ プロモーション事業部**
