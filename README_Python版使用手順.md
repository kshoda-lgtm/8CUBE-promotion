# PowerPoint解析システム（Python + Google Apps Script版）

## 🎯 システム概要

**完全無料**でPowerPointファイルからナレッジ情報を自動抽出し、Google Spreadsheetで管理するシステムです。

### 📋 処理フロー

```
PowerPoint(.pptx) → [Python処理] → JSON → [Google Apps Script] → Spreadsheet
```

---

## ⚙️ セットアップ

### 1. Python環境の準備

```bash
# python-pptxライブラリのインストール
pip install python-pptx
```

### 2. ファイル配置

```
promotion部/
├── powerpoint_processor.py    # PowerPoint解析スクリプト
├── JSON_processor.js          # Google Apps Script用JSON処理
└── README_Python版使用手順.md  # この手順書
```

---

## 🚀 使用方法

### Step 1: PowerPointファイルの処理（ローカル）

1. **コマンドプロンプトまたはターミナルを開く**

2. **Pythonスクリプトを実行**
   ```bash
   cd "C:\Users\owner\OneDrive\Desktop\vexum\promotion部"
   python powerpoint_processor.py
   ```

3. **PowerPointファイルのパスを入力**
   ```
   PowerPointファイルのパスを入力してください: C:\path\to\your\presentation.pptx
   ```

4. **処理結果の確認**
   - 同じフォルダに `presentation.json` が生成される
   - コンソールに抽出結果のサマリーが表示される

### Step 2: JSONファイルをGoogle Driveにアップロード

1. **生成されたJSONファイルをGoogle Driveにアップロード**
   - ファイル名に "powerpoint" または "pptx" を含める
   - 例: `presentation_powerpoint.json`

### Step 3: Google Apps ScriptでJSONを処理

1. **Google Apps Scriptを開く**
   ```
   https://script.google.com/home/projects/1TDH5umBKYshlbyEuvy3EZNCnLL4IbjH-5MmEI7Yv3TGmTHNM9kDXuiJz/edit
   ```

2. **JSON_processor.js の関数を実行**
   - `processJSONFiles()` を実行
   - または `testJSONProcessing()` でテスト

3. **結果確認**
   - スプレッドシートにデータが自動追加される
   - URL: https://docs.google.com/spreadsheets/d/1LVjGOulUFlrsq1TOwZR3hEXO_c4DSDPM0UkQ9doBKMw

---

## 📊 抽出可能な情報

### 自動抽出項目

| 項目 | 抽出方法 | 例 |
|------|----------|-----|
| **価格** | 正規表現パターン | "単価500円" → 500 |
| **数量** | 数値パターン | "1000個" → 1000 |
| **納期** | 期間パターン | "14営業日" → "14営業日" |
| **会社名** | 会社パターン | "株式会社ABC" → "ABC" |
| **日付** | 日付パターン | "2024年7月" → "2024/07/01" |
| **イベント種別** | キーワード判定 | "キャンペーン" → "キャンペーン" |
| **キーワード** | 重要語抽出 | "ノベルティ", "エコ" など |

### 生成されるデータ

- **信頼度スコア**: 抽出精度の指標（0-100%）
- **タグ**: 自動生成されたキーワード群
- **サマリー情報**: 価格帯、数量、関連会社一覧

---

## 🔧 カスタマイズ

### パターンマッチングの調整

`powerpoint_processor.py` の `patterns` 辞書を編集：

```python
self.patterns = {
    'price': r'(?:単価|価格|¥|円).*?(\d{1,3}(?:,\d{3})*(?:\.\d+)?)',
    'quantity': r'(?:数量|個数|枚数|ロット).*?(\d{1,3}(?:,\d{3})*)',
    # 新しいパターンを追加
    'custom_pattern': r'your_regex_here'
}
```

### キーワード辞書の更新

```python
important_words = [
    'ノベルティ', '景品', 'グッズ', 'プレゼント',
    # 新しいキーワードを追加
    'あなたの業界用語'
]
```

---

## 🚨 トラブルシューティング

### よくある問題と解決法

#### 1. `python-pptx` インストールエラー
```bash
# 管理者権限で実行
pip install --user python-pptx
```

#### 2. PowerPointファイルが開けない
- ファイルが破損していないか確認
- `.pptx` 形式であることを確認（`.ppt` は非対応）

#### 3. JSONファイルが見つからない
- Google Driveでファイル名に "powerpoint" または "pptx" が含まれているか確認
- ファイル形式が `application/json` になっているか確認

#### 4. スプレッドシートに書き込めない
- Google Apps Scriptの権限を確認
- スプレッドシートIDが正しいか確認

---

## 📈 システムの利点

### ✅ 完全無料
- Python: 無料
- Google Workspace: 既存利用
- 追加ツール: 不要

### ✅ 大きなファイル対応
- ローカルで処理するためサイズ制限なし
- Google Apps Scriptの6分制限回避

### ✅ 高精度抽出
- python-pptxによる正確なテキスト抽出
- テーブル、図形内テキストも対応

### ✅ カスタマイズ可能
- 正規表現パターンの調整
- キーワード辞書の拡張
- 出力フォーマットの変更

---

## 🔄 バッチ処理（複数ファイル対応）

### 複数ファイルを一括処理

```python
# powerpoint_processor.py に追加可能な機能
import os

def batch_process(folder_path):
    """フォルダ内の全PowerPointファイルを処理"""
    for filename in os.listdir(folder_path):
        if filename.endswith('.pptx'):
            file_path = os.path.join(folder_path, filename)
            print(f"処理中: {filename}")
            processor.process_powerpoint(file_path)
```

---

## 📞 サポート

### 技術サポート
- 📧 Email: ses.members@8-cube.co.jp
- 💻 Google Apps Script: [プロジェクトURL](https://script.google.com/home/projects/1TDH5umBKYshlbyEuvy3EZNCnLL4IbjH-5MmEI7Yv3TGmTHNM9kDXuiJz/edit)
- 📊 データベース: [スプレッドシートURL](https://docs.google.com/spreadsheets/d/1LVjGOulUFlrsq1TOwZR3hEXO_c4DSDPM0UkQ9doBKMw)

### 更新履歴
- v1.0 (2025/01/29): Python版初回リリース

---

*Copyright © 2025 株式会社エイトキューブ プロモーション事業部*