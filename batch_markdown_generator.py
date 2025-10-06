#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NotebookLM用バッチMarkdown生成スクリプト
フォルダ内の全PowerPointファイルを一括でMarkdown変換
"""

import os
import sys
from pathlib import Path
from markdown_generator import process_powerpoint_to_markdown
import time


def batch_generate_markdown(folder_path: str, api_key: str):
    """フォルダ内の全PowerPointファイルをMarkdownに変換"""

    print(f"\n{'='*60}")
    print(f"NotebookLM用バッチMarkdown生成")
    print(f"{'='*60}\n")

    # .pptxファイルを再帰的に検索
    pptx_files = list(Path(folder_path).rglob("*.pptx"))

    if not pptx_files:
        print(f"❌ No .pptx files found in: {folder_path}")
        return

    print(f"📁 Found {len(pptx_files)} PowerPoint files\n")

    success_count = 0
    error_count = 0
    results = []

    for i, pptx_file in enumerate(pptx_files, 1):
        print(f"\n[{i}/{len(pptx_files)}] {pptx_file.name}")
        print("-" * 60)

        try:
            # Markdown生成
            process_powerpoint_to_markdown(str(pptx_file), api_key)

            success_count += 1
            results.append({
                'file': pptx_file.name,
                'status': 'success',
                'markdown': pptx_file.with_suffix('.md').name
            })

            # APIレート制限を考慮して待機
            if i < len(pptx_files):
                time.sleep(2)

        except Exception as e:
            print(f"❌ ERROR: {str(e)}")
            error_count += 1
            results.append({
                'file': pptx_file.name,
                'status': 'error',
                'error': str(e)
            })

    # サマリー出力
    print(f"\n{'='*60}")
    print(f"バッチ処理完了")
    print(f"{'='*60}")
    print(f"✅ 成功: {success_count}/{len(pptx_files)}")
    print(f"❌ エラー: {error_count}/{len(pptx_files)}")

    if success_count > 0:
        print(f"\n📝 生成されたMarkdownファイル:")
        for result in results:
            if result['status'] == 'success':
                print(f"   - {result['markdown']}")

        print(f"\n💡 次のステップ:")
        print(f"   1. https://notebooklm.google.com/ にアクセス")
        print(f"   2. 新しいノートブックを作成")
        print(f"   3. 上記のMarkdownファイルをすべてアップロード")
        print(f"   4. 完了！AIに質問できます")

    print(f"{'='*60}\n")


def main():
    """メイン処理"""
    print("="*60)
    print("NotebookLM用バッチMarkdown生成ツール")
    print("="*60)

    # APIキーの確認
    api_key = os.environ.get('GEMINI_API_KEY')
    if not api_key:
        print("\nGemini API key not found in environment variable.")
        api_key = input("Enter your Gemini API key: ").strip()
        if not api_key:
            print("ERROR: API key is required")
            return

    # フォルダパスを取得
    if len(sys.argv) > 1:
        folder = sys.argv[1]
    else:
        folder = input("\nEnter folder path to process: ").strip()

    if not Path(folder).exists():
        print(f"ERROR: Folder not found: {folder}")
        sys.exit(1)

    # バッチ処理実行
    batch_generate_markdown(folder, api_key)


if __name__ == "__main__":
    main()
