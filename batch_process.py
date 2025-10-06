#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PowerPointファイルのバッチ処理スクリプト
指定フォルダ内のすべての.pptxファイルを処理してJSONに変換
"""

import os
import sys
from pathlib import Path
from powerpoint_processor import PowerPointProcessor
import json

def batch_process_folder(folder_path: str):
    """フォルダ内の全PowerPointファイルを処理"""
    processor = PowerPointProcessor()

    # .pptxファイルを再帰的に検索
    pptx_files = list(Path(folder_path).rglob("*.pptx"))

    if not pptx_files:
        print(f"No .pptx files found in: {folder_path}")
        return

    print(f"Found {len(pptx_files)} PowerPoint files")
    print("=" * 60)

    success_count = 0
    error_count = 0
    results = []

    for i, pptx_file in enumerate(pptx_files, 1):
        print(f"\n[{i}/{len(pptx_files)}] Processing: {pptx_file.name}")

        try:
            # 絶対パスを使用
            abs_path = str(pptx_file.absolute())
            result = processor.process_powerpoint(abs_path)

            if 'error' in result:
                print(f"  ERROR: {result['error']}")
                error_count += 1
                continue

            # JSON出力
            output_path = pptx_file.with_suffix('.json')
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)

            print(f"  SUCCESS: {output_path.name}")
            print(f"    - Slides: {result['file_info']['slide_count']}")
            print(f"    - Prices: {len(result['summary']['all_prices'])}")
            print(f"    - Companies: {len(result['summary']['all_companies'])}")
            print(f"    - Keywords: {len(result['summary']['all_keywords'])}")

            success_count += 1
            results.append({
                'file': pptx_file.name,
                'output': output_path.name,
                'slides': result['file_info']['slide_count'],
                'status': 'success'
            })

        except Exception as e:
            print(f"  ERROR: {str(e)}")
            error_count += 1
            results.append({
                'file': pptx_file.name,
                'status': 'error',
                'error': str(e)
            })

    # サマリー出力
    print("\n" + "=" * 60)
    print("BATCH PROCESSING SUMMARY")
    print("=" * 60)
    print(f"Total files: {len(pptx_files)}")
    print(f"Success: {success_count}")
    print(f"Errors: {error_count}")

    # サマリーJSONを保存
    summary_path = Path(folder_path) / "_batch_summary.json"
    with open(summary_path, 'w', encoding='utf-8') as f:
        json.dump({
            'total': len(pptx_files),
            'success': success_count,
            'errors': error_count,
            'results': results
        }, f, ensure_ascii=False, indent=2)

    print(f"\nSummary saved to: {summary_path}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        folder = sys.argv[1]
    else:
        folder = input("Enter folder path to process: ")

    if not Path(folder).exists():
        print(f"ERROR: Folder not found: {folder}")
        sys.exit(1)

    batch_process_folder(folder)