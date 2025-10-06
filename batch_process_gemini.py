#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PowerPointファイルのバッチ処理スクリプト（Gemini API版 v4.0）
指定フォルダ内のすべての.pptxファイルをGemini APIで処理してJSONに変換
"""

import os
import sys
from pathlib import Path
from powerpoint_processor_gemini import GeminiPowerPointProcessor
import json
import time


def batch_process_folder(folder_path: str, api_key: str):
    """フォルダ内の全PowerPointファイルをGemini APIで処理"""

    # プロセッサー初期化
    try:
        processor = GeminiPowerPointProcessor(api_key=api_key)
    except Exception as e:
        print(f"ERROR: Failed to initialize Gemini API: {e}")
        return

    # .pptxファイルを再帰的に検索
    pptx_files = list(Path(folder_path).rglob("*.pptx"))

    if not pptx_files:
        print(f"No .pptx files found in: {folder_path}")
        return

    print(f"\nFound {len(pptx_files)} PowerPoint files")
    print("=" * 60)

    success_count = 0
    error_count = 0
    results = []

    for i, pptx_file in enumerate(pptx_files, 1):
        print(f"\n[{i}/{len(pptx_files)}] Processing: {pptx_file.name}")
        print("-" * 60)

        try:
            # Gemini APIで処理
            result = processor.process_powerpoint(str(pptx_file))

            if 'error' in result:
                error_msg = result['error']
                print(f"  ERROR: {error_msg}")

                # 無料枠超過の場合は処理を停止
                if error_msg == 'FREE_TIER_LIMIT_EXCEEDED':
                    print(f"\n⚠️  バッチ処理を停止します（無料枠超過）")
                    print(f"   処理済み: {success_count}ファイル")
                    print(f"   未処理: {len(pptx_files) - i}ファイル")
                    break

                error_count += 1
                results.append({
                    'file': pptx_file.name,
                    'status': 'error',
                    'error': error_msg
                })
                continue

            # JSON出力
            output_path = pptx_file.with_suffix('.json')
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)

            # 結果表示
            analysis = result['gemini_analysis']
            print(f"  SUCCESS: {output_path.name}")
            print(f"    - Slides: {result['file_info']['slide_count']}")
            print(f"    - Client: {analysis.get('client_name', 'N/A')}")
            print(f"    - Event Type: {analysis.get('event_type', 'N/A')}")
            print(f"    - Event Date: {analysis.get('event_date', 'N/A')}")
            print(f"    - Companies: {len(analysis.get('partner_companies', []))}")
            print(f"    - Confidence: {analysis.get('confidence_score', 0)}%")

            success_count += 1
            results.append({
                'file': pptx_file.name,
                'output': output_path.name,
                'slides': result['file_info']['slide_count'],
                'confidence': analysis.get('confidence_score', 0),
                'status': 'success'
            })

            # APIレート制限を考慮して少し待機
            time.sleep(1)

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
    print("BATCH PROCESSING SUMMARY (Gemini API v4.0)")
    print("=" * 60)
    print(f"Total files: {len(pptx_files)}")
    print(f"Success: {success_count}")
    print(f"Errors: {error_count}")

    # 平均信頼度スコア
    confidence_scores = [r.get('confidence', 0) for r in results if r.get('status') == 'success']
    if confidence_scores:
        avg_confidence = sum(confidence_scores) / len(confidence_scores)
        print(f"Average Confidence: {avg_confidence:.1f}%")

    # サマリーJSONを保存
    summary_path = Path(folder_path) / "_batch_summary_gemini.json"
    with open(summary_path, 'w', encoding='utf-8') as f:
        json.dump({
            'processing_method': 'gemini_api_v4.0',
            'total': len(pptx_files),
            'success': success_count,
            'errors': error_count,
            'average_confidence': avg_confidence if confidence_scores else 0,
            'results': results
        }, f, ensure_ascii=False, indent=2)

    print(f"\nSummary saved to: {summary_path}")


def main():
    """メイン処理"""
    print("=" * 60)
    print("PowerPoint Batch Processing (Gemini API v4.0)")
    print("=" * 60)

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
    batch_process_folder(folder, api_key)


if __name__ == "__main__":
    main()