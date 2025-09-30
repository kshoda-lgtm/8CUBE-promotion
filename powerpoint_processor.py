#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PowerPoint解析・JSON変換スクリプト（無料版）
python-pptxを使用してPowerPointファイルからテキストを抽出し、
Google Apps Scriptで処理しやすいJSON形式に変換する
"""

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional

try:
    from pptx import Presentation
    from pptx.shapes.base import BaseShape
    from pptx.shapes.table import Table
    from pptx.text.text import TextFrame
except ImportError:
    print("❌ python-pptx がインストールされていません")
    print("pip install python-pptx でインストールしてください")
    exit(1)


class PowerPointProcessor:
    """PowerPoint解析・JSON変換クラス"""

    def __init__(self):
        """初期化"""
        self.patterns = {
            'price': r'(?:単価|価格|¥|円).*?(\d{1,3}(?:,\d{3})*(?:\.\d+)?)',
            'quantity': r'(?:数量|個数|枚数|ロット).*?(\d{1,3}(?:,\d{3})*)',
            'deadline': r'(?:納期|期間|日程).*?(\d+(?:日|週間|ヶ月))',
            'company': r'(?:株式会社|有限会社|\(株\)|\(有\))([^\s]+)',
            'date': r'(\d{4}年\d{1,2}月|\d{4}/\d{1,2}|\d{1,2}月)',
            'event_type': r'(キャンペーン|イベント|展示会|セミナー|プロモーション)',
        }

    def extract_text_from_slide(self, slide) -> List[str]:
        """スライドからテキストを抽出"""
        texts = []

        for shape in slide.shapes:
            text = self._extract_text_from_shape(shape)
            if text.strip():
                texts.append(text.strip())

        return texts

    def _extract_text_from_shape(self, shape: BaseShape) -> str:
        """図形からテキストを抽出"""
        text = ""

        # テキストフレーム
        if hasattr(shape, 'text_frame') and shape.text_frame:
            text += shape.text_frame.text + "\n"

        # テーブル
        if hasattr(shape, 'table') and shape.table:
            table = shape.table
            for row in table.rows:
                row_text = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text:
                        row_text.append(cell_text)
                if row_text:
                    text += " | ".join(row_text) + "\n"

        # グループ化された図形
        if hasattr(shape, 'shapes'):
            for sub_shape in shape.shapes:
                text += self._extract_text_from_shape(sub_shape)

        return text

    def analyze_text(self, text: str) -> Dict[str, Any]:
        """テキストから情報を抽出"""
        info = {
            'prices': [],
            'quantities': [],
            'deadlines': [],
            'companies': [],
            'dates': [],
            'event_types': [],
            'keywords': []
        }

        # パターンマッチング
        for key, pattern in self.patterns.items():
            matches = re.findall(pattern, text, re.IGNORECASE)
            if key == 'price':
                info['prices'] = [self._clean_number(m) for m in matches]
            elif key == 'quantity':
                info['quantities'] = [self._clean_number(m) for m in matches]
            elif key == 'deadline':
                info['deadlines'] = matches
            elif key == 'company':
                info['companies'] = matches
            elif key == 'date':
                info['dates'] = matches
            elif key == 'event_type':
                info['event_types'] = matches

        # キーワード抽出（簡易版）
        keywords = self._extract_keywords(text)
        info['keywords'] = keywords

        return info

    def _clean_number(self, number_str: str) -> Optional[int]:
        """数値文字列をクリーンアップ"""
        try:
            # カンマを除去して数値に変換
            cleaned = re.sub(r'[,\s]', '', number_str)
            return int(cleaned)
        except (ValueError, AttributeError):
            return None

    def _extract_keywords(self, text: str) -> List[str]:
        """キーワード抽出（簡易版）"""
        # 重要そうなキーワードを抽出
        important_words = [
            'ノベルティ', '景品', 'グッズ', 'プレゼント', 'キャンペーン',
            '展示会', 'イベント', 'セミナー', 'プロモーション',
            'エコ', '環境', 'SDGs', 'オリジナル', 'カスタム'
        ]

        keywords = []
        for word in important_words:
            if word in text:
                keywords.append(word)

        return keywords

    def process_powerpoint(self, file_path: str) -> Dict[str, Any]:
        """PowerPointファイルを処理してJSON化"""
        try:
            presentation = Presentation(file_path)

            result = {
                'file_info': {
                    'file_name': Path(file_path).name,
                    'processed_at': datetime.now().isoformat(),
                    'slide_count': len(presentation.slides)
                },
                'slides': [],
                'summary': {
                    'all_prices': [],
                    'all_quantities': [],
                    'all_companies': [],
                    'all_keywords': []
                }
            }

            # 各スライドを処理
            for i, slide in enumerate(presentation.slides, 1):
                slide_texts = self.extract_text_from_slide(slide)
                combined_text = "\n".join(slide_texts)

                analyzed_info = self.analyze_text(combined_text)

                slide_data = {
                    'slide_number': i,
                    'raw_texts': slide_texts,
                    'analyzed_info': analyzed_info,
                    'text_length': len(combined_text)
                }

                result['slides'].append(slide_data)

                # サマリー情報を蓄積
                result['summary']['all_prices'].extend(analyzed_info['prices'])
                result['summary']['all_quantities'].extend(analyzed_info['quantities'])
                result['summary']['all_companies'].extend(analyzed_info['companies'])
                result['summary']['all_keywords'].extend(analyzed_info['keywords'])

            # 重複を除去
            result['summary']['all_companies'] = list(set(result['summary']['all_companies']))
            result['summary']['all_keywords'] = list(set(result['summary']['all_keywords']))

            return result

        except Exception as e:
            return {
                'error': str(e),
                'file_name': Path(file_path).name,
                'processed_at': datetime.now().isoformat()
            }


def main():
    """メイン処理"""
    processor = PowerPointProcessor()

    # 使用例
    print("🚀 PowerPoint解析システム（無料版）")
    print("=" * 50)

    # ファイルパスを指定（例）
    file_path = input("PowerPointファイルのパスを入力してください: ")

    if not Path(file_path).exists():
        print(f"❌ ファイルが見つかりません: {file_path}")
        return

    print(f"📁 処理中: {file_path}")
    result = processor.process_powerpoint(file_path)

    if 'error' in result:
        print(f"❌ エラー: {result['error']}")
        return

    # JSON出力
    output_path = Path(file_path).with_suffix('.json')
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"✅ 処理完了")
    print(f"📄 出力ファイル: {output_path}")
    print(f"📊 スライド数: {result['file_info']['slide_count']}")
    print(f"💰 発見した価格: {len(result['summary']['all_prices'])}件")
    print(f"🏢 発見した会社: {len(result['summary']['all_companies'])}件")


if __name__ == "__main__":
    main()