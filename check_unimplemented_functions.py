#!/usr/bin/env python3
import os
import re
import json

# 実装された関数をチェック
implemented_functions = set()

# logic/index.tsから実装関数を抽出
logic_index_path = 'src/utils/spreadsheet/logic/index.ts'
if os.path.exists(logic_index_path):
    with open(logic_index_path, 'r', encoding='utf-8') as f:
        content = f.read()
        # ALL_FUNCTIONS配列内の関数名を抽出
        # まずALL_FUNCTIONS配列の内容を取得
        all_functions_match = re.search(r'export const ALL_FUNCTIONS = \[(.*?)\] as CustomFormula\[\];', content, re.DOTALL)
        if all_functions_match:
            all_functions_content = all_functions_match.group(1)
            # 関数名を抽出（大文字で始まる識別子）
            matches = re.findall(r'\b([A-Z][A-Z0-9_]+)\b', all_functions_content)
            # 予約語やObject.valuesなどを除外
            for match in matches:
                if match not in ['Object', 'CustomFormula', 'FUNCTION_CATEGORIES']:
                    implemented_functions.add(match)

# テストデータから全関数を抽出
test_dir = 'src/data/individualTests'
all_functions = set()
function_details = {}

for filename in os.listdir(test_dir):
    if filename.endswith('.ts'):
        filepath = os.path.join(test_dir, filename)
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
            # name: 'FUNCTION_NAME' のパターンを探す
            matches = re.findall(r"name:\s*['\"]([^'\"]+)['\"]", content)
            for func in matches:
                all_functions.add(func)
                # カテゴリも取得
                category_match = re.search(rf"name:\s*['\"]({func})['\"][^{{]+category:\s*['\"]([^'\"]+)['\"]", content, re.DOTALL)
                if category_match:
                    function_details[func] = category_match.group(2)

# 未実装関数を特定
unimplemented = all_functions - implemented_functions

print(f"総関数数: {len(all_functions)}")
print(f"実装済み関数数: {len(implemented_functions)}")
print(f"未実装関数数: {len(unimplemented)}")

if unimplemented:
    print("\n未実装関数一覧（カテゴリ別）:")
    category_groups = {}
    for func in sorted(unimplemented):
        category = function_details.get(func, "Unknown")
        if category not in category_groups:
            category_groups[category] = []
        category_groups[category].append(func)
    
    for category in sorted(category_groups.keys()):
        print(f"\n{category}:")
        for func in category_groups[category]:
            print(f"  - {func}")

# 実装されているが動的な値を返す関数
dynamic_functions = ['RAND', 'RANDBETWEEN', 'TODAY', 'NOW', 'RANDARRAY']
dynamic_implemented = [f for f in dynamic_functions if f in implemented_functions]
if dynamic_implemented:
    print(f"\n動的関数（expectedValuesを削除すべき）: {', '.join(dynamic_implemented)}")