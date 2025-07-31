#!/usr/bin/env python3
import os
import re

# テスト結果から失敗した関数を抽出
failed_functions = {
    # 01. 数学
    "SERIESSUM": {"category": "01. 数学", "failures": 1},
    
    # 02. 統計
    "SKEW": {"category": "02. 統計", "failures": 0},
    "GEOMEAN": {"category": "02. 統計", "failures": 0},
    "HARMEAN": {"category": "02. 統計", "failures": 0},
    "TRIMMEAN": {"category": "02. 統計", "failures": 0},
    "GAMMALN": {"category": "02. 統計", "failures": 0},
    "GAUSS": {"category": "02. 統計", "failures": 0},
    "STEYX": {"category": "02. 統計", "failures": 0},
    
    # 03. 文字列
    "TEXT": {"category": "03. 文字列", "failures": 0},
    "NUMBERVALUE": {"category": "03. 文字列", "failures": 0},
    "SEARCHB": {"category": "03. 文字列", "failures": 0},
    
    # 04. 日付
    "EDATE": {"category": "04. 日付", "failures": 0},
    "EOMONTH": {"category": "04. 日付", "failures": 0},
    "DATEVALUE": {"category": "04. 日付", "failures": 0},
    "TIMEVALUE": {"category": "04. 日付", "failures": 0},
    "ISOWEEKNUM": {"category": "04. 日付", "failures": 0},
    
    # 06. 検索
    "HLOOKUP": {"category": "06. 検索", "failures": 0},
    "INDIRECT": {"category": "06. 検索", "failures": 0},
    "HYPERLINK": {"category": "06. 検索", "failures": 0},
    "FORMULATEXT": {"category": "06. 検索", "failures": 0},
    "GETPIVOTDATA": {"category": "06. 検索", "failures": 0},
    
    # 07. 財務
    "RATE": {"category": "07. 財務", "failures": 0},
    "NPER": {"category": "07. 財務", "failures": 0},
    "IRR": {"category": "07. 財務", "failures": 0},
    "XNPV": {"category": "07. 財務", "failures": 0},
    "XIRR": {"category": "07. 財務", "failures": 0},
    "IPMT": {"category": "07. 財務", "failures": 0},
    "PPMT": {"category": "07. 財務", "failures": 0},
    "MIRR": {"category": "07. 財務", "failures": 0},
    "SLN": {"category": "07. 財務", "failures": 0},
    "ACCRINT": {"category": "07. 財務", "failures": 0},
    "DURATION": {"category": "07. 財務", "failures": 0},
    "MDURATION": {"category": "07. 財務", "failures": 0},
    "PRICE": {"category": "07. 財務", "failures": 0},
    "COUPDAYS": {"category": "07. 財務", "failures": 0},
    "COUPNCD": {"category": "07. 財務", "failures": 0},
    "AMORDEGRC": {"category": "07. 財務", "failures": 0},
    "CUMIPMT": {"category": "07. 財務", "failures": 0},
    "CUMPRINC": {"category": "07. 財務", "failures": 0},
    "ODDFPRICE": {"category": "07. 財務", "failures": 0},
    "ODDLPRICE": {"category": "07. 財務", "failures": 0},
    "ODDLYIELD": {"category": "07. 財務", "failures": 0},
    "TBILLPRICE": {"category": "07. 財務", "failures": 0},
    "PRICEDISC": {"category": "07. 財務", "failures": 0},
    "RECEIVED": {"category": "07. 財務", "failures": 0},
    "INTRATE": {"category": "07. 財務", "failures": 0},
    "PRICEMAT": {"category": "07. 財務", "failures": 0},
    "YIELDMAT": {"category": "07. 財務", "failures": 0},
    
    # 08. 行列
    "MINVERSE": {"category": "08. 行列", "failures": 0},
    
    # 09. 情報
    "ISTEXT": {"category": "09. 情報", "failures": 0},
    "ISNUMBER": {"category": "09. 情報", "failures": 0},
    "TYPE": {"category": "09. 情報", "failures": 0},
    "SHEET": {"category": "09. 情報", "failures": 0},
    "SHEETS": {"category": "09. 情報", "failures": 0},
    "CELL": {"category": "09. 情報", "failures": 0},
    "INFO": {"category": "09. 情報", "failures": 0},
    
    # 10. データベース
    "DSUM": {"category": "10. データベース", "failures": 0},
    "DAVERAGE": {"category": "10. データベース", "failures": 0},
    "DCOUNT": {"category": "10. データベース", "failures": 0},
    "DCOUNTA": {"category": "10. データベース", "failures": 0},
    "DPRODUCT": {"category": "10. データベース", "failures": 0},
    "DGET": {"category": "10. データベース", "failures": 0},
    
    # 11. エンジニアリング
    "IMABS": {"category": "11. エンジニアリング", "failures": 0},
    "IMSUM": {"category": "11. エンジニアリング", "failures": 0},
    "IMDIV": {"category": "11. エンジニアリング", "failures": 0},
    "IMPOWER": {"category": "11. エンジニアリング", "failures": 0},
    "IMLOG10": {"category": "11. エンジニアリング", "failures": 0},
    "IMLOG2": {"category": "11. エンジニアリング", "failures": 0},
    "PHONETIC": {"category": "11. エンジニアリング", "failures": 0},
    "IMSQRT": {"category": "11. エンジニアリング", "failures": 0},
    "IMEXP": {"category": "11. エンジニアリング", "failures": 0},
    "IMLN": {"category": "11. エンジニアリング", "failures": 0},
    "IMSIN": {"category": "11. エンジニアリング", "failures": 0},
    "IMCOS": {"category": "11. エンジニアリング", "failures": 0},
    "IMTAN": {"category": "11. エンジニアリング", "failures": 0},
    "BESSELY": {"category": "11. エンジニアリング", "failures": 0},
    "BITAND": {"category": "11. エンジニアリング", "failures": 0},
    "BITXOR": {"category": "11. エンジニアリング", "failures": 0},
    
    # 12. 動的配列
    "TRANSPOSE": {"category": "12. 動的配列", "failures": 2},
    "SEQUENCE": {"category": "12. 動的配列", "failures": 5},
    "LAMBDA": {"category": "12. 動的配列", "failures": 2},
    "HSTACK": {"category": "12. 動的配列", "failures": 4},
    "VSTACK": {"category": "12. 動的配列", "failures": 8},
    "BYROW": {"category": "12. 動的配列", "failures": 3},
    "BYCOL": {"category": "12. 動的配列", "failures": 3},
    "MAKEARRAY": {"category": "12. 動的配列", "failures": 4},
    "MAP": {"category": "12. 動的配列", "failures": 4},
    "REDUCE": {"category": "12. 動的配列", "failures": 1},
    "SCAN": {"category": "12. 動的配列", "failures": 4},
    "TAKE": {"category": "12. 動的配列", "failures": 3},
    "DROP": {"category": "12. 動的配列", "failures": 4},
    "EXPAND": {"category": "12. 動的配列", "failures": 12},
    "TOCOL": {"category": "12. 動的配列", "failures": 4},
    "TOROW": {"category": "12. 動的配列", "failures": 4},
    "CHOOSEROWS": {"category": "12. 動的配列", "failures": 4},
    "CHOOSECOLS": {"category": "12. 動的配列", "failures": 6},
    "WRAPROWS": {"category": "12. 動的配列", "failures": 6},
    "WRAPCOLS": {"category": "12. 動的配列", "failures": 6},
    
    # 13. キューブ
    "CUBEVALUE": {"category": "13. キューブ", "failures": 0},
    "CUBESETCOUNT": {"category": "13. キューブ", "failures": 0},
    
    # 14. Web・インポート
    "REGEXEXTRACT": {"category": "14. Web・インポート", "failures": 0},
    "REGEXMATCH": {"category": "14. Web・インポート", "failures": 0},
    "REGEXREPLACE": {"category": "14. Web・インポート", "failures": 0},
    "SORTN": {"category": "14. Web・インポート", "failures": 0},
    "WEBSERVICE": {"category": "14. Web・インポート", "failures": 0},
    "SPARKLINE": {"category": "14. Web・インポート", "failures": 0},
    "IMPORTDATA": {"category": "14. Web・インポート", "failures": 0},
    "IMPORTFEED": {"category": "14. Web・インポート", "failures": 0},
    "IMPORTHTML": {"category": "14. Web・インポート", "failures": 0},
    "IMPORTXML": {"category": "14. Web・インポート", "failures": 0},
    "IMPORTRANGE": {"category": "14. Web・インポート", "failures": 0},
    "IMAGE": {"category": "14. Web・インポート", "failures": 0},
    "GOOGLEFINANCE": {"category": "14. Web・インポート", "failures": 0},
    "GOOGLETRANSLATE": {"category": "14. Web・インポート", "failures": 0},
    "DETECTLANGUAGE": {"category": "14. Web・インポート", "failures": 0},
    "TO_DATE": {"category": "14. Web・インポート", "failures": 0},
    "TO_PERCENT": {"category": "14. Web・インポート", "failures": 0},
    "TO_TEXT": {"category": "14. Web・インポート", "failures": 0},
    
    # 15. その他
    "ISOMITTED": {"category": "15. その他", "failures": 0},
    "STOCKHISTORY": {"category": "15. その他", "failures": 0},
}

# 実装された関数をチェック
implemented_functions = set()

# logic/index.tsから実装関数を抽出
logic_index_path = 'src/utils/spreadsheet/logic/index.ts'
if os.path.exists(logic_index_path):
    with open(logic_index_path, 'r', encoding='utf-8') as f:
        content = f.read()
        # ALL_FUNCTIONS配列内の関数名を抽出
        all_functions_match = re.search(r'export const ALL_FUNCTIONS = \[(.*?)\] as CustomFormula\[\];', content, re.DOTALL)
        if all_functions_match:
            all_functions_content = all_functions_match.group(1)
            # 関数名を抽出（大文字で始まる識別子）
            matches = re.findall(r'\b([A-Z][A-Z0-9_]+)\b', all_functions_content)
            # 予約語やObject.valuesなどを除外
            for match in matches:
                if match not in ['Object', 'CustomFormula', 'FUNCTION_CATEGORIES']:
                    implemented_functions.add(match)

# 失敗した関数を実装状況で分類
implemented_failed = []
unimplemented_failed = []

for func_name, details in failed_functions.items():
    if func_name in implemented_functions:
        implemented_failed.append({
            "name": func_name,
            "category": details["category"],
            "failures": details["failures"]
        })
    else:
        unimplemented_failed.append({
            "name": func_name,
            "category": details["category"],
            "failures": details["failures"]
        })

# 結果を出力
print("=== 失敗している関数の分析結果 ===\n")
print(f"失敗した関数の総数: {len(failed_functions)}")
print(f"実装済みで失敗: {len(implemented_failed)}")
print(f"未実装で失敗: {len(unimplemented_failed)}")

# カテゴリ別にグループ化
def group_by_category(functions):
    categories = {}
    for func in functions:
        cat = func["category"]
        if cat not in categories:
            categories[cat] = []
        categories[cat].append(func)
    return categories

print("\n=== 実装済みだが失敗している関数（優先度高） ===")
implemented_categories = group_by_category(implemented_failed)
for category in sorted(implemented_categories.keys()):
    print(f"\n{category}:")
    for func in implemented_categories[category]:
        if func["failures"] > 0:
            print(f"  - {func['name']}: {func['failures']}個の失敗")
        else:
            print(f"  - {func['name']}: 失敗（詳細不明）")

print("\n=== 未実装の関数 ===")
unimplemented_categories = group_by_category(unimplemented_failed)
for category in sorted(unimplemented_categories.keys()):
    print(f"\n{category}:")
    for func in unimplemented_categories[category]:
        print(f"  - {func['name']}")

# 特に失敗数が多い関数をリスト
print("\n=== 失敗数が多い関数（上位10個） ===")
sorted_by_failures = sorted(
    [(f["name"], f["failures"]) for f in implemented_failed if f["failures"] > 0],
    key=lambda x: x[1],
    reverse=True
)[:10]

for func_name, failure_count in sorted_by_failures:
    print(f"  - {func_name}: {failure_count}個の失敗")