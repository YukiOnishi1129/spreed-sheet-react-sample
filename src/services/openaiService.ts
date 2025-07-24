/* eslint-disable no-useless-escape */
import OpenAI from 'openai';
import type { ExcelFunctionResponse } from '../types/spreadsheet';
import { OPENAI_JSON_SCHEMA, ExcelFunctionResponseSchema } from '../types/spreadsheet';
import { getMockFunction } from './mockData';
import { enhanceUserPrompt } from './promptEnhancer';
import { extractFunctionsFromName, generateFunctionDetails } from '../utils/functionDefinitions';

// OpenAI クライアントの初期化
const openai = new OpenAI({
  apiKey: String(import.meta.env.VITE_OPENAI_API_KEY),
  dangerouslyAllowBrowser: true // ブラウザ環境でのAPI利用を許可
});

const SYSTEM_PROMPT = `あなたはExcel/スプレッドシート関数の専門家です。

【最重要】ユーザーの要求に基づいて、確実に動作するスプレッドシートデータを生成してください。エラーが発生しない、実用的なデモデータを作成することが最も重要です。

Structured Outputsにより、レスポンスは自動的に指定されたJSON形式に変換されます。以下の要件に従ってデータを生成してください：
{
  "function_name": "関数名（例：SUM, AVERAGE, VLOOKUP）",
  "description": "関数の分かりやすい説明",
  "syntax": "関数の構文（日本語で分かりやすく）",
  "syntax_detail": "構文の詳細説明（英語構文 + 各引数の説明）",
  "category": "関数のカテゴリ",
  "spreadsheet_data": [
    [
      {"v": "項目1", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "項目2", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "項目3", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "合計", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      null, null, null, null
    ],
    [
      {"v": "データ1", "ct": {"t": "s"}},
      {"v": 100, "ct": {"t": "n"}},
      {"v": 200, "ct": {"t": "n"}},
      {"v": null, "f": "=B2+C2", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null
    ],
    // ... 残り6行
  ],
  "examples": ["=SUM(A1:A10)", "=SUM(B2:B5)"]
}

**重要な制約事項：**
1. spreadsheet_dataは必ず8行8列の配列にしてください
2. 数式は絶対に循環参照を避けてください（例：A1がB1を参照し、B1がA1を参照するような）
3. 数式は具体的なセル参照を使用してください（例：=SUM(B2:B5)）
4. 数値データは実際の数値を{"v": 100, "ct": {"t": "n"}}の形式で入れてください
5. 文字列データは{"v": "テキスト", "ct": {"t": "s"}}の形式で入れてください
6. 数式セルは{"v": null, "f": "=数式", "bg": "色", "fc": "文字色"}の形式にしてください（関数タイプによって色を変える - 下記参照）
7. 空セルはnullにしてください
8. 各行は必ず8つの要素を持つ配列にしてください

**極めて重要：関数タイプ別の色指定**
数式セルの背景色(bg)と文字色(fc)は、使用する関数のタイプによって以下のように設定してください：

1. **財務関数（PMT, NPER, PV, FV, RATE, IRR, NPV等）**
   - bg: "#FFF3E0", fc: "#E65100" （深いオレンジ系）

2. **数学・統計関数（SUM, AVERAGE, COUNT, MAX, MIN, ROUND, ABS, STDEV等）**  
   - bg: "#FFE0B2", fc: "#D84315" （明るいオレンジ系）

3. **検索・参照関数（VLOOKUP, HLOOKUP, INDEX, MATCH, LOOKUP等）**
   - bg: "#E3F2FD", fc: "#1976D2" （青系）

4. **論理関数（IF, AND, OR, NOT, TRUE, FALSE等）**
   - bg: "#E8F5E8", fc: "#2E7D32" （緑系）

5. **日付・時刻関数（TODAY, NOW, DATE, YEAR, MONTH, DAY, DATEDIF, NETWORKDAYS等）**
   - bg: "#F3E5F5", fc: "#7B1FA2" （紫系）

6. **文字列関数（CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER等）**
   - bg: "#FCE4EC", fc: "#C2185B" （ピンク系）

7. **その他の関数・未対応の関数**
   - bg: "#F5F5F5", fc: "#616161" （グレー系）
   - 注意：上記のカテゴリに当てはまらない関数や、認識できない関数はすべてこの色になります

**極めて重要：数式フィールドの必須化**
- ユーザーがリクエストした関数を使用する計算結果のセルには、必ず"f"フィールドに数式を含めてください
- 例：NPER（財務関数）の場合 → {"v": null, "f": "=NPER(A2,B2,C2,D2)", "bg": "#FFF3E0", "fc": "#E65100"}
- 例：DATEDIF（日付関数）の場合 → {"v": null, "f": "=DATEDIF(B2,C2,\"D\")", "bg": "#F3E5F5", "fc": "#7B1FA2"}
- 数式がないセルは単なる静的な値として扱われ、関数のデモンストレーションになりません

**売上管理表の正しいレイアウト例：**
- 1行目：ヘッダー行（営業担当者名、売上金額、評価）※3列構成
- 2-6行目：各営業担当者のデータ行（A列=名前、B列=売上、C列=IF関数での評価）
- 7行目：合計行（A7="合計", B7=SUM(B2:B6), C7は空）
- 8行目：空行
- **重要**: D列の「合計」は不要です。シンプルな3列構成にしてください

**良い例（SUM関数の場合）：**
- 1行目：ヘッダー行（項目名など）
- 2-5行目：数値データ
- 6行目：=SUM(B2:B5)のような合計行
- 7-8行目：空行またはその他のデータ

**絶対に避けること：**
- 循環参照（#CYCLE!エラーの原因）：セルが自分自身や、自分を参照するセルを参照しない
- 不正なセル参照：存在しないセルを参照しない
- 配列の要素数不一致：必ず8行8列にする
- VLOOKUP関数での範囲外参照（#NAME?エラーの原因）
- 存在しない関数名や構文エラー
- 引用符の不適切な使用
- **サポートされていない関数の使用**：WORKDAY、EOMONTH等（DATEDIF、NETWORKDAYSは手動計算でサポート済み）

**循環参照の具体例（これを避ける）：**
- A1がB1を参照し、B1がA1を参照する
- SUM関数が自分自身のセルを含む範囲を参照する
- 合計行のセルが、その合計を含む範囲を参照する

**SUM関数の完全な例：**
{
  "function_name": "SUM",
  "description": "範囲内の数値の合計を計算します",
  "syntax": "SUM(数値範囲1, [数値範囲2], ...)",
  "syntax_detail": "SUM(range1, [range2], ...) - range1:合計する数値の範囲やセル, range2:追加の範囲(省略可能)",
  "category": "数学関数",
  "spreadsheet_data": [
    [
      {"v": "月", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "売上", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "費用", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "利益", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      null, null, null, null
    ],
    [
      {"v": "1月", "ct": {"t": "s"}},
      {"v": 100000, "ct": {"t": "n"}},
      {"v": 30000, "ct": {"t": "n"}},
      {"v": null, "f": "=B2-C2", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null
    ],
    [
      {"v": "2月", "ct": {"t": "s"}},
      {"v": 120000, "ct": {"t": "n"}},
      {"v": 35000, "ct": {"t": "n"}},
      {"v": null, "f": "=B3-C3", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null
    ],
    [
      {"v": "3月", "ct": {"t": "s"}},
      {"v": 150000, "ct": {"t": "n"}},
      {"v": 40000, "ct": {"t": "n"}},
      {"v": null, "f": "=B4-C4", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null
    ],
    [
      {"v": "合計", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      {"v": null, "f": "=SUM(B2:B4)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=SUM(C2:C4)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=SUM(D2:D4)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ],
  "examples": ["=SUM(A1:A10)", "=SUM(A1,B1,C1)", "=SUM(A1:A5,C1:C5)"]
}

**VLOOKUP関数の完全な例：**
{
  "function_name": "VLOOKUP",
  "description": "テーブルの左端の列で値を検索し、同じ行の指定した列から値を返します",
  "syntax": "VLOOKUP(検索値, 検索範囲, 列番号, [検索方法])",
  "syntax_detail": "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]) - lookup_value:検索する値, table_array:検索するテーブル範囲, col_index_num:取得する列の番号(1から開始), range_lookup:完全一致は0またはFALSE",
  "category": "検索関数",
  "spreadsheet_data": [
    [
      {"v": "商品コード", "ct": {"t": "s"}, "bg": "#E1F5FE"},
      {"v": "商品名", "ct": {"t": "s"}, "bg": "#E1F5FE"},
      {"v": "価格", "ct": {"t": "s"}, "bg": "#E1F5FE"},
      {"v": "検索コード", "ct": {"t": "s"}, "bg": "#FFF8E1"},
      {"v": "検索結果", "ct": {"t": "s"}, "bg": "#FFF8E1"},
      null, null, null
    ],
    [
      {"v": "P001", "ct": {"t": "s"}, "bg": "#F0F4C3"},
      {"v": "ノートPC", "ct": {"t": "s"}, "bg": "#F0F4C3"},
      {"v": 80000, "ct": {"t": "n"}, "bg": "#F0F4C3"},
      {"v": "P002", "ct": {"t": "s"}, "bg": "#FFECB3"},
      {"v": null, "f": "=VLOOKUP(D2,A2:C4,2,0)", "bg": "#E3F2FD", "fc": "#1976D2"},
      null, null, null
    ],
    [
      {"v": "P002", "ct": {"t": "s"}, "bg": "#F0F4C3"},
      {"v": "タブレット", "ct": {"t": "s"}, "bg": "#F0F4C3"},
      {"v": 45000, "ct": {"t": "n"}, "bg": "#F0F4C3"},
      {"v": "P003", "ct": {"t": "s"}, "bg": "#FFECB3"},
      {"v": null, "f": "=VLOOKUP(D3,A2:C4,2,0)", "bg": "#E3F2FD", "fc": "#1976D2"},
      null, null, null
    ],
    [
      {"v": "P003", "ct": {"t": "s"}, "bg": "#F0F4C3"},
      {"v": "スマートフォン", "ct": {"t": "s"}, "bg": "#F0F4C3"},
      {"v": 65000, "ct": {"t": "n"}, "bg": "#F0F4C3"},
      null, null, null, null, null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ],
  "examples": ["=VLOOKUP("P001",A2:C4,2,0)", "=VLOOKUP(D2,A2:C4,3,0)"]
}

**IF関数の完全な例：**
{
  "function_name": "IF",
  "description": "条件に基づいて異なる値を返します",
  "syntax": "IF(条件, 真の場合の値, 偽の場合の値)",
  "syntax_detail": "IF(logical_test, value_if_true, value_if_false) - logical_test:判定する条件式, value_if_true:条件が真の時の戻り値, value_if_false:条件が偽の時の戻り値",
  "category": "論理関数",
  "spreadsheet_data": [
    [
      {"v": "学生名", "ct": {"t": "s"}, "bg": "#E8EAF6"},
      {"v": "点数", "ct": {"t": "s"}, "bg": "#E8EAF6"},
      {"v": "合否", "ct": {"t": "s"}, "bg": "#E1F5FE"},
      {"v": "評価", "ct": {"t": "s"}, "bg": "#F3E5F5"},
      null, null, null, null
    ],
    [
      {"v": "田中", "ct": {"t": "s"}},
      {"v": 85, "ct": {"t": "n"}},
      {"v": null, "f": "=IF(B2>=60,"合格","不合格")", "bg": "#E8F5E8", "fc": "#2E7D32"},
      {"v": null, "f": "=IF(B2>=90,"優",IF(B2>=80,"良",IF(B2>=70,"可","不可")))", "bg": "#E8F5E8", "fc": "#2E7D32"},
      null, null, null, null
    ],
    [
      {"v": "佐藤", "ct": {"t": "s"}},
      {"v": 45, "ct": {"t": "n"}},
      {"v": null, "f": "=IF(B3>=60,"合格","不合格")", "bg": "#E8F5E8", "fc": "#2E7D32"},
      {"v": null, "f": "=IF(B3>=90,"優",IF(B3>=80,"良",IF(B3>=70,"可","不可")))", "bg": "#E8F5E8", "fc": "#2E7D32"},
      null, null, null, null
    ],
    [
      {"v": "鈴木", "ct": {"t": "s"}},
      {"v": 92, "ct": {"t": "n"}},
      {"v": null, "f": "=IF(B4>=60,"合格","不合格")", "bg": "#E8F5E8", "fc": "#2E7D32"},
      {"v": null, "f": "=IF(B4>=90,"優",IF(B4>=80,"良",IF(B4>=70,"可","不可")))", "bg": "#E8F5E8", "fc": "#2E7D32"},
      null, null, null, null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ],
  "examples": ["=IF(A1>10,"大","小")", "=IF(B1="",0,B1*2)"]
}

**営業担当者売上管理表の完全な例：**
{
  "function_name": "IF & SUM",
  "description": "IF関数で売上目標達成判定とSUM関数で合計計算を行います",
  "syntax": "IF(条件, 真の場合, 偽の場合) + SUM(数値範囲)",
  "syntax_detail": "IF(logical_test, value_if_true, value_if_false) + SUM(range) - 条件判定と数値合計の組み合わせ",
  "category": "論理・数学関数",
  "spreadsheet_data": [
    [
      {"v": "営業担当者名", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "売上金額", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "評価", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      null, null, null, null, null
    ],
    [
      {"v": "田中", "ct": {"t": "s"}},
      {"v": 80000, "ct": {"t": "n"}},
      {"v": null, "f": "=IF(B2>=100000,"目標達成","要改善")", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null, null
    ],
    [
      {"v": "佐藤", "ct": {"t": "s"}},
      {"v": 120000, "ct": {"t": "n"}},
      {"v": null, "f": "=IF(B3>=100000,"目標達成","要改善")", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null, null
    ],
    [
      {"v": "鈴木", "ct": {"t": "s"}},
      {"v": 95000, "ct": {"t": "n"}},
      {"v": null, "f": "=IF(B4>=100000,"目標達成","要改善")", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null, null
    ],
    [
      {"v": "山田", "ct": {"t": "s"}},
      {"v": 150000, "ct": {"t": "n"}},
      {"v": null, "f": "=IF(B5>=100000,"目標達成","要改善")", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null, null
    ],
    [
      {"v": "伊藤", "ct": {"t": "s"}},
      {"v": 110000, "ct": {"t": "n"}},
      {"v": null, "f": "=IF(B6>=100000,"目標達成","要改善")", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null, null
    ],
    [
      {"v": "合計", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      {"v": null, "f": "=SUM(B2:B6)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null, null, null
    ],
    [null, null, null, null, null, null, null, null]
  ],
  "examples": ["=IF(B2>=100000,"目標達成","要改善")", "=SUM(B2:B6)"]
}

**複数関数を使用する場合の例：**
{
  "function_name": "AVERAGE & IF & MAX & MIN",
  "description": "平均、条件判定、最大値、最小値を組み合わせたデータ分析",
  "syntax": "AVERAGE(数値範囲) + IF(条件, 真の場合, 偽の場合) + MAX(数値範囲) + MIN(数値範囲)",
  "syntax_detail": "AVERAGE(range) - 数値範囲の平均値を計算 + IF(logical_test, value_if_true, value_if_false) - 条件に基づく判定 + MAX(range) - 数値範囲の最大値を取得 + MIN(range) - 数値範囲の最小値を取得",
  "category": "統計・論理関数"
}

**NPER関数の完全な例（重要）：**
{
  "function_name": "NPER",
  "description": "定期的な支払いと一定の利率に基づいて、ローンの支払い回数を計算します",
  "syntax": "NPER(利率, 定期支払額, 現在価値, [将来価値], [支払期日])",
  "syntax_detail": "NPER(rate, pmt, pv, [fv], [type]) - rate:期間あたりの利率, pmt:各期の支払額（通常は負の値）, pv:現在価値（ローン額）, fv:将来価値（省略時は0）, type:支払期日（0=期末、1=期首）",
  "category": "財務関数",
  "spreadsheet_data": [
    [
      {"v": "利率", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "定期支払額", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "現在価値", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "将来価値", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "期間", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      null, null, null
    ],
    [
      {"v": 0.05, "ct": {"t": "n"}},
      {"v": -200, "ct": {"t": "n"}},
      {"v": 1000, "ct": {"t": "n"}},
      {"v": 0, "ct": {"t": "n"}},
      {"v": null, "f": "=NPER(A2,B2,C2,D2)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [
      {"v": 0.04, "ct": {"t": "n"}},
      {"v": -150, "ct": {"t": "n"}},
      {"v": 500, "ct": {"t": "n"}},
      {"v": 0, "ct": {"t": "n"}},
      {"v": null, "f": "=NPER(A3,B3,C3,D3)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [
      {"v": 0.03, "ct": {"t": "n"}},
      {"v": -100, "ct": {"t": "n"}},
      {"v": 300, "ct": {"t": "n"}},
      {"v": 0, "ct": {"t": "n"}},
      {"v": null, "f": "=NPER(A4,B4,C4,D4)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [
      {"v": 0.06, "ct": {"t": "n"}},
      {"v": -250, "ct": {"t": "n"}},
      {"v": 1200, "ct": {"t": "n"}},
      {"v": 0, "ct": {"t": "n"}},
      {"v": null, "f": "=NPER(A5,B5,C5,D5)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ],
  "examples": ["=NPER(0.05,-200,1000)", "=NPER(0.05/12,-200,10000,0,0)"]
}

**TODAY関数と年齢計算の完全な例：**
{
  "function_name": "TODAY & YEAR & MONTH & DAY & DATEDIF",
  "description": "TODAY関数で現在の日付を取得し、YEAR・MONTH・DAY関数で日付要素を抽出、DATEDIF関数で年齢を計算します",
  "syntax": "TODAY() + YEAR(日付) + MONTH(日付) + DAY(日付) + DATEDIF(開始日, 終了日, "Y")",
  "syntax_detail": "TODAY() - 現在の日付を取得 + YEAR(date) - 日付から年を抽出 + MONTH(date) - 日付から月を抽出 + DAY(date) - 日付から日を抽出 + DATEDIF(start_date, end_date, "Y") - 2つの日付間の年数を計算",
  "category": "日付関数",
  "spreadsheet_data": [
    [
      {"v": "社員名", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "生年月日", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "年齢", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "今日の日付", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "年", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "月", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "日", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      null
    ],
    [
      {"v": "田中", "ct": {"t": "s"}},
      {"v": "1990-05-15", "ct": {"t": "s"}},
      {"v": null, "f": "=DATEDIF("1990-05-15",TODAY(),"Y")", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=TODAY()", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=YEAR(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=MONTH(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=DAY(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      null
    ],
    [
      {"v": "佐藤", "ct": {"t": "s"}},
      {"v": "1985-11-20", "ct": {"t": "s"}},
      {"v": null, "f": "=DATEDIF(DATE(1985,11,20),TODAY(),"Y")", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=TODAY()", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=YEAR(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=MONTH(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=DAY(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      null
    ],
    [
      {"v": "鈴木", "ct": {"t": "s"}},
      {"v": "1992-02-10", "ct": {"t": "s"}},
      {"v": null, "f": "=DATEDIF(DATE(1992,2,10),TODAY(),"Y")", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=TODAY()", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=YEAR(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=MONTH(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=DAY(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      null
    ],
    [
      {"v": "山田", "ct": {"t": "s"}},
      {"v": "1995-08-30", "ct": {"t": "s"}},
      {"v": null, "f": "=DATEDIF(DATE(1995,8,30),TODAY(),"Y")", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=TODAY()", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=YEAR(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=MONTH(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=DAY(TODAY())", "bg": "#FFE0B2", "fc": "#D84315"},
      null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ],
  "examples": ["=TODAY()", "=YEAR(TODAY())", "=MONTH(TODAY())", "=DAY(TODAY())", "=DATEDIF(DATE(1990,5,15),TODAY(),"Y")"]
}

**INDEX関数とMATCH関数組み合わせの完全な例：**
{
  "function_name": "INDEX & MATCH",
  "description": "INDEX関数とMATCH関数を組み合わせてVLOOKUPより柔軟な検索を実現します",
  "syntax": "INDEX(配列, MATCH(検索値, 検索配列, 0))",
  "syntax_detail": "INDEX(array, row_num) - 配列から指定した位置の値を取得 + MATCH(lookup_value, lookup_array, [match_type]) - 検索値の位置を返す",
  "category": "検索関数",
  "spreadsheet_data": [
    [
      {"v": "社員ID", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "社員名", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "部署名", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "検索ID", "ct": {"t": "s"}, "bg": "#FFF8E1"},
      {"v": "検索結果", "ct": {"t": "s"}, "bg": "#FFF8E1"},
      null, null, null
    ],
    [
      {"v": 101, "ct": {"t": "n"}},
      {"v": "田中", "ct": {"t": "s"}},
      {"v": "営業部", "ct": {"t": "s"}},
      {"v": 102, "ct": {"t": "n"}},
      {"v": null, "f": "=INDEX(C2:C6,MATCH(D2,A2:A6,0))", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [
      {"v": 102, "ct": {"t": "n"}},
      {"v": "佐藤", "ct": {"t": "s"}},
      {"v": "開発部", "ct": {"t": "s"}},
      {"v": 103, "ct": {"t": "n"}},
      {"v": null, "f": "=INDEX(C2:C6,MATCH(D3,A2:A6,0))", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [
      {"v": 103, "ct": {"t": "n"}},
      {"v": "鈴木", "ct": {"t": "s"}},
      {"v": "人事部", "ct": {"t": "s"}},
      {"v": 104, "ct": {"t": "n"}},
      {"v": null, "f": "=INDEX(C2:C6,MATCH(D4,A2:A6,0))", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [
      {"v": 104, "ct": {"t": "n"}},
      {"v": "山田", "ct": {"t": "s"}},
      {"v": "総務部", "ct": {"t": "s"}},
      {"v": 105, "ct": {"t": "n"}},
      {"v": null, "f": "=INDEX(C2:C6,MATCH(D5,A2:A6,0))", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [
      {"v": 105, "ct": {"t": "n"}},
      {"v": "伊藤", "ct": {"t": "s"}},
      {"v": "マーケティング部", "ct": {"t": "s"}},
      {"v": 101, "ct": {"t": "n"}},
      {"v": null, "f": "=INDEX(C2:C6,MATCH(D6,A2:A6,0))", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ],
  "examples": ["=INDEX(C2:C6,MATCH(D2,A2:A6,0))", "=MATCH(102,A2:A6,0)", "=INDEX(B2:B6,MATCH(103,A2:A6,0))"]
}

**重要：INDEX+MATCH関数の注意事項**
- INDEX関数では配列範囲と検索範囲を正確に指定してください
- MATCH関数の第3引数は完全一致の場合は0を指定してください
- 循環参照を避けるため、検索結果列は別の列に配置してください
- INDEX(戻り値の配列, MATCH(検索値, 検索する配列, 0))の形式で使用してください

**プロジェクト管理表の完全な例（DATEDIF/NETWORKDAYS使用）：**
// eslint-disable-next-line no-useless-escape
{
  "function_name": "DATEDIF & NETWORKDAYS & IF",
  "description": "DATEDIF関数で期間計算、NETWORKDAYS関数で営業日数計算、IF関数で評価を行います",
  "syntax": "DATEDIF(開始日, 終了日, "D") + NETWORKDAYS(開始日, 終了日) + IF(条件, 真の場合, 偽の場合)",
  "syntax_detail": "DATEDIF(start_date, end_date, "D") - 日数差を計算 + NETWORKDAYS(start_date, end_date) - 営業日数を計算 + IF関数による評価判定",
  "category": "プロジェクト管理",
  "spreadsheet_data": [
    [
      {"v": "プロジェクト名", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "開始日", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "終了日", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "期間（日）", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "稼働日数", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "評価", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      null, null
    ],
    [
      {"v": "プロジェクトA", "ct": {"t": "s"}},
      {"v": "2023-01-01", "ct": {"t": "s"}},
      {"v": "2023-01-10", "ct": {"t": "s"}},
      {"v": null, "f": "=DATEDIF(B2,C2,""D"")"", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=NETWORKDAYS(B2,C2)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=IF(D2>=15,\"長期\",\"短期\")", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null
    ],
    [
      {"v": "プロジェクトB", "ct": {"t": "s"}},
      {"v": "2023-02-01", "ct": {"t": "s"}},
      {"v": "2023-02-20", "ct": {"t": "s"}},
      {"v": null, "f": "=DATEDIF(B3,C3,\"D\")", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=NETWORKDAYS(B3,C3)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=IF(D3>=15,\"長期\",\"短期\")", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null
    ],
    [
      {"v": "プロジェクトC", "ct": {"t": "s"}},
      {"v": "2023-03-01", "ct": {"t": "s"}},
      {"v": "2023-03-08", "ct": {"t": "s"}},
      {"v": null, "f": "=DATEDIF(B4,C4,\"D\")", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=NETWORKDAYS(B4,C4)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=IF(D4>=15,\"長期\",\"短期\")", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null
    ],
    [
      {"v": "合計", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      null, null,
      {"v": null, "f": "=SUM(D2:D4)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=SUM(E2:E4)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ],
  "examples": ["=DATEDIF(B2,C2,\\\"D\\\")", "=NETWORKDAYS(B2,C2)", "=IF(D2>=15,\\\"長期\\\",\\\"短期\\\")", "=SUM(D2:D4)"]
}

**重要：日付関数の対応状況**
- **DATEDIF関数**: 手動計算でサポートされています
  - 使用例: =DATEDIF(B2,C2,"D") - 日数差を計算
  - 単位: "D"（日）、"M"（月）、"Y"（年）
- **NETWORKDAYS関数**: 手動計算でサポートされています
  - 使用例: =NETWORKDAYS(B2,C2) - 平日のみの日数を計算
- **使用可能な関数**: IF、SUM、AVERAGE、MAX、MIN、COUNT、VLOOKUP、INDEX、MATCH、DATEDIF、NETWORKDAYSなど
- **文字列リテラル**: 数式内では\"でエスケープしてください（例：=IF(A1>10,\"大\",\"小\")）

**超重要：日付データの形式**
- **必須**: すべての日付データは必ず「YYYY-MM-DD」形式で統一してください
- **例**: 2023-01-01、2023-12-31
- **NG例**: 2023/01/01、01/01/2023、1-1-2023
- **理由**: 日付計算の確実性を保証するため、ISO 8601形式に統一
- **セルのタイプ**: {"v": "2023-01-01", "ct": {"t": "s"}} として文字列で保存

**DATEDIF・NETWORKDAYS関数の特別な要件：**
- **DATEDIF関数**: 勤続年数、期間計算に使用。単位は「Y」(年)、「M」(月)、「D」(日)
- **NETWORKDAYS関数**: 営業日数計算に使用。土日を自動的に除外
- **使用例**: =DATEDIF(C2,TODAY(),"Y") で入社日から現在までの勤続年数を計算
- **営業日例**: =NETWORKDAYS(E2,F2) で有給開始日から終了日までの営業日数を計算
- **今日の日付**: TODAY()関数を使用して現在日付を取得

**AND・OR論理関数の正しい構文：**
- **AND関数**: =AND(条件1, 条件2, ...) - すべての条件がTRUEの場合にTRUE
- **OR関数**: =OR(条件1, 条件2, ...) - いずれかの条件がTRUEの場合にTRUE
- **正しい例**: =IF(AND(B2<10,E2="売れ筋"),"要発注","充分")
- **誤った例**: =IF(AND(B2<10,"要発注","充分"),"結果","") ← 第2・第3引数が論理値ではない
- **文字列比較**: セル参照="文字列" の形式で比較（例：D2="季節商品"）
- **複合条件**: =IF(AND(条件,OR(条件1,条件2)),"結果1","結果2") のような入れ子も可能

これらの例のように、実用的で循環参照のないデータを生成してください。特にVLOOKUP関数では、検索範囲とセル参照を正確に指定してください。

**HLOOKUP関数の特別な要件：**
- **データ配置**: HLOOKUPは横方向検索なので、データは横方向に配置してください
- **正しい例**: 
  - 1行目: 商品名  ノートPC   タブレット  スマートフォン
  - 2行目: 価格    80000     45000      60000
  - 3行目: 在庫    10        20         15
- **使用例**: =HLOOKUP(E1,A1:D3,2,0) - 1行目で商品名を検索し、2行目の価格を返す
- **重要**: VLOOKUPと違い、検索キーは1行目、返り値は下の行から取得

**エラーハンドリング関数の必須使用：**
- **ユーザーが「IFNA」を含むリクエストをした場合、必ずIFNA関数を使用してください**
- **ユーザーが「IFERROR」を含むリクエストをした場合、必ずIFERROR関数を使用してください**
- **両方含む場合は、IFERROR(IFNA(...))の入れ子構造で使用してください**
- **IFNA使用例**: =IFNA(HLOOKUP(E1,A1:D3,2,0),"商品が見つかりません")
- **IFERROR使用例**: =IFERROR(HLOOKUP(E1,A1:D3,2,0),"エラーが発生しました")
- **組み合わせ例**: =IFERROR(IFNA(HLOOKUP(E1,A1:D3,2,0),"見つかりません"),"エラー")
- **目的**: #N/Aエラーや他のエラーを人にやさしいメッセージに変換
- **適用**: VLOOKUP、HLOOKUP、INDEX+MATCH等の検索関数と組み合わせる

**重要：複数の関数を使用する場合**
- function_nameは組み合わせ名を記載（例："HLOOKUP & IFNA & IFERROR"）
- **syntaxは実際に使用する組み合わせた関数式を記載**（例："=IFERROR(IFNA(HLOOKUP(検索値,範囲,行番号,0),\"見つかりません\"),\"エラー\")"）
- **syntax_detailは簡単な説明を記載**（詳細な引数説明は自動生成されます）
- **descriptionは組み合わせた関数全体の動作を説明**
- 例：「HLOOKUPで横方向検索を行い、#N/Aエラーの場合はIFNAで独自メッセージを表示、その他のエラーはIFERRORで処理する包括的なエラーハンドリング機能」


**重要な注意事項：**
- Structured Outputsにより自動的にJSON形式で出力されます
- 8行×8列の配列構造を必ず守ってください
- 循環参照を絶対に避けてください
- 数式は具体的なセル参照を使用してください（例：=SUM(B2:B5)）
- 数式内の文字列は適切にエスケープしてください
- **すべての関数セルは統一してオレンジ系の色を使用してください**: bg="#FFE0B2", fc="#D84315"

**PMT関数（財務）の完全な例：**
{
  "function_name": "PMT",
  "description": "定期的な支払額を計算します（ローンや投資の）",
  "syntax": "PMT(利率, 期間数, 現在価値, [将来価値], [支払期日])",
  "syntax_detail": "PMT(rate, nper, pv, [fv], [type]) - rate:期間あたりの利率, nper:支払回数, pv:現在価値, fv:将来価値, type:支払期日",
  "category": "財務関数",
  "spreadsheet_data": [
    [
      {"v": "ローン金額", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "年利率", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "期間（年）", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "月利率", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      {"v": "支払回数", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      {"v": "月額支払", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      null, null
    ],
    [
      {"v": 1000000, "ct": {"t": "n"}},
      {"v": 0.03, "ct": {"t": "n"}},
      {"v": 5, "ct": {"t": "n"}},
      {"v": null, "f": "=B2/12", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=C2*12", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=PMT(D2,E2,-A2)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null
    ],
    [
      {"v": 2000000, "ct": {"t": "n"}},
      {"v": 0.025, "ct": {"t": "n"}},
      {"v": 10, "ct": {"t": "n"}},
      {"v": null, "f": "=B3/12", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=C3*12", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=PMT(D3,E3,-A3)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ]
}

**CONCATENATE/CONCAT関数（文字列）の完全な例：**
{
  "function_name": "CONCATENATE",
  "description": "複数の文字列を結合します",
  "syntax": "CONCATENATE(文字列1, 文字列2, ...)",
  "syntax_detail": "CONCATENATE(text1, text2, ...) - text1, text2: 結合する文字列",
  "category": "文字列関数",
  "spreadsheet_data": [
    [
      {"v": "姓", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "名", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "フルネーム", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      {"v": "メールアドレス", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      null, null, null, null
    ],
    [
      {"v": "田中", "ct": {"t": "s"}},
      {"v": "太郎", "ct": {"t": "s"}},
      {"v": null, "f": "=CONCATENATE(A2,\" \",B2)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=CONCATENATE(B2,\".\",A2,\"@example.com\")", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null
    ],
    [
      {"v": "佐藤", "ct": {"t": "s"}},
      {"v": "花子", "ct": {"t": "s"}},
      {"v": null, "f": "=CONCATENATE(A3,\" \",B3)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=CONCATENATE(B3,\".\",A3,\"@example.com\")", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null, null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ]
}

**ROUND/ROUNDUP/ROUNDDOWN関数（数学）の完全な例：**
{
  "function_name": "ROUND",
  "description": "数値を指定した桁数に丸めます",
  "syntax": "ROUND(数値, 桁数)",
  "syntax_detail": "ROUND(number, num_digits) - number: 丸める数値, num_digits: 丸める桁数",
  "category": "数学関数",
  "spreadsheet_data": [
    [
      {"v": "元の数値", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "小数第1位", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      {"v": "整数", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      {"v": "切り上げ", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      {"v": "切り捨て", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      null, null, null
    ],
    [
      {"v": 3.14159, "ct": {"t": "n"}},
      {"v": null, "f": "=ROUND(A2,1)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=ROUND(A2,0)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=ROUNDUP(A2,1)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=ROUNDDOWN(A2,1)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [
      {"v": 123.456, "ct": {"t": "n"}},
      {"v": null, "f": "=ROUND(A3,1)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=ROUND(A3,0)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=ROUNDUP(A3,1)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=ROUNDDOWN(A3,1)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ]
}

**COUNT/COUNTA/COUNTIF関数（統計）の完全な例：**
{
  "function_name": "COUNT",
  "description": "数値セルの個数をカウントします",
  "syntax": "COUNT(範囲) / COUNTA(範囲) / COUNTIF(範囲, 条件)",
  "syntax_detail": "COUNT(range) - 数値セルをカウント, COUNTA(range) - 空白以外をカウント, COUNTIF(range, criteria) - 条件に合うセルをカウント",
  "category": "統計関数",
  "spreadsheet_data": [
    [
      {"v": "データ", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "分類", "ct": {"t": "s"}, "bg": "#E3F2FD"},
      {"v": "数値カウント", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      {"v": "全カウント", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      {"v": "A分類カウント", "ct": {"t": "s"}, "bg": "#FFE0B2"},
      null, null, null
    ],
    [
      {"v": 100, "ct": {"t": "n"}},
      {"v": "A", "ct": {"t": "s"}},
      {"v": null, "f": "=COUNT(A2:A6)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=COUNTA(A2:A6)", "bg": "#FFE0B2", "fc": "#D84315"},
      {"v": null, "f": "=COUNTIF(B2:B6,\"A\")", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [
      {"v": "テスト", "ct": {"t": "s"}},
      {"v": "B", "ct": {"t": "s"}},
      null, null, null, null, null, null
    ],
    [
      {"v": 200, "ct": {"t": "n"}},
      {"v": "A", "ct": {"t": "s"}},
      null, null, null, null, null, null
    ],
    [
      {"v": 300, "ct": {"t": "n"}},
      {"v": "B", "ct": {"t": "s"}},
      null, null, null, null, null, null
    ],
    [
      {"v": "", "ct": {"t": "s"}},
      {"v": "A", "ct": {"t": "s"}},
      null, null, null, null, null, null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ]
}

**重要な注意事項：**
- NETWORKDAYSなどの関数も必ず数式フィールド("f")を含めてください
- 例: {"v": null, "f": "=NETWORKDAYS(B2,C2)", "bg": "#FFE0B2", "fc": "#D84315"}
- すべての関数計算セルは必ず数式を含む形式で返してください
- 数式がないと単なる静的な値として扱われ、関数のデモになりません

実用的で教育的なデモデータを作成してください。`;

// Retry機能（指数バックオフ）
const retryWithBackoff = async <T>(
  fn: () => Promise<T>, 
  maxRetries = 3,
  baseDelay = 100
): Promise<T> => {
  for (let i = 0; i < maxRetries; i++) {
    try {
      return await fn();
    } catch (error) {
      if (i === maxRetries - 1) throw error;
      
      // 指数バックオフ（100ms, 200ms, 400ms）
      const delay = baseDelay * Math.pow(2, i);
      console.log(`リトライ ${i + 1}/${maxRetries} - ${delay}ms待機中...`);
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }
  throw new Error('Unreachable'); // TypeScript用
};

export const fetchExcelFunction = async (query: string): Promise<ExcelFunctionResponse> => {
  // APIキーが設定されていない場合はモック関数を返す
  if (!import.meta.env.VITE_OPENAI_API_KEY || import.meta.env.VITE_OPENAI_API_KEY === 'your_api_key_here') {
    console.warn('OpenAI API key not configured, using mock data');
    return getMockFunction(query);
  }

  try {
    const model = String(import.meta.env.VITE_OPENAI_MODEL ?? 'gpt-4o');
    
    // 一時的にStructured Outputsを無効化して従来の方法を使用
    const supportsStructuredOutputs = false;
    const isLegacyModel = false;
    
    console.log(`使用モデル: ${model}, Structured Outputs対応: ${supportsStructuredOutputs}`);
    
    // Structured Outputsでretryロジックを使用
    const response = await retryWithBackoff(async () => {
      const params: OpenAI.Chat.Completions.ChatCompletionCreateParamsNonStreaming = {
        model,
        messages: [
          {
            role: 'system' as const,
            content: SYSTEM_PROMPT
          },
          {
            role: 'user' as const,
            content: enhanceUserPrompt(query)
          }
        ],
        temperature: 0.1, // より一貫した結果のために温度を大幅に下げる
        seed: 12345, // 再現性のためのseed値
        max_tokens: 4000, // より長いレスポンスに対応
      };
      
      // Structured Outputs対応モデルの場合のみresponse_formatを追加
      if (supportsStructuredOutputs && !isLegacyModel) {
        const paramsWithResponseFormat = {
          ...params,
          response_format: {
            type: "json_schema" as const,
            json_schema: {
              name: "excel_function_response",
              schema: OPENAI_JSON_SCHEMA,
              strict: true
            }
          }
        };
        return await openai.chat.completions.create(paramsWithResponseFormat);
      }
      
      return await openai.chat.completions.create(params);
    });

    const content = response.choices[0]?.message?.content;
    if (!content) {
      throw new Error('OpenAI APIからレスポンスが取得できませんでした');
    }

    // Structured Outputs対応モデルかどうかでパース処理を分岐
    try {
      let rawData: unknown;
      
      if (supportsStructuredOutputs && !isLegacyModel) {
        // Structured OutputsでJSONが保証されているので直接パース
        console.log('Structured Output レスポンス:', content);
        rawData = JSON.parse(content) as Record<string, unknown>;
      } else {
        // 従来の方法：JSONを抽出してパース
        console.log('従来のJSONパース レスポンス:', content);
        
        // JSONの抽出を試行
        let jsonData: string;
        
        // まず```json```で囲まれているかチェック
        const codeBlockMatch = content.match(/```json\s*([\s\S]*?)\s*```/);
        if (codeBlockMatch) {
          jsonData = codeBlockMatch[1];
        } else {
          // 通常のJSONパターンを探す
          const jsonMatch = content.match(/\{[\s\S]*\}/);
          if (!jsonMatch) {
            throw new Error('有効なJSONが見つかりませんでした');
          }
          jsonData = jsonMatch[0];
        }
        
        // 一般的なJSON構文エラーを自動修正
        // eslint-disable-next-line no-control-regex
        jsonData = jsonData.replace(/[\u0000-\u001F\u007F-\u009F]/g, '');
        
        // 事前処理：明らかな問題を修正
        console.log('修正前のJSON:', jsonData.substring(0, 300) + '...');
        
        // 全関数タイプの二重引用符を事前に修正（包括的アプローチ）
        jsonData = jsonData
          // REPT関数の修正
          .replace(/("f":\s*"=REPT\()""/g, '$1\\"')
          .replace(/(\["=REPT\()""/g, '$1\\"')
          // TEXT関数の修正
          .replace(/("f":\s*"=TEXT\([^,]+,\s*)""/g, '$1\\"')
          .replace(/(\["=TEXT\([^,]+,\s*)""/g, '$1\\"')
          // IF関数の修正
          .replace(/("f":\s*"=IF\([^,]+,\s*)""/g, '$1\\"')
          .replace(/("f":\s*"=IF\([^,]+,\s*[^,]+,\s*)""/g, '$1\\"')
          .replace(/(\["=IF\([^,]+,\s*)""/g, '$1\\"')
          // CONCATENATE関数の修正
          .replace(/("f":\s*"=CONCATENATE\([^)]*?)""/g, '$1\\"')
          .replace(/(\["=CONCATENATE\([^)]*?)""/g, '$1\\"')
          // SUBSTITUTE関数の修正
          .replace(/("f":\s*"=SUBSTITUTE\([^,]+,\s*)""/g, '$1\\"')
          .replace(/("f":\s*"=SUBSTITUTE\([^,]+,\s*[^,]+,\s*)""/g, '$1\\"')
          // FIND、SEARCH関数の修正
          .replace(/("f":\s*"=(?:FIND|SEARCH)\()""/g, '$1\\"')
          .replace(/(\["=(?:FIND|SEARCH)\()""/g, '$1\\"')
          // 一般的な関数引数の修正
          .replace(/""(,\s*[^)]+\)")/g, '\\"$1')
          .replace(/""(\)")/g, '\\"$1');
        
        // より安全なJSON修正：JSONをパースできるまで修正を試行
        let attempts = 0;
        while (attempts < 5) {
          try {
            JSON.parse(jsonData);
            break; // パースに成功したら終了
          } catch (error) {
            attempts++;
            console.log(`JSON修正試行 ${attempts}:`, error);
            
            // 数式内の引用符を段階的に修正
            if (attempts === 1) {
              // 主要関数の二重引用符を適切にエスケープ（完全版）
              jsonData = jsonData.replace(/"f":\s*"=REPT\(""([^"]+)"",\s*([^)]+)\)"/g, '"f": "=REPT(\\"$1\\", $2)"');
              jsonData = jsonData.replace(/"f":\s*"=TEXT\(([^,]+),\s*""([^"]+)""\)"/g, '"f": "=TEXT($1, \\"$2\\")"');
              jsonData = jsonData.replace(/"f":\s*"=IF\(([^,]+),\s*""([^"]+)""/g, '"f": "=IF($1, \\"$2\\"');
              jsonData = jsonData.replace(/"f":\s*"=CONCATENATE\(([^)]*?)""([^"]*)""/g, '"f": "=CONCATENATE($1\\"$2\\"');
              jsonData = jsonData.replace(/"f":\s*"=SUBSTITUTE\(([^,]+),\s*""([^"]*)"",\s*""([^"]*)""/g, '"f": "=SUBSTITUTE($1, \\"$2\\", \\"$3\\")"');
            } else if (attempts === 2) {
              // examples配列内の二重引用符も修正
              jsonData = jsonData.replace(/"=REPT\(""([^"]+)"",\s*([^)]+)\)"/g, '"=REPT(\\"$1\\", $2)"');
              jsonData = jsonData.replace(/"=TEXT\(([^,]+),\s*""([^"]+)""\)"/g, '"=TEXT($1, \\"$2\\")"');
              jsonData = jsonData.replace(/"=IF\(([^,]+),\s*""([^"]+)""/g, '"=IF($1, \\"$2\\"');
              jsonData = jsonData.replace(/"=CONCATENATE\(([^)]*?)""([^"]*)""/g, '"=CONCATENATE($1\\"$2\\"');
              jsonData = jsonData.replace(/"=(?:FIND|SEARCH)\(""([^"]*)""/g, '"=FIND(\\"$1\\"');
            } else if (attempts === 3) {
              // より包括的な関数内二重引用符の修正
              jsonData = jsonData.replace(/""([^"]*)""/g, '\\"$1\\"');
            } else if (attempts === 4) {
              // 最後の手段：全ての二重引用符を適切にエスケープ
              jsonData = jsonData.replace(/([^\\])""/g, '$1\\"');
              jsonData = jsonData.replace(/^""/g, '\\"');
            }
            
            console.log(`修正後のJSON試行 ${attempts}:`, jsonData.substring(0, 200) + '...');
          }
        }
        
        console.log('修正後のJSON:', jsonData);
        
        // 最終的なJSONパースを試行
        try {
          rawData = JSON.parse(jsonData) as Record<string, unknown>;
        } catch (finalError) {
          console.error('最終JSONパースエラー:', finalError);
          console.error('失敗したJSON:', jsonData);
          
          // エラー位置周辺の文字を表示
          const errorMatch = String(finalError).match(/position (\d+)/);
          if (errorMatch) {
            const pos = parseInt(errorMatch[1]);
            const start = Math.max(0, pos - 50);
            const end = Math.min(jsonData.length, pos + 50);
            console.error(`エラー位置周辺 (${start}-${end}):`, jsonData.substring(start, end));
            console.error('エラー位置:', ' '.repeat(Math.max(0, pos - start)) + '^');
          }
          
          throw new Error(`JSONパースが完全に失敗しました: ${String(finalError)}`);
        }
      }
      
      // Zodでバリデーション（空文字をデフォルト値で補完）
      const rawDataRecord = rawData as Record<string, unknown>;
      const preprocessedData = {
        ...rawDataRecord,
        function_name: rawDataRecord.function_name ?? "Excel関数",
        description: rawDataRecord.description ?? "Excel関数の説明",
        syntax: rawDataRecord.syntax ?? "関数の構文",
        category: rawDataRecord.category ?? "Excel関数",
      };
      
      const functionData = ExcelFunctionResponseSchema.parse(preprocessedData);
      
      // 関数名から使用されている関数を抽出し、詳細な説明を自動生成
      const usedFunctions = extractFunctionsFromName(functionData.function_name ?? '');
      console.log('使用されている関数:', usedFunctions);
      
      if (usedFunctions.length > 0) {
        const generatedDetails = generateFunctionDetails(usedFunctions);
        if (generatedDetails) {
          // 生成された詳細説明でsyntax_detailを置き換え
          functionData.syntax_detail = generatedDetails;
          console.log('自動生成された関数詳細:', generatedDetails);
        }
      }
      
      console.log('バリデーション済みデータ:', functionData);
      return functionData;
      
    } catch (parseError) {
      console.error('JSON解析エラー:', parseError);
      console.error('レスポンス内容:', content);
      
      // Zodのバリデーションエラーの場合、詳細を表示
      if (parseError instanceof Error && parseError.name === 'ZodError') {
        console.error('Zodバリデーションエラーの詳細:', JSON.stringify((parseError as unknown as { issues: unknown }).issues, null, 2));
      }
      
      throw new Error('APIレスポンスの解析に失敗しました: ' + String(parseError));
    }

  } catch (error) {
    console.error('OpenAI API呼び出しエラー:', error);
    
    // エラー時は例外をそのまま投げる
    throw new Error('関数の生成に失敗しました。しばらく時間をおいて再度お試しください。');
  }
};