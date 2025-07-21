import OpenAI from 'openai';
import type { ExcelFunctionResponse } from '../types/spreadsheet';
import { getMockFunction } from './mockData';
import { enhanceUserPrompt } from './promptEnhancer';

// OpenAI クライアントの初期化
const openai = new OpenAI({
  apiKey: import.meta.env.VITE_OPENAI_API_KEY,
  dangerouslyAllowBrowser: true // ブラウザ環境でのAPI利用を許可
});

const SYSTEM_PROMPT = `あなたはExcel/スプレッドシート関数の専門家です。

【最重要】ユーザーの自由な要求に基づいて、確実に動作するスプレッドシートデータをJSON形式で生成してください。エラーが発生しない、実用的なデモデータを作成することが最も重要です。

以下の形式で厳密にJSONを返してください：
{
  "function_name": "関数名（例：SUM, AVERAGE, VLOOKUP）",
  "description": "関数の分かりやすい説明",
  "syntax": "関数の構文",
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
6. 数式セルは{"v": null, "f": "=数式", "bg": "背景色", "fc": "文字色"}の形式にしてください
7. 空セルはnullにしてください
8. 各行は必ず8つの要素を持つ配列にしてください

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

**循環参照の具体例（これを避ける）：**
- A1がB1を参照し、B1がA1を参照する
- SUM関数が自分自身のセルを含む範囲を参照する
- 合計行のセルが、その合計を含む範囲を参照する

**SUM関数の完全な例：**
{
  "function_name": "SUM",
  "description": "範囲内の数値の合計を計算します",
  "syntax": "SUM(range1, [range2], ...)",
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
      {"v": null, "f": "=B2-C2", "bg": "#E8F5E8", "fc": "#2E7D32"},
      null, null, null, null
    ],
    [
      {"v": "2月", "ct": {"t": "s"}},
      {"v": 120000, "ct": {"t": "n"}},
      {"v": 35000, "ct": {"t": "n"}},
      {"v": null, "f": "=B3-C3", "bg": "#E8F5E8", "fc": "#2E7D32"},
      null, null, null, null
    ],
    [
      {"v": "3月", "ct": {"t": "s"}},
      {"v": 150000, "ct": {"t": "n"}},
      {"v": 40000, "ct": {"t": "n"}},
      {"v": null, "f": "=B4-C4", "bg": "#E8F5E8", "fc": "#2E7D32"},
      null, null, null, null
    ],
    [
      {"v": "合計", "ct": {"t": "s"}, "bg": "#E8F5E8"},
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
  "syntax": "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
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
      {"v": null, "f": "=VLOOKUP(D2,A2:C4,2,0)", "bg": "#FFE0B2", "fc": "#D84315"},
      null, null, null
    ],
    [
      {"v": "P002", "ct": {"t": "s"}, "bg": "#F0F4C3"},
      {"v": "タブレット", "ct": {"t": "s"}, "bg": "#F0F4C3"},
      {"v": 45000, "ct": {"t": "n"}, "bg": "#F0F4C3"},
      {"v": "P003", "ct": {"t": "s"}, "bg": "#FFECB3"},
      {"v": null, "f": "=VLOOKUP(D3,A2:C4,2,0)", "bg": "#FFE0B2", "fc": "#D84315"},
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
  "examples": ["=VLOOKUP(\"P001\",A2:C4,2,0)", "=VLOOKUP(D2,A2:C4,3,0)"]
}

**IF関数の完全な例：**
{
  "function_name": "IF",
  "description": "条件に基づいて異なる値を返します",
  "syntax": "IF(logical_test, value_if_true, value_if_false)",
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
      {"v": null, "f": "=IF(B2>=60,\"合格\",\"不合格\")", "bg": "#4CAF50", "fc": "#FFFFFF"},
      {"v": null, "f": "=IF(B2>=90,\"優\",IF(B2>=80,\"良\",IF(B2>=70,\"可\",\"不可\")))", "bg": "#9C27B0", "fc": "#FFFFFF"},
      null, null, null, null
    ],
    [
      {"v": "佐藤", "ct": {"t": "s"}},
      {"v": 45, "ct": {"t": "n"}},
      {"v": null, "f": "=IF(B3>=60,\"合格\",\"不合格\")", "bg": "#4CAF50", "fc": "#FFFFFF"},
      {"v": null, "f": "=IF(B3>=90,\"優\",IF(B3>=80,\"良\",IF(B3>=70,\"可\",\"不可\")))", "bg": "#9C27B0", "fc": "#FFFFFF"},
      null, null, null, null
    ],
    [
      {"v": "鈴木", "ct": {"t": "s"}},
      {"v": 92, "ct": {"t": "n"}},
      {"v": null, "f": "=IF(B4>=60,\"合格\",\"不合格\")", "bg": "#4CAF50", "fc": "#FFFFFF"},
      {"v": null, "f": "=IF(B4>=90,\"優\",IF(B4>=80,\"良\",IF(B4>=70,\"可\",\"不可\")))", "bg": "#9C27B0", "fc": "#FFFFFF"},
      null, null, null, null
    ],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ],
  "examples": ["=IF(A1>10,\"大\",\"小\")", "=IF(B1=\"\",0,B1*2)"]
}

これらの例のように、実用的で循環参照のないデータを生成してください。特にVLOOKUP関数では、検索範囲とセル参照を正確に指定してください。

**重要：複数の関数を使用する場合**
- function_nameは主要な関数名または組み合わせ名を記載（例："IF & SUM"）
- 複数の関数を組み合わせた実用的なデータ構造にする
- 各関数の役割を明確にする

**JSON構文の注意点：**
- 数式内の文字列は必ずエスケープしてください（例："=IF(A1>10,\\"大\\",\\"小\\")"）
- 文字列内にダブルクォートがある場合は \\" を使用
- 正しいJSON形式を厳守してください

JSONのみを返してください。`;

export const fetchExcelFunction = async (query: string): Promise<ExcelFunctionResponse> => {
  // APIキーが設定されていない場合はモック関数を返す
  if (!import.meta.env.VITE_OPENAI_API_KEY || import.meta.env.VITE_OPENAI_API_KEY === 'your_api_key_here') {
    console.warn('OpenAI API key not configured, using mock data');
    return getMockFunction(query);
  }

  try {
    const response = await openai.chat.completions.create({
      model: import.meta.env.VITE_OPENAI_MODEL || 'gpt-3.5-turbo',
      messages: [
        {
          role: 'system',
          content: SYSTEM_PROMPT
        },
        {
          role: 'user',
          content: enhanceUserPrompt(query)
        }
      ],
      temperature: 0.3, // より一貫した結果のために温度を下げる
      max_tokens: 3000, // より長いレスポンスに対応
    });

    const content = response.choices[0]?.message?.content;
    if (!content) {
      throw new Error('OpenAI APIからレスポンスが取得できませんでした');
    }

    // JSONを解析
    try {
      // JSONの抽出を試行
      let jsonData;
      
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
      // まず、不正な制御文字を削除
      jsonData = jsonData.replace(/[\u0000-\u001F\u007F-\u009F]/g, '');
      
      // より安全なJSON修正：JSONをパースできるまで修正を試行
      let attempts = 0;
      while (attempts < 3) {
        try {
          JSON.parse(jsonData);
          break; // パースに成功したら終了
        } catch (error) {
          attempts++;
          // 数式内の引用符を段階的に修正
          if (attempts === 1) {
            // エスケープされていない引用符を修正
            jsonData = jsonData.replace(/"f":\s*"=([^"]*)"([^"]*)"([^"]*)"/g, '"f": "=$1\\"$2\\"$3"');
          } else if (attempts === 2) {
            // より広範囲の引用符を修正
            jsonData = jsonData.replace(/([^\\])"([^\\])/g, '$1\\"$2');
          }
        }
      }
      
      console.log('修正後のJSON:', jsonData);
      
      const functionData = JSON.parse(jsonData) as ExcelFunctionResponse;
      
      // データの検証
      if (!functionData.function_name || !functionData.spreadsheet_data) {
        throw new Error('不正なレスポンス形式です');
      }

      return functionData;
    } catch (parseError) {
      console.error('JSON解析エラー:', parseError);
      console.error('レスポンス内容:', content);
      throw new Error('APIレスポンスの解析に失敗しました');
    }

  } catch (error) {
    console.error('OpenAI API呼び出しエラー:', error);
    
    // エラー時は例外をそのまま投げる
    throw new Error('関数の生成に失敗しました。しばらく時間をおいて再度お試しください。');
  }
};