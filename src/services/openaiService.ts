import OpenAI from 'openai';
import type { ExcelFunctionResponse } from '../types/spreadsheet';
import { getMockFunction } from './mockData';

// OpenAI クライアントの初期化
const openai = new OpenAI({
  apiKey: import.meta.env.VITE_OPENAI_API_KEY,
  dangerouslyAllowBrowser: true // ブラウザ環境でのAPI利用を許可
});

const SYSTEM_PROMPT = `あなたはExcel/スプレッドシート関数の専門家です。ユーザーの要求に基づいて、Excel/スプレッドシート関数の実用的なデモデータをJSON形式で生成してください。

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
- 循環参照（#CYCLE!エラーの原因）
- 不正なセル参照
- 配列の要素数不一致

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

この例のように、実用的で循環参照のないデータを生成してください。JSONのみを返してください。`;

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
          content: `ユーザー要求: "${query}"\n\n上記の要求に基づいて、実用的で循環参照のないスプレッドシートデータを含むJSONを生成してください。必ず8行8列の配列にし、数式は具体的なセル参照を使用してください。`
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
      const jsonMatch = content.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        throw new Error('有効なJSONが見つかりませんでした');
      }

      const functionData = JSON.parse(jsonMatch[0]) as ExcelFunctionResponse;
      
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
    
    // エラー時はモック関数にフォールバック
    console.warn('APIエラーのためモックデータを使用します');
    return getMockFunction(query);
  }
};