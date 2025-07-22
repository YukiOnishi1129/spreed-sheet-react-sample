import { z } from 'zod';

// セルの値スキーマ
export const CellSchema = z.object({
  value: z.union([z.string(), z.number(), z.null()]).optional(),
  formula: z.string().optional(),
  className: z.string().optional(),
  title: z.string().optional(),
  'data-formula': z.string().optional(),
  DataEditor: z.undefined().optional(),
}).nullable();

// スプレッドシート全体のデータスキーマ
export const SpreadsheetDataSchema = z.array(
  z.array(CellSchema)
);

// 検索フォームのスキーマ
export const SearchFormSchema = z.object({
  query: z.string().min(1, "検索ワードを入力してください"),
});

// セル選択状態のスキーマ
export const CellSelectionSchema = z.object({
  address: z.string(),
  formula: z.string(),
});

// ChatGPT APIのセルデータスキーマ（厳密なスキーマ）
export const APIResponseCellSchema = z.object({
  v: z.union([z.string(), z.number(), z.null()]).optional(), // 値
  f: z.string().optional(), // 数式
  ct: z.object({ t: z.enum(['s', 'n']) }).optional(), // セルタイプ
  bg: z.string().optional(), // 背景色
  fc: z.string().optional(), // フォントカラー
}).nullable();

// ChatGPT APIレスポンスのスキーマ（柔軟なスキーマ）
export const ExcelFunctionResponseSchema = z.object({
  function_name: z.string().optional(),
  description: z.string().optional(),
  syntax: z.string().optional(),
  syntax_detail: z.string().optional(),
  category: z.string().optional(),
  spreadsheet_data: z.array(
    z.array(APIResponseCellSchema)
  ),
  examples: z.array(z.string()).optional(),
});

// JSON Schema用のスキーマ定義（OpenAI Structured Outputs用）
export const OPENAI_JSON_SCHEMA = {
  type: "object",
  properties: {
    function_name: {
      type: "string",
      description: "Excel関数名（例：SUM, VLOOKUP, IF）"
    },
    description: {
      type: "string",
      description: "関数の分かりやすい説明"
    },
    syntax: {
      type: "string",
      description: "関数の構文"
    },
    syntax_detail: {
      type: "string",
      description: "構文の詳細説明（各引数の説明）"
    },
    category: {
      type: "string",
      description: "関数のカテゴリ"
    },
    spreadsheet_data: {
      type: "array",
      description: "8行×8列のスプレッドシートデータ",
      items: {
        type: "array",
        items: {
          type: ["object", "null"],
          properties: {
            v: {
              type: ["string", "number", "null"],
              description: "セルの値"
            },
            f: {
              type: "string",
              description: "数式（=で始まる）"
            },
            ct: {
              type: "object",
              properties: {
                t: {
                  type: "string",
                  enum: ["s", "n"],
                  description: "セルタイプ（s:文字列, n:数値）"
                }
              },
              required: ["t"],
              additionalProperties: false
            },
            bg: {
              type: "string",
              description: "背景色（#RRGGBB形式）"
            },
            fc: {
              type: "string",
              description: "フォント色（#RRGGBB形式）"
            }
          },
          additionalProperties: false
        },
        minItems: 8,
        maxItems: 8
      },
      minItems: 8,
      maxItems: 8
    },
    examples: {
      type: "array",
      items: {
        type: "string"
      },
      description: "使用例の配列"
    }
  },
  required: ["function_name", "description", "syntax", "category", "spreadsheet_data"],
  additionalProperties: false
} as const;

// フォーム全体のスキーマ
export const SpreadsheetFormSchema = z.object({
  spreadsheetData: SpreadsheetDataSchema,
  searchQuery: z.string(),
  selectedCell: CellSelectionSchema,
  currentFunction: ExcelFunctionResponseSchema.nullable(),
});

// 型定義をエクスポート
export type Cell = z.infer<typeof CellSchema>;
export type SpreadsheetData = z.infer<typeof SpreadsheetDataSchema>;
export type SearchForm = z.infer<typeof SearchFormSchema>;
export type CellSelection = z.infer<typeof CellSelectionSchema>;
export type ExcelFunctionResponse = z.infer<typeof ExcelFunctionResponseSchema>;
export type SpreadsheetForm = z.infer<typeof SpreadsheetFormSchema>;

// processSpreadsheetData関数用の型（ExcelFunctionResponseと同じ構造）
export type ProcessSpreadsheetDataInput = ExcelFunctionResponse;