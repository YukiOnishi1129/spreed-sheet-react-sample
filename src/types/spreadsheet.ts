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

// ChatGPT APIレスポンスのスキーマ
export const ExcelFunctionResponseSchema = z.object({
  function_name: z.string(),
  description: z.string(),
  syntax: z.string(),
  category: z.string(),
  spreadsheet_data: z.array(z.array(z.any())),
  examples: z.array(z.string()).optional(),
});

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