// 関数処理の共通型定義

export interface CellData {
  value: unknown;
  formula?: string;
  [key: string]: unknown;
}

export type FormulaResult = number | string | boolean | Date | null | number[][] | string[][] | (number | string)[][];

export interface FormulaContext {
  data: CellData[][];
  row: number;
  col: number;
}

// カスタム関数のインターフェース
export interface CustomFormula {
  name: string;
  pattern: RegExp;
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => FormulaResult;
  isSupported?: boolean; // HyperFormulaでサポートされているか
}

// エラー定義
export const FormulaError = {
  VALUE: '#VALUE!',
  NAME: '#NAME!',
  REF: '#REF!',
  NUM: '#NUM!',
  DIV0: '#DIV/0!',
  NULL: '#NULL!',
  NA: '#N/A',
  CYCLE: '#CYCLE!'
} as const;

export type FormulaErrorType = typeof FormulaError[keyof typeof FormulaError];