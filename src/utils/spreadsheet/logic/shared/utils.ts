// 関数処理の共通ユーティリティ

import type { FormulaContext } from './types';
import { FormulaError } from './types';
import { parseDate as parseDateNew } from './dateUtils';

// セル参照（A1形式）から座標を取得
export const parseCellReference = (cellRef: string): { col: number; row: number } | null => {
  const match = cellRef.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  
  // 列番号を計算（A=0, B=1, ..., Z=25, AA=26, ...）
  let col = 0;
  for (let i = 0; i < match[1].length; i++) {
    col = col * 26 + (match[1].charCodeAt(i) - 64);
  }
  col--; // 0ベースに変換
  
  const row = parseInt(match[2]) - 1; // 0ベースに変換
  
  return { col, row };
};

// セル範囲（A1:B2形式）から座標範囲を取得
export const parseCellRange = (range: string): { start: { col: number; row: number }; end: { col: number; row: number } } | null => {
  const match = range.match(/^([A-Z]+\d+):([A-Z]+\d+)$/);
  if (!match) return null;
  
  const start = parseCellReference(match[1]);
  const end = parseCellReference(match[2]);
  
  if (!start || !end) return null;
  
  return { start, end };
};

// セル参照から値を取得
export const getCellValue = (cellRef: string, context: FormulaContext): unknown => {
  const coords = parseCellReference(cellRef);
  if (!coords) {
    console.log(`getCellValue: Failed to parse cell reference "${cellRef}"`);
    return FormulaError.REF;
  }
  
  const { col, row } = coords;
  if (row < 0 || row >= context.data.length || col < 0 || col >= context.data[0]?.length) {
    console.log(`getCellValue: Cell reference out of bounds "${cellRef}" (row: ${row}, col: ${col})`);
    return FormulaError.REF;
  }
  
  const cellData = context.data[row][col];
  
  // 配列データの場合、セルは直接値として格納されている
  // 文字列、数値、null、undefinedの場合はそのまま返す
  if (typeof cellData === 'string' || typeof cellData === 'number' || cellData === null || cellData === undefined) {
    return cellData;
  }
  
  // セルデータがオブジェクトの場合のみプロパティを探す
  if (cellData && typeof cellData === 'object') {
    return cellData.value ?? cellData.v ?? cellData._ ?? cellData;
  }
  
  return cellData;
};

// セル範囲から値の配列を取得
export const getCellRangeValues = (range: string, context: FormulaContext): unknown[] => {
  const rangeCoords = parseCellRange(range);
  if (!rangeCoords) return [];
  
  const { start, end } = rangeCoords;
  const values: unknown[] = [];
  
  for (let row = start.row; row <= end.row; row++) {
    for (let col = start.col; col <= end.col; col++) {
      if (row >= 0 && row < context.data.length && col >= 0 && col < context.data[0]?.length) {
        const value = context.data[row][col]?.value;
        if (value !== null && value !== undefined && value !== '') {
          values.push(value);
        }
      }
    }
  }
  
  return values;
};

// 日付文字列をDateオブジェクトに変換（互換性のため）
export const parseDate = (dateValue: unknown): Date | null => {
  return parseDateNew(String(dateValue));
};

// Dateオブジェクトをそのまま返す（互換性のため）
export const toDate = (dateObj: Date | null): Date | null => {
  return dateObj;
};

// 値を数値に変換
export const toNumber = (value: unknown): number | null => {
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const num = parseFloat(value);
    if (!isNaN(num)) return num;
  }
  if (value instanceof Date) {
    return excelDateSerial(value);
  }
  return null;
};

// DateオブジェクトをExcelシリアル値に変換
export const excelDateSerial = (date: Date): number => {
  const excelEpoch = new Date(1899, 11, 30);
  return Math.floor((date.getTime() - excelEpoch.getTime()) / (1000 * 60 * 60 * 24));
};

// 文字列の引用符を除去
export const unquoteString = (str: string): string => {
  if (str.startsWith('"') && str.endsWith('"')) {
    return str.slice(1, -1).replace(/\\"/g, '"');
  }
  return str;
};