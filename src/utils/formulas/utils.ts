// 関数処理の共通ユーティリティ

import type { FormulaContext } from './types';
import { FormulaError } from './types';
import dayjs from 'dayjs';
import customParseFormat from 'dayjs/plugin/customParseFormat';

// dayjsプラグインを有効化
dayjs.extend(customParseFormat);

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
export const getCellValue = (cellRef: string, context: FormulaContext): any => {
  const coords = parseCellReference(cellRef);
  if (!coords) return FormulaError.REF;
  
  const { col, row } = coords;
  if (row < 0 || row >= context.data.length || col < 0 || col >= context.data[0]?.length) {
    return FormulaError.REF;
  }
  
  return context.data[row][col]?.value;
};

// セル範囲から値の配列を取得
export const getCellRangeValues = (range: string, context: FormulaContext): any[] => {
  const rangeCoords = parseCellRange(range);
  if (!rangeCoords) return [];
  
  const { start, end } = rangeCoords;
  const values: any[] = [];
  
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

// 日付文字列をdayjsオブジェクトに変換
export const parseDate = (dateValue: any): dayjs.Dayjs | null => {
  if (!dateValue) return null;
  
  // すでにdayjsオブジェクトの場合
  if (dayjs.isDayjs(dateValue)) return dateValue;
  
  // Dateオブジェクトの場合
  if (dateValue instanceof Date) return dayjs(dateValue);
  
  // 文字列の場合
  if (typeof dateValue === 'string') {
    // 対応する日付フォーマット
    const formats = [
      'YYYY-MM-DD',
      'YYYY/MM/DD',
      'MM/DD/YYYY',
      'DD/MM/YYYY',
      'YYYY-M-D',
      'YYYY/M/D',
      'M/D/YYYY',
      'D/M/YYYY'
    ];
    
    for (const format of formats) {
      const parsed = dayjs(dateValue, format, true); // strict mode
      if (parsed.isValid()) {
        return parsed;
      }
    }
    
    // フォーマット指定なしで試す
    const parsed = dayjs(dateValue);
    if (parsed.isValid()) {
      return parsed;
    }
  }
  
  // 数値の場合（Excelシリアル値）
  if (typeof dateValue === 'number') {
    // Excel epoch: 1900/1/1 (実際は1899/12/30)
    const excelEpoch = dayjs('1899-12-30');
    return excelEpoch.add(dateValue, 'day');
  }
  
  return null;
};

// dayjsオブジェクトをDateオブジェクトに変換（互換性のため）
export const toDate = (dayjsObj: dayjs.Dayjs | null): Date | null => {
  return dayjsObj ? dayjsObj.toDate() : null;
};

// 値を数値に変換
export const toNumber = (value: any): number | null => {
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