import type { Matrix } from 'react-spreadsheet';
import type { CellData, FormulaResult } from '../../utils/formulas/types';
import type { SpreadsheetData } from '../../types/spreadsheet';
import { hasSpreadsheetValue } from './typeGuards';

// Matrix<unknown>をCellData[][]に変換する関数
export const convertMatrixToCellData = (matrix: Matrix<unknown>): CellData[][] => {
  return matrix.map(row => 
    row ? row.map(cell => {
      if (hasSpreadsheetValue(cell)) {
        return { value: cell.value };
      }
      return { value: '' };
    }) : []
  );
};

// SpreadsheetDataをreact-spreadsheet互換のMatrixに安全に変換
export const convertToMatrixCellBase = (data: SpreadsheetData): Matrix<{ value: string | number | null } | undefined> => {
  return data.map(row => 
    row ? row.map(cell => {
      if (cell && typeof cell === 'object') {
        // react-spreadsheet互換のオブジェクトを作成
        const result: { value: string | number | null; [key: string]: unknown } = {
          value: cell.value !== undefined ? 
            (typeof cell.value === 'string' || typeof cell.value === 'number' || cell.value === null ? cell.value : String(cell.value)) 
            : ''
        };
        
        // 追加のプロパティがあれば含める
        if (cell.className) result.className = cell.className;
        if (cell.formula) result.formula = cell.formula;
        if (cell.title) result.title = cell.title;
        if (cell['data-formula']) result['data-formula'] = cell['data-formula'];
        
        return result;
      }
      return undefined; // nullではなくundefinedを返す
    }) : []
  );
};

// FormulaResultをセルの値に変換するヘルパー関数
export const convertFormulaResult = (result: FormulaResult): string | number | null => {
  if (result === null) return null;
  if (typeof result === 'boolean') return result ? 'TRUE' : 'FALSE';
  if (result instanceof Date) return result.toLocaleDateString('ja-JP');
  if (typeof result === 'string' || typeof result === 'number') return result;
  return String(result);
};

// Excelのシリアル値を日付文字列に変換する関数
export const formatExcelDate = (serialDate: number): string => {
  if (typeof serialDate !== 'number' || isNaN(serialDate)) {
    return serialDate?.toString() ?? '';
  }
  
  // Excelのシリアル値は1900年1月1日を1とする（ただし1900年はうるう年ではないのでずれがある）
  // 一般的な変換式を使用
  const excelEpoch = new Date(1899, 11, 30); // 1899年12月30日
  const date = new Date(excelEpoch.getTime() + serialDate * 24 * 60 * 60 * 1000);
  
  // 日付として有効かチェック
  if (isNaN(date.getTime())) {
    return serialDate.toString();
  }
  
  // 現在の年に近い値の場合のみ日付として表示
  const currentYear = new Date().getFullYear();
  if (date.getFullYear() < currentYear - 50 || date.getFullYear() > currentYear + 50) {
    return serialDate.toString();
  }
  
  // YYYY/MM/DD形式で返す
  return date.toLocaleDateString('ja-JP', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).replace(/\//g, '/');
};