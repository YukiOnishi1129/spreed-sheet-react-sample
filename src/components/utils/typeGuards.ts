import type { ApiCellData, ExcelFunctionResponse } from '../../types/spreadsheet';
import type { FormulaResult } from '../../utils/formulas/types';

// 型ガード関数
export const hasFormulaProperty = (cell: ApiCellData | null): cell is ApiCellData & { f: string } => {
  return cell?.f != null && typeof cell.f === 'string';
};

export const hasValueProperty = (cell: ApiCellData | null): cell is ApiCellData & { v: string | number | null } => {
  return cell != null && cell.v !== undefined;
};

export const hasBackgroundProperty = (cell: ApiCellData | null): cell is ApiCellData & { bg: string } => {
  return cell?.bg != null && typeof cell.bg === 'string';
};

export const isFormulaCell = (cell: ApiCellData | null): cell is ApiCellData & { f: string } => {
  return hasFormulaProperty(cell);
};

// react-spreadsheetで使われるセル（valueプロパティを持つ）の型ガード
export const hasSpreadsheetValue = (cell: unknown): cell is { value: unknown } => {
  return cell != null && typeof cell === 'object' && 'value' in cell;
};

// データ数式プロパティを持つセルの型ガード
export const hasDataFormula = (cell: unknown): cell is { 'data-formula': unknown } => {
  return cell != null && typeof cell === 'object' && 'data-formula' in cell;
};

// FormulaResult型の型ガード
export const isFormulaResult = (value: unknown): value is FormulaResult => {
  return value === null || 
         typeof value === 'string' || 
         typeof value === 'number' || 
         typeof value === 'boolean' || 
         value instanceof Date;
};

// ExcelFunctionResponse型の型ガード
export const isExcelFunctionResponse = (value: unknown): value is ExcelFunctionResponse => {
  if (value == null || typeof value !== 'object') {
    return false;
  }
  
  if (!('spreadsheet_data' in value)) {
    return false;
  }
  
  // TypeScriptは'in'チェック後にプロパティの存在を認識する
  return Array.isArray(value.spreadsheet_data);
};