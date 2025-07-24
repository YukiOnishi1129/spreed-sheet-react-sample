import type { Matrix } from 'react-spreadsheet';
import { HyperFormula } from 'hyperformula';
import type { SpreadsheetData, ExcelFunctionResponse } from '../../types/spreadsheet';
import { matchFormula } from "../../utils/functions";
import { 
  hasSpreadsheetValue, 
  hasDataFormula, 
  isFormulaCell,
} from './typeGuards';
import { 
  convertMatrixToCellData, 
} from './conversions';

// HyperFormulaで再計算する関数
export const recalculateFormulas = (
  data: Matrix<unknown>, 
  currentFunction: ExcelFunctionResponse | null,
  setValue: (name: 'spreadsheetData', value: SpreadsheetData) => void
) => {
  console.log('再計算開始:', data);
  
  if (!currentFunction?.spreadsheet_data) return;
  
  // すべての数式セルを特定し、HyperFormulaで再計算
  const formulaCells: {row: number, col: number, formula: string}[] = [];
  
  // 元のテンプレートから数式セルを特定
  currentFunction.spreadsheet_data.forEach((row, rowIndex) => {
    if (row) {
      row.forEach((cell, colIndex) => {
        if (isFormulaCell(cell)) {
          formulaCells.push({
            row: rowIndex,
            col: colIndex,
            formula: cell.f
          });
        }
      });
    }
  });
  
  if (formulaCells.length === 0) return;
  
  // HyperFormulaで計算するためのデータ準備
  const rawData: (string | number | null)[][] = data.map((row, rowIndex) => {
    if (!row) return Array.from({ length: 8 }, () => '');
    
    return row.map((cell, colIndex) => {
      // 数式セルの場合は数式を返す
      const formulaCell = formulaCells.find(f => f.row === rowIndex && f.col === colIndex);
      if (formulaCell) {
        return formulaCell.formula;
      }
      // 通常のセルの場合は値を返す
      if (hasSpreadsheetValue(cell)) {
        const value = cell.value;
        if (typeof value === 'string' || typeof value === 'number' || value === null) {
          return value;
        }
        return String(value);
      }
      return '';
    });
  });
  
  // HyperFormulaで再計算
  try {
    const hf = HyperFormula.buildFromArray(rawData, {
      licenseKey: 'gpl-v3',
      useColumnIndex: false,
      useArrayArithmetic: true,
      smartRounding: true,
      useStats: true,
      precisionEpsilon: 1e-13,
      precisionRounding: 14
    });
    
    // 更新されたデータを作成
    const updatedData: SpreadsheetData = data.map((row, rowIndex) => {
      if (!row) return [];
      
      return row.map((cell, colIndex) => {
        // 数式セルの場合は再計算結果を使用
        const formulaCell = formulaCells.find(f => f.row === rowIndex && f.col === colIndex);
        if (formulaCell) {
          // HyperFormulaで計算を試行
          try {
            const result = hf.getCellValue({ sheet: 0, row: rowIndex, col: colIndex });
            
            // HyperFormula結果のフォーマット処理
            let displayValue = result ?? '#ERROR!';
            const formula = formulaCell.formula;
            
            if (typeof result === 'number' && !isNaN(result)) {
              // TODAY関数の場合は日付文字列に変換
              if (formula.includes('TODAY()') && !formula.includes('DATEDIF(')) {
                const excelEpoch = new Date(1899, 11, 30);
                const date = new Date(excelEpoch.getTime() + result * 24 * 60 * 60 * 1000);
                displayValue = date.toLocaleDateString('ja-JP');
              }
            }
            
            return {
              value: typeof displayValue === 'string' || typeof displayValue === 'number' || displayValue === null ? displayValue : String(displayValue),
              formula: formulaCell.formula,
              title: `数式: ${formulaCell.formula}`,
              className: hasSpreadsheetValue(cell) && typeof cell === 'object' && 'className' in cell ? String(cell.className) : undefined,
              'data-formula': formulaCell.formula,
              DataEditor: undefined
            };
          } catch (error) {
            console.warn('HyperFormula計算エラー:', error, formulaCell.formula);
            
            return {
              value: '#ERROR!',
              formula: formulaCell.formula,
              title: `数式: ${formulaCell.formula}`,
              className: hasSpreadsheetValue(cell) && typeof cell === 'object' && 'className' in cell ? String(cell.className) : undefined,
              'data-formula': formulaCell.formula,
              DataEditor: undefined
            };
          }
        }
        
        // 通常のセルの場合、適切なCell型オブジェクトを返す
        if (hasSpreadsheetValue(cell)) {
          return {
            value: typeof cell.value === 'string' || typeof cell.value === 'number' || cell.value === null ? cell.value : String(cell.value),
            className: typeof cell === 'object' && 'className' in cell ? String(cell.className) : undefined,
            formula: typeof cell === 'object' && 'formula' in cell ? String(cell.formula) : undefined,
            title: typeof cell === 'object' && 'title' in cell ? String(cell.title) : undefined,
            'data-formula': hasDataFormula(cell) && cell['data-formula'] != null ? String(cell['data-formula']) : undefined,
            DataEditor: undefined
          };
        }
        
        return null;
      });
    });
    
    setValue('spreadsheetData', updatedData);
    
  } catch (error) {
    console.error('HyperFormula再計算エラー:', error);
    // エラー時は手動計算のみ実行
    const updatedData: SpreadsheetData = data.map((row, rowIndex) => {
      if (!row) return [];
      
      return row.map((cell, colIndex) => {
        const formulaCell = formulaCells.find(f => f.row === rowIndex && f.col === colIndex);
        if (formulaCell) {
          // モジュラー関数を使用した手動計算
          const matchResult = matchFormula(formulaCell.formula);
          if (matchResult) {
            try {
              const result = matchResult.function.calculate(matchResult.matches, {
                data: convertMatrixToCellData(data),
                row: rowIndex,
                col: colIndex
              });
              
              return {
                value: String(result),
                formula: formulaCell.formula,
                title: `数式: ${formulaCell.formula}`,
                className: hasSpreadsheetValue(cell) && typeof cell === 'object' && 'className' in cell ? String(cell.className) : undefined,
                'data-formula': formulaCell.formula,
                DataEditor: undefined
              };
            } catch (calcError) {
              console.warn('モジュラー関数計算エラー:', calcError, formulaCell.formula);
            }
          }
          // フォールバック：エラーとして処理
          return {
            value: '#ERROR!',
            formula: formulaCell.formula,
            title: `数式: ${formulaCell.formula}`,
            className: hasSpreadsheetValue(cell) && typeof cell === 'object' && 'className' in cell ? String(cell.className) : undefined,
            'data-formula': formulaCell.formula,
            DataEditor: undefined
          };
        }
        
        // 通常のセルの場合、適切なCell型オブジェクトを返す
        if (hasSpreadsheetValue(cell)) {
          return {
            value: typeof cell.value === 'string' || typeof cell.value === 'number' || cell.value === null ? cell.value : String(cell.value),
            className: typeof cell === 'object' && 'className' in cell ? String(cell.className) : undefined,
            formula: typeof cell === 'object' && 'formula' in cell ? String(cell.formula) : undefined,
            title: typeof cell === 'object' && 'title' in cell ? String(cell.title) : undefined,
            'data-formula': hasDataFormula(cell) && cell['data-formula'] != null ? String(cell['data-formula']) : undefined,
            DataEditor: undefined
          };
        }
        
        return null;
      });
    });
    
    setValue('spreadsheetData', updatedData);
  }
};