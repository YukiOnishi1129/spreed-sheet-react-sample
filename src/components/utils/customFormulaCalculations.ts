import type { Matrix } from 'react-spreadsheet';
import type { SpreadsheetData, ExcelFunctionResponse } from '../../types/spreadsheet';
import { matchFormula } from "../../utils/spreadsheet/logic";
import type { CellData, FormulaContext, FormulaResult } from "../../utils/spreadsheet/logic";
import { 
  hasSpreadsheetValue, 
  hasDataFormula, 
  isFormulaCell,
} from './typeGuards';

// カスタム関数システムで再計算する関数
export const recalculateFormulas = (
  data: Matrix<unknown>, 
  currentFunction: ExcelFunctionResponse | null,
  setValue: (name: 'spreadsheetData', value: SpreadsheetData) => void
) => {
  console.log('カスタム関数システムで再計算開始:', data);
  
  if (!currentFunction?.spreadsheet_data) return;
  
  try {
    // Matrix<unknown>をCellData[][]に変換
    const cellData: CellData[][] = data.map((row) =>
      row.map((cell) => {
        if (!cell) return { value: '' };
        
        // スプレッドシートの値を持つセル
        if (hasSpreadsheetValue(cell)) {
          return { value: cell.value ?? '' };
        }
        
        // 数式を持つセル
        if (isFormulaCell(cell)) {
          const formula = hasDataFormula(cell) ? cell['data-formula'] : undefined;
          return { 
            value: hasSpreadsheetValue(cell) ? (cell.value ?? '') : '',
            formula: typeof formula === 'string' ? formula : undefined
          };
        }
        
        // 通常の値
        if (typeof cell === 'string' || typeof cell === 'number' || cell === null) {
          return { value: cell };
        }
        
        return { value: String(cell) };
      })
    );
    
    // 数式セルを計算
    const calculatedData = calculateAllFormulas(cellData);
    
    // SpreadsheetData形式に変換
    const processedData: SpreadsheetData = calculatedData.map((row, rowIndex) =>
      row.map((cell, colIndex) => {
        const originalCell = currentFunction.spreadsheet_data[rowIndex]?.[colIndex];
        
        const result: {
          value?: string | number | null;
          formula?: string;
          className?: string;
          title?: string;
          'data-formula'?: string;
        } = {};
        
        // 計算結果を設定
        if (cell.formula) {
          result.formula = cell.formula;
          result['data-formula'] = cell.formula;
          
          if (typeof cell.value === 'string' || typeof cell.value === 'number' || cell.value === null) {
            result.value = cell.value;
          } else {
            result.value = String(cell.value);
          }
        } else {
          if (typeof cell.value === 'string' || typeof cell.value === 'number' || cell.value === null) {
            result.value = cell.value;
          } else {
            result.value = String(cell.value);
          }
        }
        
        // 元のセルの背景色を保持
        if (originalCell?.bg) {
          result.className = `bg-[${originalCell.bg}]`;
        }
        
        return result;
      })
    );
    
    console.log('再計算完了:', processedData);
    setValue('spreadsheetData', processedData);
    
  } catch (error) {
    console.error('再計算中にエラーが発生しました:', error);
  }
};

// 全ての数式を計算する関数
function calculateAllFormulas(data: CellData[][]): CellData[][] {
  const result: CellData[][] = data.map(row => row.map(cell => ({ ...cell })));
  const maxIterations = 10; // 循環参照を防ぐための最大反復回数
  
  for (let iteration = 0; iteration < maxIterations; iteration++) {
    let hasChanges = false;
    
    for (let rowIndex = 0; rowIndex < result.length; rowIndex++) {
      for (let colIndex = 0; colIndex < result[rowIndex].length; colIndex++) {
        const cell = result[rowIndex][colIndex];
        
        if (cell.formula && (cell.value === undefined || cell.value === '' || typeof cell.value === 'string')) {
          const calculatedValue = calculateSingleFormula(cell.formula, result, rowIndex, colIndex);
          
          if (calculatedValue !== cell.value) {
            cell.value = calculatedValue;
            hasChanges = true;
          }
        }
      }
    }
    
    if (!hasChanges) {
      break;
    }
  }
  
  return result;
}

// 単一の数式を計算する関数
function calculateSingleFormula(formula: string, data: CellData[][], currentRow: number, currentCol: number): FormulaResult {
  try {
    // 先頭の = を除去
    const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;
    
    // カスタム関数のマッチングを試行
    console.log(`Attempting to match formula: ${cleanFormula}`);
    const matchResult = matchFormula(cleanFormula);
    
    if (matchResult) {
      console.log(`数式 ${cleanFormula} を関数 ${matchResult.function.name} で計算中`);
      
      // FormulaContextを作成
      const context: FormulaContext = {
        data,
        row: currentRow,
        col: currentCol
      };
      
      // 関数を実行
      const result = matchResult.function.calculate(matchResult.matches, context);
      console.log(`計算結果:`, result);
      
      return result;
    } else {
      console.warn(`未対応の数式: ${cleanFormula}`);
      return `#NAME?`; // 未対応の関数エラー
    }
  } catch (error) {
    console.error(`数式計算エラー: ${formula}`, error);
    return `#VALUE!`; // 計算エラー
  }
}

// セルの依存関係を解析する関数（将来的な最適化用）
export function analyzeDependencies(data: CellData[][]): Map<string, string[]> {
  const dependencies = new Map<string, string[]>();
  
  data.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      if (cell.formula) {
        const cellAddress = `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`;
        const refs = extractCellReferences(cell.formula);
        dependencies.set(cellAddress, refs);
      }
    });
  });
  
  return dependencies;
}

// 数式からセル参照を抽出する関数
function extractCellReferences(formula: string): string[] {
  const cellRefPattern = /[A-Z]+\d+/g;
  const matches = formula.match(cellRefPattern);
  return matches ?? [];
}

// 計算順序を決定する関数（トポロジカルソート）
export function calculateOptimalOrder(dependencies: Map<string, string[]>): string[] {
  const visited = new Set<string>();
  const visiting = new Set<string>();
  const result: string[] = [];
  
  function visit(cell: string) {
    if (visiting.has(cell)) {
      console.warn(`循環参照を検出: ${cell}`);
      return;
    }
    
    if (visited.has(cell)) {
      return;
    }
    
    visiting.add(cell);
    
    const deps = dependencies.get(cell) ?? [];
    deps.forEach(dep => visit(dep));
    
    visiting.delete(cell);
    visited.add(cell);
    result.push(cell);
  }
  
  for (const cell of dependencies.keys()) {
    visit(cell);
  }
  
  return result;
}