import type { Matrix } from 'react-spreadsheet';
import type { SpreadsheetData, ExcelFunctionResponse } from '../../types/spreadsheet';
import { matchFormula } from "../../utils/spreadsheet/logic";
import type { CellData, FormulaContext, FormulaResult } from "../../utils/spreadsheet/logic";
import { 
  hasSpreadsheetValue, 
  hasDataFormula
} from './typeGuards';
import { formatCellValue } from '../../utils/spreadsheet/formatting';

// カスタム関数システムで再計算する関数
export const recalculateFormulas = (
  data: Matrix<unknown>, 
  currentFunction: ExcelFunctionResponse | null,
  setValue: (name: 'spreadsheetData', value: SpreadsheetData) => void
) => {
  if (!currentFunction?.spreadsheet_data) return;
  
  try {
    // Matrix<unknown>をCellData[][]に変換
    const cellData: CellData[][] = data.map((row) =>
      row.map((cell) => {
        if (!cell) return { value: '' };
        
        
        // data-formulaプロパティがある場合を優先
        if (hasDataFormula(cell)) {
          const formula = cell['data-formula'];
          return { 
            value: hasSpreadsheetValue(cell) ? (cell.value ?? '') : '',
            formula: typeof formula === 'string' ? formula : undefined
          };
        }
        
        // formulaプロパティを直接持つ場合
        if (typeof cell === 'object' && 'formula' in cell && cell.formula) {
          return {
            value: hasSpreadsheetValue(cell) ? (cell.value ?? '') : '',
            formula: typeof cell.formula === 'string' ? cell.formula : undefined
          };
        }
        
        // スプレッドシートの値を持つセル
        if (hasSpreadsheetValue(cell)) {
          return { value: cell.value ?? '' };
        }
        
        // 通常の値（数値はそのまま保持）
        if (typeof cell === 'string' || typeof cell === 'number' || cell === null) {
          return { value: cell };
        }
        
        // オブジェクトの場合、valueプロパティを探す
        if (cell && typeof cell === 'object' && 'value' in cell) {
          return { value: cell.value };
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
          
          
          // 日付のフォーマットを適用
          const formattedValue = formatCellValue(cell.value, cell.formula);
          if (typeof formattedValue === 'string' || typeof formattedValue === 'number' || formattedValue === null) {
            result.value = formattedValue;
          } else {
            result.value = String(formattedValue);
          }
        } else {
          // 非数式セルでも日付フォーマットを適用
          const formattedValue = formatCellValue(cell.value);
          if (typeof formattedValue === 'string' || typeof formattedValue === 'number' || formattedValue === null) {
            result.value = formattedValue;
          } else {
            result.value = String(formattedValue);
          }
        }
        
        // 元のセルの背景色を保持
        if (originalCell?.bg) {
          result.className = `bg-[${originalCell.bg}]`;
        }
        
        return result;
      })
    );
    
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
        
        if (cell.formula) {
          const calculatedValue = calculateSingleFormula(cell.formula, result, rowIndex, colIndex);
          
          
          // 常に計算結果を設定する
          cell.value = calculatedValue;
          hasChanges = true;
        }
      }
    }
    
    if (!hasChanges) {
      break;
    }
  }
  
  return result;
}

// ネストされた関数を評価する
function evaluateNestedFormula(formula: string, context: FormulaContext): string {
  let evaluatedFormula = formula;
  let hasChanges = true;
  let iterations = 0;
  const maxIterations = 10;
  
  
  while (hasChanges && iterations < maxIterations) {
    hasChanges = false;
    iterations++;
    
    // 最も内側の関数から評価する - より正確なパターンマッチング
    const innerFunctionRegex = /([A-Z]+(?:\.[A-Z]+)?)\(([^()]*)\)/;
    const matches = [...evaluatedFormula.matchAll(new RegExp(innerFunctionRegex, 'g'))];
    
    
    // 最も深い（最後の）マッチから処理
    for (let i = matches.length - 1; i >= 0; i--) {
      const match = matches[i];
      const fullMatch = match[0];
      const functionName = match[1];
      
      
      // この関数を評価
      const matchResult = matchFormula(fullMatch);
      
      
      if (matchResult) {
        try {
          const result = matchResult.function.calculate(matchResult.matches, context);
          
          // 結果を文字列に変換して置換
          let resultStr: string;
          if (result === null || result === undefined) {
            resultStr = '0';
          } else if (typeof result === 'number') {
            resultStr = result.toString();
          } else if (typeof result === 'string') {
            // エラー値の場合はそのまま、通常の文字列は引用符で囲む
            if (result.startsWith('#')) {
              resultStr = result;
            } else {
              resultStr = `"${result}"`;
            }
          } else {
            resultStr = String(result);
          }
          evaluatedFormula = evaluatedFormula.replace(fullMatch, resultStr);
          hasChanges = true;
          
        } catch (error) {
          console.error(`Error evaluating ${functionName}:`, error);
          return formula; // Return original formula on error
        }
      }
    }
  }
  
  
  return evaluatedFormula;
}

// 単一の数式を計算する関数
export function calculateSingleFormula(formula: string, data: CellData[][], currentRow: number, currentCol: number): FormulaResult {
  try {
    // 先頭の = を除去
    const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;
    
    // FormulaContextを作成
    const context: FormulaContext = {
      data,
      row: currentRow,
      col: currentCol
    };
    
    // ネストされた関数を評価
    const evaluatedFormula = evaluateNestedFormula(cleanFormula, context);
    
    // 評価後の式が引用符で囲まれた文字列の場合は、直接返す
    if (evaluatedFormula !== cleanFormula && evaluatedFormula.startsWith('"') && evaluatedFormula.endsWith('"')) {
      return evaluatedFormula.slice(1, -1);
    }
    
    // カスタム関数のマッチングを試行
    let matchResult = matchFormula(evaluatedFormula);
    
    // もし関数が見つからなかった場合
    if (!matchResult && evaluatedFormula !== cleanFormula) {
      // 引用符で囲まれた文字列の場合は、その値を返す
      if (evaluatedFormula.startsWith('"') && evaluatedFormula.endsWith('"')) {
        return evaluatedFormula.slice(1, -1);
      }
      
      // それ以外の場合のみ、元の式でも試す
      matchResult = matchFormula(cleanFormula);
    }
    
    if (matchResult) {
      // 関数を実行
      const result = matchResult.function.calculate(matchResult.matches, context);
      return result;
    } else {
      // 評価済みの式が単純な値の場合
      const num = parseFloat(evaluatedFormula);
      if (!isNaN(num)) {
        return num;
      }
      if (evaluatedFormula.startsWith('"') && evaluatedFormula.endsWith('"')) {
        return evaluatedFormula.slice(1, -1);
      }
      
      // エラー値の場合はそのまま返す
      if (evaluatedFormula.startsWith('#')) {
        return evaluatedFormula as FormulaResult;
      }
      
      console.warn(`未対応の数式: ${evaluatedFormula} (original: ${cleanFormula})`);
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