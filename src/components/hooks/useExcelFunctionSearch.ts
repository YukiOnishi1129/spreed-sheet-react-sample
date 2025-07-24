import { useCallback } from 'react';
import { matchFormula, ALL_FUNCTIONS, getFunctionType } from "../../utils/functions";
import type { CellData, FormulaContext, FormulaResult } from "../../utils/functions";
import type { SpreadsheetData, ExcelFunctionResponse } from '../../types/spreadsheet';
import { fetchExcelFunction } from '../../services/openaiService';
import {
  hasValueProperty,
  hasFormulaProperty,
  hasBackgroundProperty,
  isFormulaResult
} from '../utils/typeGuards';
import {
  convertFormulaResult
} from '../utils/conversions';

interface UseExcelFunctionSearchProps {
  isSubmitting: boolean;
  setValue: {
    (name: 'currentFunction', value: ExcelFunctionResponse): void;
    (name: 'spreadsheetData', value: SpreadsheetData): void;
  };
  setLoadingMessage: (message: string) => void;
}

export const useExcelFunctionSearch = ({ isSubmitting, setValue, setLoadingMessage }: UseExcelFunctionSearchProps) => {
  
  const executeSearch = useCallback(async (query: string) => {
    if (isSubmitting) return; // 既に実行中の場合は何もしない
    
    try {
      setLoadingMessage('AIがあなたのリクエストを分析しています');
      const response = await fetchExcelFunction(query);
      
      setLoadingMessage('スプレッドシートデータを準備しています');
      setValue('currentFunction', response);
      
      // カスタム関数システム用のデータ準備
      const cellData: CellData[][] = response.spreadsheet_data.map((row) => 
        row.map((cell) => {
          if (!cell) return { value: '' };
          
          
          if (hasFormulaProperty(cell)) {
            return { value: '', formula: cell.f };
          }
          if (hasValueProperty(cell)) {
            return { value: cell.v ?? '' };
          }
          return { value: '' };
        })
      );
      
      setLoadingMessage('数式を計算しています');
      
      // カスタム関数システムで数式を計算
      const calculatedData = await calculateCustomFormulas(cellData);
      
      setLoadingMessage('結果を整形しています');
      
      // 計算結果をSpreadsheetData形式に変換
      const processedData: SpreadsheetData = calculatedData.map((row, rowIndex) =>
        row.map((cell, colIndex) => {
          const originalCell = response.spreadsheet_data[rowIndex]?.[colIndex];
          
          // SpreadsheetDataの形式に合わせたセルオブジェクトを作成
          const result: {
            value?: string | number | null;
            formula?: string;
            className?: string;
            title?: string;
            'data-formula'?: string;
          } = {};
          
          // 計算結果または元の値を設定
          if (cell.formula) {
            // 数式セルの場合
            result.formula = cell.formula;
            result['data-formula'] = cell.formula;
            
            // 値の型チェックと変換
            if (typeof cell.value === 'string' || typeof cell.value === 'number' || cell.value === null) {
              result.value = cell.value;
            } else {
              result.value = String(cell.value);
            }
          } else {
            // 値セルの場合
            if (isFormulaResult(cell.value)) {
              const converted = convertFormulaResult(cell.value);
              if (typeof converted === 'string' || typeof converted === 'number' || converted === null) {
                result.value = converted;
              } else {
                result.value = String(converted);
              }
            } else {
              // 値の型チェック
              if (typeof cell.value === 'string' || typeof cell.value === 'number' || cell.value === null) {
                result.value = cell.value;
              } else {
                result.value = String(cell.value);
              }
            }
          }
          
          // 背景色情報を className として設定
          if (originalCell && hasBackgroundProperty(originalCell) && originalCell.bg) {
            // Tailwind CSSの動的クラスは機能しないので、特定の色に対してマッピング
            if (originalCell.bg === '#FFE0B2') {
              result.className = 'bg-orange-200';
            } else if (originalCell.bg === '#E3F2FD') {
              result.className = 'bg-blue-100';
            } else {
              result.className = `bg-[${originalCell.bg}]`;
            }
          }
          
          // 数式セルの場合、関数タイプに基づいてCSSクラスを設定
          if (cell?.formula) {
            const formulaWithoutEquals = cell.formula.startsWith('=') ? cell.formula.substring(1) : cell.formula;
            const functionMatch = formulaWithoutEquals.match(/^([A-Z_]+)\s*\(/);
            
            if (functionMatch) {
              const functionName = functionMatch[1];
              const functionType = getFunctionType(functionName);
              
              switch (functionType) {
                case 'math':
                  result.className = 'math-formula-cell';
                  break;
                case 'statistical':
                  result.className = 'math-formula-cell'; // 統計関数も数学系として扱う
                  break;
                case 'financial':
                  result.className = 'financial-formula-cell';
                  break;
                case 'text':
                  result.className = 'text-formula-cell';
                  break;
                case 'datetime':
                  result.className = 'date-formula-cell';
                  break;
                case 'logical':
                  result.className = 'logic-formula-cell';
                  break;
                case 'lookup':
                  result.className = 'lookup-formula-cell';
                  break;
                default:
                  result.className = 'other-formula-cell';
                  break;
              }
            }
          }
          
          return result;
        })
      );
      
      setValue('spreadsheetData', processedData);
      
    } catch (error) {
      console.error('Excel関数検索でエラーが発生しました:', error);
      setLoadingMessage('エラーが発生しました');
    }
  }, [isSubmitting, setValue, setLoadingMessage]);

  return { executeSearch };
};

// カスタム関数システムで数式を計算する関数
async function calculateCustomFormulas(data: CellData[][]): Promise<CellData[][]> {
  const result: CellData[][] = data.map(row => row.map(cell => ({ ...cell })));
  const maxIterations = 10; // 循環参照を防ぐための最大反復回数
  
  for (let iteration = 0; iteration < maxIterations; iteration++) {
    let hasChanges = false;
    
    for (let rowIndex = 0; rowIndex < result.length; rowIndex++) {
      for (let colIndex = 0; colIndex < result[rowIndex].length; colIndex++) {
        const cell = result[rowIndex][colIndex];
        
        if (cell?.formula && (cell.value === undefined || cell.value === '' || typeof cell.value === 'string')) {
          const calculatedValue = calculateFormula(cell.formula, result, rowIndex, colIndex);
          
          if (calculatedValue !== cell.value) {
            cell.value = calculatedValue;
            hasChanges = true;
          }
        }
      }
    }
    
    // 変更がなくなったら終了
    if (!hasChanges) {
      break;
    }
  }
  
  return result;
}

// 単一の数式を計算する関数
function calculateFormula(formula: string, data: CellData[][], currentRow: number, currentCol: number): FormulaResult {
  try {
    // 先頭の = を除去
    const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;
    
    // カスタム関数のマッチングを試行
    const matchResult = matchFormula(cleanFormula);
    
    if (matchResult) {
      // FormulaContextを作成
      const context: FormulaContext = {
        data,
        row: currentRow,
        col: currentCol
      };
      
      // 関数を実行
      const result = matchResult.function.calculate(matchResult.matches, context);
      return result;
    } else {
      return `#NAME?`; // 未対応の関数エラー
    }
  } catch (error) {
    console.error(`数式計算エラー: ${formula}`, error);
    return `#VALUE!`; // 計算エラー
  }
}

// 利用可能な関数の一覧を取得
export const getAvailableFunctions = () => {
  return ALL_FUNCTIONS.map(func => ({
    name: func.name,
    pattern: func.pattern.toString()
  }));
};

// 関数名で関数情報を取得
export const getFunctionInfo = (functionName: string) => {
  return ALL_FUNCTIONS.find(func => 
    func.name.toUpperCase() === functionName.toUpperCase()
  );
};

// サポートされている関数のカテゴリ情報
export const getSupportedFunctionCategories = () => {
  const categories: Record<string, string[]> = {
    '数学・三角関数': [],
    '統計関数': [],
    '文字列操作関数': [],
    '日付・時刻関数': [],
    '論理関数': [],
    '検索・参照関数': [],
    '情報関数': [],
    'データベース関数': [],
    '財務関数': [],
    'その他': []
  };
  
  ALL_FUNCTIONS.forEach(func => {
    // 関数名に基づいて適切なカテゴリに分類
    const name = func.name;
    if (['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'ROUND', 'ABS', 'SQRT', 'POWER', 'MOD', 'SIN', 'COS', 'TAN'].some(mathFunc => name.includes(mathFunc))) {
      categories['数学・三角関数'].push(name);
    } else if (['VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'LOOKUP'].includes(name)) {
      categories['検索・参照関数'].push(name);
    } else if (['IF', 'AND', 'OR', 'NOT'].includes(name)) {
      categories['論理関数'].push(name);
    } else if (['CONCATENATE', 'LEFT', 'RIGHT', 'MID', 'LEN', 'UPPER', 'LOWER'].some(textFunc => name.includes(textFunc))) {
      categories['文字列操作関数'].push(name);
    } else if (['DATE', 'TODAY', 'NOW', 'YEAR', 'MONTH', 'DAY'].some(dateFunc => name.includes(dateFunc))) {
      categories['日付・時刻関数'].push(name);
    } else if (['DSUM', 'DAVERAGE', 'DCOUNT'].some(dbFunc => name.includes(dbFunc))) {
      categories['データベース関数'].push(name);
    } else if (['MEDIAN', 'MODE', 'STDEV', 'VAR'].some(statFunc => name.includes(statFunc))) {
      categories['統計関数'].push(name);
    } else if (['ISBLANK', 'ISERROR', 'TYPE'].some(infoFunc => name.includes(infoFunc))) {
      categories['情報関数'].push(name);
    } else if (['PMT', 'PV', 'FV', 'NPV', 'IRR'].includes(name)) {
      categories['財務関数'].push(name);
    } else {
      categories['その他'].push(name);
    }
  });
  
  return categories;
};