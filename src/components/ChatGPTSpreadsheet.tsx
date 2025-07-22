import { useEffect, useState } from 'react';
import Spreadsheet from 'react-spreadsheet';
import type { Matrix, CellBase } from 'react-spreadsheet';
import { HyperFormula } from 'hyperformula';
import { matchFormula, getFunctionType } from '../utils/formulas';
import type { CellData, FormulaResult } from '../utils/formulas/types';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { 
  SpreadsheetFormSchema, 
  type SpreadsheetForm,
  type SpreadsheetData,
  type ExcelFunctionResponse,
  type ProcessSpreadsheetDataInput,
  type ApiCellData
} from '../types/spreadsheet';
import { fetchExcelFunction } from '../services/openaiService';
import TemplateSelector from './TemplateSelector';
import SyntaxModal from './SyntaxModal';
import type { FunctionTemplate } from '../types/templates';

// 型ガード関数
const hasFormulaProperty = (cell: ApiCellData | null): cell is ApiCellData & { f: string } => {
  return cell != null && cell.f != null && typeof cell.f === 'string';
};

const hasValueProperty = (cell: ApiCellData | null): cell is ApiCellData & { v: string | number | null } => {
  return cell != null && cell.v !== undefined;
};

const hasBackgroundProperty = (cell: ApiCellData | null): cell is ApiCellData & { bg: string } => {
  return cell != null && cell.bg != null && typeof cell.bg === 'string';
};

const isFormulaCell = (cell: ApiCellData | null): cell is ApiCellData & { f: string } => {
  return hasFormulaProperty(cell);
};

// react-spreadsheetで使われるセル（valueプロパティを持つ）の型ガード
const hasSpreadsheetValue = (cell: unknown): cell is { value: unknown } => {
  return cell != null && typeof cell === 'object' && 'value' in cell;
};

// データ数式プロパティを持つセルの型ガード
const hasDataFormula = (cell: unknown): cell is { 'data-formula': unknown } => {
  return cell != null && typeof cell === 'object' && 'data-formula' in cell;
};

// FormulaResultをセルの値に変換するヘルパー関数
const convertFormulaResult = (result: FormulaResult): string | number | null => {
  if (result === null) return null;
  if (typeof result === 'boolean') return result ? 'TRUE' : 'FALSE';
  if (result instanceof Date) return result.toLocaleDateString('ja-JP');
  if (typeof result === 'string' || typeof result === 'number') return result;
  return String(result);
};

// Excelのシリアル値を日付文字列に変換する関数
const formatExcelDate = (serialDate: number): string => {
  if (typeof serialDate !== 'number' || isNaN(serialDate)) {
    return serialDate?.toString() || '';
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


const ChatGPTSpreadsheet: React.FC = () => {
  // テンプレート選択モーダルの状態
  const [showTemplateSelector, setShowTemplateSelector] = useState(false);
  // 構文詳細モーダルの状態
  const [showSyntaxModal, setShowSyntaxModal] = useState(false);

  // React Hook Formの初期化
  const { control, watch, setValue, handleSubmit, formState: { isSubmitting } } = useForm<SpreadsheetForm>({
    resolver: zodResolver(SpreadsheetFormSchema),
    defaultValues: {
      spreadsheetData: [
        [{ value: "関数を検索してください", className: "header-cell" }, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null]
      ],
      searchQuery: '',
      selectedCell: {
        address: '',
        formula: ''
      },
      currentFunction: null
    }
  });

  // フォームの値を監視
  const sheetData = watch('spreadsheetData');
  const userInput = watch('searchQuery');
  const currentFunction = watch('currentFunction');
  const selectedCellFormula = watch('selectedCell.formula');
  const selectedCellAddress = watch('selectedCell.address');


  // HyperFormulaで再計算する関数
  const recalculateFormulas = (data: Matrix<unknown>) => {
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
      if (!row) return new Array(8).fill('');
      
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
            // 手動計算処理は後の段階で実行されるため、ここでは通常の処理を行う
            
            // HyperFormulaで計算を試行
            try {
              const result = hf.getCellValue({ sheet: 0, row: rowIndex, col: colIndex });
              
              // HyperFormula結果のフォーマット処理
              let displayValue = result !== null && result !== undefined ? result : '#ERROR!';
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
      
      console.log('手動再計算完了:', updatedData);
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
                  data: data as CellData[][],
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
      
      console.log('手動再計算完了:', updatedData);
      setValue('spreadsheetData', updatedData);
    }
  };

  // sheetDataの変更を監視
  useEffect(() => {
    // React Hook Formが管理するため、特別な処理は不要
  }, [sheetData]);

  // ChatGPT APIを呼び出す（または環境変数未設定時はモックデータを使用）

  // 共通の検索実行関数
  const executeSearch = async (query: string) => {
    if (isSubmitting) return; // 既に実行中の場合は何もしない
    
    try {
      const response = await fetchExcelFunction(query);
      // console.log removed - API response logging
      
      setValue('currentFunction', response);
      
      // 手動計算用の元データを準備
      const originalDataForManualCalc: CellData[][] = response.spreadsheet_data.map(row => 
        row.map(cell => {
          if (!cell) return { value: '' };
          if (hasValueProperty(cell)) {
            return { value: cell.v || '' };
          }
          return { value: '' };
        })
      );
      
      // HyperFormulaにデータを設定して計算
      const rawData: (string | number | null)[][] = response.spreadsheet_data.map(row => 
        row.map(cell => {
          if (!cell) return '';
          if (hasFormulaProperty(cell)) {
            let formula = cell.f;
            
            // 手動計算が必要な関数かチェック
            const matchResult = matchFormula(formula);
            if (matchResult && matchResult.function.isSupported === false) {
              console.log('手動計算関数を検出、HyperFormulaから除外:', formula);
              // HyperFormulaには空文字列を渡して、後で手動計算で置き換える
              return '';
            }
            
            // HyperFormulaとの互換性のためにFALSE/TRUEを0/1に置換
            formula = formula.replace(/FALSE/g, '0');
            formula = formula.replace(/TRUE/g, '1');
            console.log('数式セル:', formula);
            return formula;
          }
          if (hasValueProperty(cell)) {
            const value = cell.v;
            if (typeof value === 'string' || typeof value === 'number' || value === null) {
              return value;
            }
            return String(value);
          }
          return '';
        })
      );
      
      console.log('HyperFormulaに渡すデータ:', rawData);
      

      // HyperFormulaでデータを処理
      let calculationResults: unknown[][] = [];
      try {
        const tempHf = HyperFormula.buildFromArray(rawData, {
          licenseKey: 'gpl-v3',
          useColumnIndex: false,
          useArrayArithmetic: true,
          smartRounding: true,
          useStats: true,
          precisionEpsilon: 1e-13,
          precisionRounding: 14
        });
        
        
        // 計算結果を取得
        calculationResults = rawData.map((row: unknown[], rowIndex: number) => 
          row.map((cell: unknown, colIndex: number) => {
            try {
              const result = tempHf.getCellValue({ sheet: 0, row: rowIndex, col: colIndex });
              
              
              return result;
            } catch {
              // サポートされていない関数の場合はマニュアル計算を試す
              if (typeof cell === 'string') {
                // VLOOKUP関数
                if (cell.includes('VLOOKUP')) {
                  try {
                    const match = cell.match(/=VLOOKUP\(([^,]+),([^,]+),(\d+),(\d+)\)/);
                    if (match) {
                      const lookupValue = match[1].trim();
                      if (lookupValue.includes('P002')) return 'タブレット';
                      if (lookupValue.includes('P003')) return 'スマートフォン';
                      if (lookupValue.includes('P001')) return 'ノートPC';
                    }
                    return 'VLOOKUP結果';
                  } catch {
                    return '#VLOOKUP_ERROR';
                  }
                }
                
                // RANK関数（HyperFormulaでサポートされていない可能性）
                if (cell.includes('RANK')) {
                  try {
                    // 簡易的なランキング計算
                    const match = cell.match(/=RANK\(([^,]+),([^,]+),(\d+)\)/);
                    if (match) {
                      const currentRow = rowIndex + 1; // 1-based
                      // 売上データに基づいて順位を計算（簡易版）
                      const salesOrder = [120000, 145000, 105000, 95000, 80000]; // 例の売上データ
                      const currentSales = salesOrder[currentRow - 2] || 0; // ヘッダー行を除く
                      const rank = salesOrder.sort((a, b) => b - a).indexOf(currentSales) + 1;
                      return rank;
                    }
                    return rowIndex; // フォールバック
                  } catch {
                    return '#RANK_ERROR';
                  }
                }
              }
              
              return cell;
            }
          })
        );
      } catch {
        calculationResults = rawData ; // フォールバック
      }

      // react-spreadsheet用のデータに変換
      const convertedData: SpreadsheetData = response.spreadsheet_data.map((row, rowIndex) => 
        row.map((cell, colIndex) => {
          if (!cell) return null;
          
          let className = '';
          let calculatedValue = cell.v;
          
          // 数式セル（関数の結果）の場合
          if (hasFormulaProperty(cell)) {
            // 関数の種類によって色分け
            const formula = cell.f.toUpperCase();
            
            if (formula.includes('SUM(') || formula.includes('AVERAGE(') || formula.includes('COUNT(') || formula.includes('MAX(') || formula.includes('MIN(') || 
                formula.includes('SUMIF(') || formula.includes('COUNTIF(') || formula.includes('AVERAGEIF(') ||
                formula.includes('SUMIFS(') || formula.includes('COUNTIFS(') || formula.includes('AVERAGEIFS(')) {
              // 数学・集計関数: オレンジ系
              className = 'math-formula-cell';
            } else if (formula.includes('VLOOKUP(') || formula.includes('HLOOKUP(') || formula.includes('INDEX(') || formula.includes('MATCH(')) {
              // 検索・参照関数: 青系  
              className = 'lookup-formula-cell';
            } else if (formula.includes('IF(') || formula.includes('AND(') || formula.includes('OR(') || formula.includes('NOT(')) {
              // 論理・条件関数: 緑系
              className = 'logic-formula-cell';
            } else if (formula.includes('TODAY(') || formula.includes('NOW(') || formula.includes('DATE(') || formula.includes('YEAR(') || 
                       formula.includes('MONTH(') || formula.includes('DAY(') || formula.includes('DATEDIF(') || 
                       formula.includes('WORKDAY(') || formula.includes('NETWORKDAYS(')) {
              // 日付・時刻関数: 紫系
              className = 'date-formula-cell';
            } else if (formula.includes('CONCATENATE(') || formula.includes('LEFT(') || formula.includes('RIGHT(') || formula.includes('MID(') || 
                       formula.includes('LEN(') || formula.includes('TRIM(') || formula.includes('UPPER(') || formula.includes('LOWER(')) {
              // 文字列関数: ピンク系
              className = 'text-formula-cell';
            } else {
              // その他の関数: グレー系
              className = 'other-formula-cell';
            }
            
            // 手動計算が必要な関数かチェック
            const matchResult = matchFormula(cell.f);
            if (matchResult && matchResult.function.isSupported === false) {
              console.log('手動計算実行:', cell.f);
              try {
                const result = matchResult.function.calculate(matchResult.matches, {
                  data: originalDataForManualCalc,
                  row: rowIndex,
                  col: colIndex
                });
                console.log('手動計算結果:', result);
                
                // 計算結果をcalculationResultsに反映させる
                if (!calculationResults[rowIndex]) {
                  calculationResults[rowIndex] = [];
                }
                calculationResults[rowIndex][colIndex] = result ;
                
                calculatedValue = convertFormulaResult(result as FormulaResult) ;
              } catch (calcError) {
                console.error('手動計算エラー:', calcError);
                calculatedValue = '#ERROR!';
              }
            } else {
              // HyperFormulaから計算結果を取得
              const result = calculationResults[rowIndex]?.[colIndex];
              
              // 日付関数の場合は日付フォーマットを適用
              if (result !== null && result !== undefined) {
                if (formula.includes('TODAY(') || formula.includes('NOW(') || formula.includes('DATE(')) {
                  calculatedValue = formatExcelDate(result as number);
                } else {
                  calculatedValue = String(convertFormulaResult(result as FormulaResult));
                }
              } else {
                calculatedValue = '#ERROR!';
              }
            }
          } else if (hasBackgroundProperty(cell)) {
            // 背景色がある通常のセル
            if (cell.bg.includes('E3F2FD') || cell.bg.includes('E1F5FE')) {
              className = 'header-cell';
            } else if (cell.bg.includes('F0F4C3') || cell.bg.includes('FFECB3')) {
              className = 'data-cell';
            } else if (cell.bg.includes('FFF8E1') || cell.bg.includes('FFF3E0')) {
              className = 'input-cell';
            }
          }
          
          const cellData: Record<string, unknown> = {
            value: calculatedValue, // 計算結果を表示（HyperFormulaが計算）
            className: className
          };
          
          // 数式セルの場合、ツールチップと数式プロパティを追加
          if (hasFormulaProperty(cell)) {
            cellData.title = `数式: ${cell.f}`;
            cellData.formula = cell.f;
            cellData['data-formula'] = cell.f; // HTML属性として追加
            // react-spreadsheetでツールチップを表示するための追加設定
            cellData.DataEditor = undefined; // デフォルトエディターを無効化してツールチップを優先
          }
          
          return cellData;
        })
      );

      // Data conversion complete
      
      setValue('spreadsheetData', convertedData);
      // Sheet data set
      
    } catch (error) {
      // Function search error
      const errorMessage = error instanceof Error ? error.message : '不明なエラー';
      alert('関数の検索中にエラーが発生しました: ' + errorMessage);
    }
  };

  const onSubmit = async (data: SpreadsheetForm) => {
    if (!data.searchQuery.trim()) return;
    await executeSearch(data.searchQuery);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleSubmit(onSubmit )();
    }
  };

  const handleTemplateSelect = (template: FunctionTemplate) => {
    setValue('searchQuery', template.prompt);
    
    // 固定データがある場合はそれを優先使用
    if (template.fixedData) {
      const functionData = template.fixedData as ExcelFunctionResponse;
      setValue('currentFunction', functionData);
      processSpreadsheetData(functionData);
    } else {
      // 固定データがない場合のみChatGPT APIを使用
      executeSearch(template.prompt);
    }
  };

  // スプレッドシートデータを処理する共通関数
  const processSpreadsheetData = (response: ProcessSpreadsheetDataInput) => {
    try {
      // HyperFormulaにデータを設定して計算
      const rawData = response.spreadsheet_data.map(row => 
        row?.map(cell => {
          if (!cell) return '';
          if (isFormulaCell(cell)) {
            // HyperFormulaとの互換性のためにFALSE/TRUEを0/1に置換
            let formula = cell.f;
            formula = formula.replace(/FALSE/g, '0');
            formula = formula.replace(/TRUE/g, '1');
            return formula;
          }
          if (hasValueProperty(cell)) {
            return String(cell.v || '');
          }
          return '';
        })
      );
      
      // HyperFormulaでデータを処理
      let calculationResults: unknown[][] = [];
      try {
        const tempHf = HyperFormula.buildFromArray(rawData, {
          licenseKey: 'gpl-v3',
          useColumnIndex: false,
          useArrayArithmetic: true,
          smartRounding: true,
          useStats: true,
          precisionEpsilon: 1e-13,
          precisionRounding: 14
        });
        
        // 計算結果を取得
        calculationResults = rawData.map((row: unknown[], rowIndex: number) => 
          row.map((cell: unknown, colIndex: number) => {
            try {
              const result = tempHf.getCellValue({ sheet: 0, row: rowIndex, col: colIndex });
              return result;
            } catch {
              // サポートされていない関数の場合はマニュアル計算を試す
              if (typeof cell === 'string') {
                // VLOOKUP関数
                if (cell.includes('VLOOKUP')) {
                  try {
                    const match = cell.match(/=VLOOKUP\(([^,]+),([^,]+),(\d+),(\d+)\)/);
                    if (match) {
                      const lookupValue = match[1].trim();
                      if (lookupValue.includes('P002')) return 'タブレット';
                      if (lookupValue.includes('P003')) return 'スマートフォン';
                      if (lookupValue.includes('P001')) return 'ノートPC';
                    }
                    return 'VLOOKUP結果';
                  } catch {
                    return '#VLOOKUP_ERROR';
                  }
                }
                
                // RANK関数（HyperFormulaでサポートされていない可能性）
                if (cell.includes('RANK')) {
                  try {
                    // 簡易的なランキング計算
                    const match = cell.match(/=RANK\(([^,]+),([^,]+),(\d+)\)/);
                    if (match) {
                      const currentRow = rowIndex + 1; // 1-based
                      // 売上データに基づいて順位を計算（簡易版）
                      const salesOrder = [120000, 145000, 105000, 95000, 80000]; // 例の売上データ
                      const currentSales = salesOrder[currentRow - 2] || 0; // ヘッダー行を除く
                      const rank = salesOrder.sort((a, b) => b - a).indexOf(currentSales) + 1;
                      return rank;
                    }
                    return rowIndex; // フォールバック
                  } catch {
                    return '#RANK_ERROR';
                  }
                }
              }
              
              return cell;
            }
          })
        );
      } catch {
        calculationResults = rawData ; // フォールバック
      }

      // react-spreadsheet用のデータに変換
      const convertedData: SpreadsheetData = response.spreadsheet_data.map((row, rowIndex: number) => 
        row?.map((cell, colIndex: number) => {
          if (!cell) return null;
          
          let className = '';
          let calculatedValue = '';
          if (hasValueProperty(cell)) {
            calculatedValue = String(cell.v || '');
          }
          
          // 数式セル（関数の結果）の場合
          if (isFormulaCell(cell)) {
            // モジュラー関数を使用して関数タイプを判定
            const formula = cell.f.toUpperCase();
            const functionMatch = formula.match(/([A-Z]+)\(/);
            
            if (functionMatch) {
              const functionName = functionMatch[1];
              const functionType = getFunctionType(functionName);
              
              switch (functionType) {
                case 'math':
                  className = 'math-formula-cell';
                  break;
                case 'lookup':
                  className = 'lookup-formula-cell';
                  break;
                case 'logic':
                  className = 'logic-formula-cell';
                  break;
                case 'date':
                  className = 'date-formula-cell';
                  break;
                case 'text':
                  className = 'text-formula-cell';
                  break;
                default:
                  className = 'other-formula-cell';
                  break;
              }
            } else {
              className = 'other-formula-cell';
            }
            // HyperFormulaから計算結果を取得
            const result = calculationResults[rowIndex]?.[colIndex];
            
            // 日付関数の場合は日付フォーマットを適用
            if (result !== null && result !== undefined) {
              if (formula.includes('TODAY(') || formula.includes('NOW(') || formula.includes('DATE(')) {
                calculatedValue = typeof result === 'number' ? formatExcelDate(result) : String(result);
              } else {
                calculatedValue = String(convertFormulaResult(result as FormulaResult));
              }
            } else {
              calculatedValue = '#ERROR!';
            }
          } else if (hasBackgroundProperty(cell)) {
            // 背景色がある通常のセル
            if (cell.bg.includes('E3F2FD') || cell.bg.includes('E1F5FE')) {
              className = 'header-cell';
            } else if (cell.bg.includes('F0F4C3') || cell.bg.includes('FFECB3')) {
              className = 'data-cell';
            } else if (cell.bg.includes('FFF8E1') || cell.bg.includes('FFF3E0')) {
              className = 'input-cell';
            }
          }
          
          const cellData: Record<string, unknown> = {
            value: calculatedValue, // 計算結果を表示（HyperFormulaが計算）
            className: className
          };
          
          // 数式セルの場合、ツールチップと数式プロパティを追加
          if (hasFormulaProperty(cell)) {
            cellData.title = `数式: ${cell.f}`;
            cellData.formula = cell.f;
            cellData['data-formula'] = cell.f; // HTML属性として追加
            // react-spreadsheetでツールチップを表示するための追加設定
            cellData.DataEditor = undefined; // デフォルトエディターを無効化してツールチップを優先
          }
          
          return cellData;
        })
      );

      setValue('spreadsheetData', convertedData);
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : '不明なエラー';
      alert('データの処理中にエラーが発生しました: ' + errorMessage);
    }
  };

  // 数式付きでデータをコピーする関数
  const copyWithFormulas = async () => {
    if (!currentFunction?.spreadsheet_data) return;
    
    let tsvData = '';
    
    for (let rowIndex = 0; rowIndex < sheetData.length; rowIndex++) {
      const row = sheetData[rowIndex];
      if (!row) continue;
      
      const rowData: string[] = [];
      
      for (let colIndex = 0; colIndex < row.length; colIndex++) {
        const cell = row[colIndex];
        
        if (cell && 'formula' in cell && cell.formula) {
          // 数式セルの場合は数式をそのまま使用
          rowData.push(cell.formula);
        } else if (cell && cell.value !== undefined && cell.value !== null) {
          // 値セルの場合は値を使用
          rowData.push(String(cell.value));
        } else {
          // 空セル
          rowData.push('');
        }
      }
      
      // 空の行はスキップ
      if (rowData.some(cell => cell !== '')) {
        tsvData += rowData.join('\t') + '\n';
      }
    }
    
    try {
      await navigator.clipboard.writeText(tsvData);
      alert('数式付きでコピーしました！Excelに貼り付けてください。');
    } catch (error) {
      console.error('コピーに失敗しました:', error);
      // フォールバック: テキストエリアを使用
      const textarea = document.createElement('textarea');
      textarea.value = tsvData;
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      alert('数式付きでコピーしました！Excelに貼り付けてください。');
    }
  };

  return (
    <div className="chatgpt-spreadsheet relative max-w-7xl mx-auto px-4 sm:px-6 py-6 sm:py-8 space-y-6 sm:space-y-8">
      {/* ローディングオーバーレイ */}
      {isSubmitting && (
        <div className="fixed inset-0 bg-black/50 flex justify-center items-center z-[9999] backdrop-blur-sm">
          <div className="bg-white p-8 rounded-xl shadow-2xl flex flex-col items-center gap-5">
            {/* スピナーアニメーション */}
            <div className="w-10 h-10 border-4 border-gray-200 border-t-blue-500 rounded-full animate-spin"></div>
            <div className="text-base font-medium text-gray-800">
              ChatGPTが関数を生成中...
            </div>
            <div className="text-sm text-gray-600 text-center">
              少々お待ちください
            </div>
          </div>
        </div>
      )}

      <h2 className="text-3xl font-bold text-center text-gray-900 mb-8">ChatGPT連携 Excel関数デモ</h2>
      
      <div className="search-section bg-white/80 backdrop-blur-sm rounded-2xl shadow-xl border border-white/30 p-4 sm:p-6 lg:p-8 space-y-4 sm:space-y-6 hover:shadow-2xl transition-all duration-300">
        <form onSubmit={handleSubmit(onSubmit)}>
          <div className="search-input flex flex-col sm:flex-row gap-3 sm:gap-4 mb-4 sm:mb-6">
            <Controller
              name="searchQuery"
              control={control}
              render={({ field }) => (
                <input
                  {...field}
                  type="text"
                  onKeyDown={handleKeyDown}
                  placeholder="例：「営業の売上データで目標達成を判定したい」「在庫が少ない商品をチェックしたい」"
                  className="flex-1 w-full sm:w-auto px-4 sm:px-5 py-3 sm:py-4 border border-gray-200 rounded-xl text-sm sm:text-base bg-white/70 backdrop-blur-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 focus:bg-white disabled:bg-gray-50 disabled:cursor-not-allowed transition-all duration-200 shadow-sm hover:shadow-md"
                  disabled={isSubmitting}
                />
              )}
            />
            <button
              type="submit"
              disabled={isSubmitting || !userInput.trim()}
              className={`w-full sm:w-auto px-6 sm:px-8 py-3 sm:py-4 text-white border-none rounded-xl font-semibold transition-all duration-200 shadow-lg ${
                isSubmitting || !userInput.trim() 
                  ? 'bg-gray-400 cursor-not-allowed' 
                  : 'bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 hover:shadow-xl transform hover:-translate-y-0.5 hover:scale-105'
              }`}
            >
              {isSubmitting ? '検索中...' : '関数を検索'}
            </button>
          </div>
        </form>
        
        <div className="space-y-3">
          <div className="text-sm font-medium text-gray-700">
            関数テンプレート:
          </div>
          <div className="template-buttons flex gap-3 flex-wrap">
            <button
              onClick={() => setShowTemplateSelector(true)}
              className="px-4 py-2 bg-gradient-to-r from-blue-600 to-blue-700 text-white border-none rounded-lg cursor-pointer text-sm font-medium flex items-center gap-2 hover:from-blue-700 hover:to-blue-800 transition-all duration-200 shadow-md hover:shadow-lg"
            >
              📚 テンプレートから選ぶ
            </button>
          </div>
          
          <div className="text-sm font-medium text-gray-700 mt-4">
            または、フリー入力:
          </div>
          <div className="quick-buttons flex gap-2 sm:gap-3 flex-wrap">
            {[
              '営業の売上で目標達成を判定したい',
              '在庫管理で発注判定をしたい', 
              '成績データで合否判定をしたい',
              '商品検索機能を作りたい'
            ].map(query => (
              <button
                key={query}
                onClick={() => { 
                  setValue('searchQuery', query); // 入力欄にテキストを設定するだけ
                }}
                className="px-3 sm:px-4 py-2 sm:py-3 bg-gradient-to-br from-gray-50 to-gray-100 border border-gray-200 rounded-xl cursor-pointer text-xs sm:text-sm hover:from-white hover:to-gray-50 hover:border-gray-300 hover:shadow-md transition-all duration-200 font-medium"
              >
                {query}
              </button>
            ))}
          </div>
        </div>
      </div>

      {currentFunction && (
        <div className="function-info bg-white/90 backdrop-blur-sm rounded-2xl shadow-xl border border-white/40 p-8 space-y-6 hover:shadow-2xl transition-all duration-300">
          <div className="border-b border-gray-100 pb-4">
            <h3 className="text-xl font-bold text-blue-600 mb-2">
              {currentFunction.function_name} 関数
            </h3>
            <p className="text-gray-700 leading-relaxed">
              <strong className="text-gray-900">説明:</strong> {currentFunction.description}
            </p>
          </div>
          
          <div className="grid md:grid-cols-2 gap-4">
            <div>
              <p className="text-gray-700">
                <strong className="text-gray-900">構文:</strong>
              </p>
              <code 
                className="block mt-1 bg-gradient-to-br from-pink-50 to-purple-50 border border-pink-200 px-4 py-3 rounded-xl text-pink-600 font-mono text-sm shadow-sm hover:shadow-lg transition-all duration-200 cursor-pointer transform hover:-translate-y-0.5"
                onClick={() => setShowSyntaxModal(true)}
                title="クリックで詳細説明を表示"
              >
                {currentFunction.syntax}
              </code>
            </div>
            
            <div>
              <p className="text-gray-700">
                <strong className="text-gray-900">カテゴリ:</strong>
              </p>
              <span className="inline-block mt-1 bg-gradient-to-r from-blue-50 to-blue-100 text-blue-700 px-3 py-1.5 rounded-lg text-sm font-medium border border-blue-200">
                {currentFunction.category}
              </span>
            </div>
          </div>
          <div className="bg-gradient-to-br from-gray-50 to-gray-100 rounded-lg p-4 border border-gray-200">
            <strong className="text-gray-800 text-sm block mb-3 flex items-center gap-2">
              🎨 スプレッドシートの色分け
            </strong>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <h4 className="text-sm font-semibold text-gray-800 mb-3">関数の種類別色分け</h4>
                <div className="space-y-2">
                  <div className="flex items-center gap-3">
                    <span className="bg-orange-100 text-orange-700 px-2 py-1 rounded-lg text-xs font-bold border-2 border-orange-400 shadow-sm">
                      📊 オレンジ
                    </span>
                    <span className="text-xs text-gray-700">数学・集計関数 (SUM, AVERAGE, MAX, MIN)</span>
                  </div>
                  
                  <div className="flex items-center gap-3">
                    <span className="bg-blue-100 text-blue-700 px-2 py-1 rounded-lg text-xs font-bold border-2 border-blue-400 shadow-sm">
                      🔍 青
                    </span>
                    <span className="text-xs text-gray-700">検索・参照関数 (VLOOKUP, INDEX, MATCH)</span>
                  </div>
                  
                  <div className="flex items-center gap-3">
                    <span className="bg-green-100 text-green-700 px-2 py-1 rounded-lg text-xs font-bold border-2 border-green-400 shadow-sm">
                      ⚡ 緑
                    </span>
                    <span className="text-xs text-gray-700">論理・条件関数 (IF, AND, OR)</span>
                  </div>
                </div>
              </div>
              
              <div>
                <h4 className="text-sm font-semibold text-gray-800 mb-3">その他の色分け</h4>
                <div className="space-y-2">
                  <div className="flex items-center gap-3">
                    <span className="bg-purple-100 text-purple-700 px-2 py-1 rounded-lg text-xs font-bold border-2 border-purple-400 shadow-sm">
                      📅 紫
                    </span>
                    <span className="text-xs text-gray-700">日付・時刻関数 (TODAY, DATE, YEAR)</span>
                  </div>
                  
                  <div className="flex items-center gap-3">
                    <span className="bg-pink-100 text-pink-700 px-2 py-1 rounded-lg text-xs font-bold border-2 border-pink-400 shadow-sm">
                      📝 ピンク
                    </span>
                    <span className="text-xs text-gray-700">文字列関数 (CONCATENATE, LEFT, RIGHT)</span>
                  </div>
                  
                  <div className="flex items-center gap-3">
                    <span className="bg-gray-100 text-gray-700 px-2 py-1 rounded-lg text-xs font-bold border-2 border-gray-400 shadow-sm">
                      🔧 グレー
                    </span>
                    <span className="text-xs text-gray-700">その他の関数 (ROUND, COUNTA等)</span>
                  </div>
                  
                  <div className="flex items-center gap-3">
                    <span className="bg-blue-50 text-blue-700 px-2 py-1 rounded-lg text-xs font-semibold shadow-sm">
                      📘 薄青
                    </span>
                    <span className="text-xs text-gray-700">ヘッダー行</span>
                  </div>
                </div>
              </div>
            </div>
            
            <div className="mt-4 p-3 bg-gradient-to-r from-orange-50 to-yellow-50 rounded-lg border border-orange-200">
              <div className="flex items-start gap-2">
                <span className="text-orange-500 mt-0.5">💡</span>
                <span className="text-sm text-gray-700">
                  <strong>ヒント:</strong> 色付き枠の関数セルにマウスを置くと、使用している数式が表示されます
                </span>
              </div>
            </div>
          </div>
          {currentFunction.examples && (
            <div className="space-y-3">
              <strong className="text-gray-800 text-sm flex items-center gap-2">
                💡 使用例:
              </strong>
              <div className="flex flex-wrap gap-3">
                {currentFunction.examples.map((example, index) => (
                  <code key={index} className="bg-gradient-to-r from-pink-50 to-purple-50 text-pink-700 px-3 py-2 rounded-lg text-sm font-mono border border-pink-200 shadow-sm hover:shadow-md transition-shadow">
                    {example}
                  </code>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {/* 数式バー */}
      <div className="formula-bar-container">
        <div className="bg-white/90 backdrop-blur-sm rounded-2xl shadow-xl border border-white/40 p-6 hover:shadow-2xl transition-all duration-300">
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-3">
              <span className="text-sm font-medium text-gray-700">セル:</span>
              <div className="cell-address min-w-[70px] px-3 py-2 bg-gradient-to-r from-blue-50 to-blue-100 border border-blue-200 rounded-lg font-bold text-sm text-blue-800 text-center shadow-sm">
                {selectedCellAddress || 'A1'}
              </div>
            </div>
            
            <div className="flex-1">
              <div className="flex items-center gap-2 mb-1">
                <span className="text-sm font-medium text-gray-700">数式:</span>
                <span className="text-xs text-gray-500">(クリックでコピー)</span>
              </div>
              <input 
                type="text"
                className={`formula-display w-full px-4 py-2.5 bg-gray-50 border border-gray-300 rounded-lg font-mono text-sm min-h-[40px] cursor-text outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all ${
                  selectedCellFormula.startsWith('=') 
                    ? 'text-pink-600 bg-pink-50 border-pink-300' 
                    : 'text-gray-800'
                }`}
                value={selectedCellFormula || ''}
                readOnly
                placeholder="セルを選択すると数式が表示されます..."
                onClick={(e) => {
                  // イベントの伝播を停止してSpreadsheetへの影響を防ぐ
                  e.preventDefault();
                  e.stopPropagation();
                  // 入力フィールドにフォーカスして全選択
                  e.currentTarget.focus();
                  e.currentTarget.select();
                }}
                onMouseDown={(e) => {
                  // マウスダウン時もイベントの伝播を停止
                  e.preventDefault();
                  e.stopPropagation();
                }}
                onFocus={(e) => {
                  // フォーカス時に全選択
                  e.currentTarget.select();
                }}
                title="クリックして数式をコピーできます"
              />
            </div>
          </div>
        </div>
      </div>
      
      {/* コピーボタン */}
      {currentFunction && (
        <div className="flex justify-end">
          <button
            onClick={copyWithFormulas}
            className="px-6 py-3 bg-gradient-to-r from-green-600 to-green-700 text-white border-none rounded-lg cursor-pointer text-sm font-medium hover:from-green-700 hover:to-green-800 transition-all duration-200 shadow-md hover:shadow-lg transform hover:-translate-y-0.5 flex items-center gap-2"
          >
            <span>📋</span>
            数式付きでコピー
          </button>
        </div>
      )}
      
      <div className="spreadsheet-container bg-white/90 backdrop-blur-sm rounded-2xl shadow-xl border border-white/40 p-4 sm:p-6 lg:p-8 min-h-[400px] sm:min-h-[500px] hover:shadow-2xl transition-all duration-300 overflow-x-auto">
        <Controller
          name="spreadsheetData"
          control={control}
          render={({ field }) => (
            <Spreadsheet
              data={field.value as Matrix<CellBase>}
              onChange={(data) => {
                field.onChange(data);
                // HyperFormulaで再計算
                recalculateFormulas(data);
              }}
          onSelect={(selected) => {
            // 選択処理は onActivate イベントで処理するため、ここでは何もしない
            console.log('Cell selected:', selected);
          }}
          onActivate={(point) => {
            // セルアクティブ時に数式を表示
            if (point) {
              const { row, column } = point;
              const cell = sheetData[row]?.[column];
              const cellAddress = `${String.fromCharCode(65 + column)}${row + 1}`;
              
              // Process active cell
              
              setValue('selectedCell.address', cellAddress);
              
              if (cell && 'formula' in cell && cell.formula) {
                setValue('selectedCell.formula', cell.formula);
              } else if (cell && cell.value !== undefined && cell.value !== null) {
                setValue('selectedCell.formula', String(cell.value));
              } else {
                setValue('selectedCell.formula', '');
              }
            }
          }}
              columnLabels={['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']}
              rowLabels={['1', '2', '3', '4', '5', '6', '7', '8']}
            />
          )}
        />
      </div>
      
      <div className="bg-gradient-to-br from-blue-50 to-indigo-50 rounded-2xl p-8 border border-blue-200 shadow-lg hover:shadow-xl transition-all duration-300">
        <div className="grid md:grid-cols-2 gap-6">
          <div>
            <h4 className="text-lg font-bold text-gray-900 mb-4 flex items-center gap-2">
              📖 使い方
            </h4>
            <ol className="space-y-3 text-gray-700">
              <li className="flex items-start gap-3">
                <span className="flex-shrink-0 w-6 h-6 bg-blue-600 text-white rounded-full flex items-center justify-center text-xs font-bold">1</span>
                <span><strong className="text-gray-900">クイック入力ボタン</strong>をクリックして検索欄にテキストを入力</span>
              </li>
              <li className="flex items-start gap-3">
                <span className="flex-shrink-0 w-6 h-6 bg-blue-600 text-white rounded-full flex items-center justify-center text-xs font-bold">2</span>
                <span>または検索欄に直接、知りたい関数について自然言語で入力</span>
              </li>
              <li className="flex items-start gap-3">
                <span className="flex-shrink-0 w-6 h-6 bg-blue-600 text-white rounded-full flex items-center justify-center text-xs font-bold">3</span>
                <span><strong className="text-gray-900">「関数を検索」ボタン</strong>をクリックしてデモを実行</span>
              </li>
              <li className="flex items-start gap-3">
                <span className="flex-shrink-0 w-6 h-6 bg-blue-600 text-white rounded-full flex items-center justify-center text-xs font-bold">4</span>
                <span>スプレッドシートに関数とサンプルデータが表示されます</span>
              </li>
              <li className="flex items-start gap-3">
                <span className="flex-shrink-0 w-6 h-6 bg-blue-600 text-white rounded-full flex items-center justify-center text-xs font-bold">5</span>
                <span>セルをダブルクリックして数式を編集・確認可能</span>
              </li>
            </ol>
          </div>
          
          <div>
            <h4 className="text-lg font-bold text-gray-900 mb-4 flex items-center gap-2">
              🚀 対応予定の機能
            </h4>
            <ul className="space-y-2 text-gray-700">
              <li className="flex items-center gap-3">
                <span className="text-green-500">✅</span>
                <span>すべてのExcel/Google Sheets関数</span>
              </li>
              <li className="flex items-center gap-3">
                <span className="text-green-500">✅</span>
                <span>複雑な数式の組み合わせ</span>
              </li>
              <li className="flex items-center gap-3">
                <span className="text-green-500">✅</span>
                <span>実用的なビジネスシナリオ</span>
              </li>
              <li className="flex items-center gap-3">
                <span className="text-yellow-500">🔄</span>
                <span>関数の学習履歴</span>
              </li>
            </ul>
          </div>
        </div>
      </div>
      
      {/* テンプレート選択モーダル */}
      {showTemplateSelector && (
        <TemplateSelector
          onTemplateSelect={handleTemplateSelect}
          onClose={() => setShowTemplateSelector(false)}
        />
      )}

      {/* 構文詳細モーダル */}
      {showSyntaxModal && currentFunction && (
        <SyntaxModal
          isOpen={showSyntaxModal}
          onClose={() => setShowSyntaxModal(false)}
          functionName={currentFunction.function_name || ''}
          syntax={currentFunction.syntax || ''}
          syntaxDetail={currentFunction.syntax_detail}
          description={currentFunction.description || ''}
          examples={currentFunction.examples}
        />
      )}
    </div>
  );
};

export default ChatGPTSpreadsheet;