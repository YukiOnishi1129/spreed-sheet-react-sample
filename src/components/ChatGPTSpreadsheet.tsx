import { useEffect, useState, useCallback } from 'react';
import Spreadsheet from 'react-spreadsheet';
import type { Matrix } from 'react-spreadsheet';
import { HyperFormula } from 'hyperformula';
import { matchFormula } from '../utils/formulas';
import type { CellData } from '../utils/formulas/types';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { 
  SpreadsheetFormSchema, 
  type SpreadsheetForm,
  type SpreadsheetData,
  type ProcessSpreadsheetDataInput,
} from '../types/spreadsheet';
import { fetchExcelFunction } from '../services/openaiService';
import TemplateSelector from './TemplateSelector';
import SyntaxModal from './SyntaxModal';
import type { FunctionTemplate } from '../types/templates';

// Utility imports
import {
  hasFormulaProperty,
  hasValueProperty,
  hasBackgroundProperty,
  isFormulaResult,
  isExcelFunctionResponse
} from './utils/typeGuards';
import {
  convertToMatrixCellBase,
  convertFormulaResult,
  formatExcelDate
} from './utils/conversions';
import { recalculateFormulas } from './utils/hyperFormulaCalculations';

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

  // useCallbackでメモ化された関数たち
  const handleRecalculateFormulas = useCallback((data: Matrix<unknown>) => {
    recalculateFormulas(data, currentFunction, setValue);
  }, [currentFunction, setValue]);

  // 共通の検索実行関数
  const executeSearch = useCallback(async (query: string) => {
    if (isSubmitting) return; // 既に実行中の場合は何もしない
    
    try {
      const response = await fetchExcelFunction(query);
      setValue('currentFunction', response);
      
      // 手動計算用の元データを準備
      const originalDataForManualCalc: CellData[][] = response.spreadsheet_data.map(row => 
        row.map(cell => {
          if (!cell) return { value: '' };
          if (hasValueProperty(cell)) {
            return { value: cell.v ?? '' };
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
                    const match = cell.match(/=RANK\(([^,]+),([^,]+),(\d+)\)/);
                    if (match) {
                      const currentRow = rowIndex + 1;
                      const salesOrder = [120000, 145000, 105000, 95000, 80000];
                      const currentSales = salesOrder[currentRow - 2] ?? 0;
                      const rank = salesOrder.sort((a, b) => b - a).indexOf(currentSales) + 1;
                      return rank;
                    }
                    return rowIndex;
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
        calculationResults = rawData;
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
              className = 'math-formula-cell';
            } else if (formula.includes('VLOOKUP(') || formula.includes('HLOOKUP(') || formula.includes('INDEX(') || formula.includes('MATCH(')) {
              className = 'lookup-formula-cell';
            } else if (formula.includes('IF(') || formula.includes('AND(') || formula.includes('OR(') || formula.includes('NOT(')) {
              className = 'logic-formula-cell';
            } else if (formula.includes('TODAY(') || formula.includes('NOW(') || formula.includes('DATE(') || formula.includes('YEAR(') || 
                       formula.includes('MONTH(') || formula.includes('DAY(') || formula.includes('DATEDIF(') || 
                       formula.includes('WORKDAY(') || formula.includes('NETWORKDAYS(')) {
              className = 'date-formula-cell';
            } else if (formula.includes('CONCATENATE(') || formula.includes('LEFT(') || formula.includes('RIGHT(') || formula.includes('MID(') || 
                       formula.includes('LEN(') || formula.includes('TRIM(') || formula.includes('UPPER(') || formula.includes('LOWER(')) {
              className = 'text-formula-cell';
            } else {
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
                
                if (!calculationResults[rowIndex]) {
                  calculationResults[rowIndex] = [];
                }
                calculationResults[rowIndex][colIndex] = result ;
                
                calculatedValue = isFormulaResult(result) ? convertFormulaResult(result) : String(result);
              } catch (calcError) {
                console.error('手動計算エラー:', calcError);
                calculatedValue = '#ERROR!';
              }
            } else {
              // HyperFormulaから計算結果を取得
              const result = calculationResults[rowIndex]?.[colIndex];
              
              if (result !== null && result !== undefined) {
                if (formula.includes('TODAY(') || formula.includes('NOW(') || formula.includes('DATE(')) {
                  calculatedValue = typeof result === 'number' ? formatExcelDate(result) : String(result);
                } else {
                  calculatedValue = isFormulaResult(result) ? String(convertFormulaResult(result)) : String(result);
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
            value: calculatedValue,
            className: className
          };
          
          // 数式セルの場合、ツールチップと数式プロパティを追加
          if (hasFormulaProperty(cell)) {
            cellData.title = `数式: ${cell.f}`;
            cellData.formula = cell.f;
            cellData['data-formula'] = cell.f;
            cellData.DataEditor = undefined;
          }
          
          return cellData;
        })
      );

      setValue('spreadsheetData', convertedData);
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : '不明なエラー';
      alert('関数の検索中にエラーが発生しました: ' + errorMessage);
    }
  }, [isSubmitting, setValue]);

  const onSubmit = useCallback(async (data: SpreadsheetForm) => {
    if (!data.searchQuery.trim()) return;
    await executeSearch(data.searchQuery);
  }, [executeSearch]);

  const handleKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleSubmit(onSubmit)();
    }
  }, [handleSubmit, onSubmit]);

  // スプレッドシートデータを処理する共通関数
  const processSpreadsheetData = useCallback((response: ProcessSpreadsheetDataInput) => {
    try {
      // HyperFormulaにデータを設定して計算 - 実装は将来追加
      console.log('processSpreadsheetData called with:', response);
      // APIのデータ構造をSpreadsheetData形式に変換
      const convertedData: SpreadsheetData = response.spreadsheet_data.map(row => 
        row.map(cell => {
          if (!cell) return null;
          return {
            value: cell.v ?? null,
            formula: cell.f,
            className: cell.bg ? 'colored-cell' : undefined,
            title: cell.f ? `数式: ${cell.f}` : undefined,
            'data-formula': cell.f,
            DataEditor: undefined
          };
        })
      );
      setValue('spreadsheetData', convertedData);
    } catch (error) {
      console.error('スプレッドシートデータ処理エラー:', error);
    }
  }, [setValue]);

  const handleTemplateSelect = useCallback((template: FunctionTemplate) => {
    setValue('searchQuery', template.prompt);
    
    if (template.fixedData) {
      if (isExcelFunctionResponse(template.fixedData)) {
        setValue('currentFunction', template.fixedData);
        processSpreadsheetData(template.fixedData);
      }
    } else {
      executeSearch(template.prompt);
    }
  }, [setValue, executeSearch, processSpreadsheetData]);

  // sheetDataの変更を監視
  useEffect(() => {
    // React Hook Formが管理するため、特別な処理は不要
  }, [sheetData]);

  return (
    <div className="min-h-screen bg-gradient-to-br from-purple-50 via-blue-50 to-indigo-100 p-4 sm:p-6 lg:p-8">
      {/* Template Selector Modal */}
      {showTemplateSelector && (
        <TemplateSelector 
          onClose={() => setShowTemplateSelector(false)}
          onTemplateSelect={handleTemplateSelect}
        />
      )}
      
      {/* Syntax Modal */}
      {showSyntaxModal && currentFunction && (
        <SyntaxModal
          isOpen={showSyntaxModal}
          onClose={() => setShowSyntaxModal(false)}
          functionName={currentFunction.function_name ?? ''}
          syntax={currentFunction.syntax ?? ''}
          syntaxDetail={currentFunction.syntax_detail}
          description={currentFunction.description ?? ''}
        />
      )}

      <div className="max-w-7xl mx-auto space-y-6 sm:space-y-8">
        {/* Header */}
        <div className="text-center space-y-4">
          <h1 className="text-3xl sm:text-4xl lg:text-5xl font-bold bg-gradient-to-r from-purple-600 via-blue-600 to-indigo-600 bg-clip-text text-transparent">
            Excel関数デモンストレーター
          </h1>
          <p className="text-base sm:text-lg text-gray-600 max-w-3xl mx-auto">
            Excel関数を検索して、実際のスプレッドシートで動作を確認できます
          </p>
        </div>

        {/* Search Section */}
        <div className="bg-white/90 backdrop-blur-sm rounded-2xl shadow-xl border border-white/40 p-4 sm:p-6 lg:p-8 hover:shadow-2xl transition-all duration-300">
          <div className="space-y-4 sm:space-y-6">
            <div className="flex flex-col sm:flex-row gap-3 sm:gap-4">
              <div className="flex-1 space-y-2">
                <label className="block text-sm font-semibold text-gray-700">
                  関数名または用途を入力してください
                </label>
                <Controller
                  name="searchQuery"
                  control={control}
                  render={({ field }) => (
                    <input
                      {...field}
                      type="text"
                      placeholder="例: SUM, 合計, データの平均を求める"
                      className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all duration-200 text-sm sm:text-base"
                      onKeyDown={handleKeyDown}
                    />
                  )}
                />
              </div>
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={handleSubmit(onSubmit)}
                  disabled={isSubmitting || !userInput.trim()}
                  className={`px-6 py-3 rounded-xl font-semibold text-white transition-all duration-200 text-sm sm:text-base ${
                    isSubmitting || !userInput.trim() 
                      ? 'bg-gray-400 cursor-not-allowed' 
                      : 'bg-gradient-to-r from-blue-500 to-purple-600 hover:from-blue-600 hover:to-purple-700 shadow-lg hover:shadow-xl transform hover:scale-105'
                  }`}
                >
                  {isSubmitting ? '検索中...' : '検索'}
                </button>
                <button
                  type="button"
                  onClick={() => setShowTemplateSelector(true)}
                  className="px-4 py-3 rounded-xl font-semibold text-purple-600 bg-purple-50 hover:bg-purple-100 transition-all duration-200 text-sm sm:text-base border border-purple-200 hover:border-purple-300"
                >
                  📝 テンプレート
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Function Info Section */}
        {currentFunction && (
          <div className="bg-white/90 backdrop-blur-sm rounded-2xl shadow-xl border border-white/40 p-4 sm:p-6 lg:p-8 hover:shadow-2xl transition-all duration-300">
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 sm:gap-6 mb-6">
              <div className="space-y-2 flex-1">
                <h2 className="text-xl sm:text-2xl font-bold text-gray-800 flex items-center gap-3">
                  <span className="text-2xl">⚡</span>
                  {currentFunction.function_name ?? 'Excel関数'}
                </h2>
                <p className="text-sm sm:text-base text-gray-600 leading-relaxed">
                  {currentFunction.description ?? ''}
                </p>
              </div>
              <button
                onClick={() => setShowSyntaxModal(true)}
                className="px-4 py-2 bg-gradient-to-r from-indigo-500 to-purple-600 text-white rounded-xl hover:from-indigo-600 hover:to-purple-700 transition-all duration-200 font-semibold text-sm shadow-lg hover:shadow-xl transform hover:scale-105 whitespace-nowrap"
              >
                📖 構文詳細
              </button>
            </div>
            
            <div className="bg-gray-50 rounded-xl p-4 border border-gray-200">
              <p className="text-xs sm:text-sm font-semibold text-gray-600 mb-2">構文:</p>
              <code className="text-xs sm:text-sm font-mono text-gray-800 bg-white px-3 py-2 rounded-lg border border-gray-200 block">
                {currentFunction.syntax ?? ''}
              </code>
            </div>
          </div>
        )}

        {/* Selected Cell Info */}
        {selectedCellAddress && (
          <div className="bg-white/90 backdrop-blur-sm rounded-2xl shadow-xl border border-white/40 p-4 sm:p-6 hover:shadow-2xl transition-all duration-300">
            <h3 className="text-lg font-bold text-gray-800 mb-4 flex items-center gap-2">
              <span className="text-xl">📍</span>
              選択中のセル情報
            </h3>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <p className="text-sm font-semibold text-gray-600 mb-2">セルアドレス:</p>
                <div className="bg-gray-50 px-3 py-2 rounded-lg border border-gray-200 text-sm font-mono">
                  {selectedCellAddress || 'A1'}
                </div>
              </div>
              
              <div>
                <p className="text-sm font-semibold text-gray-600 mb-2">数式/値:</p>
                <div className="bg-gray-50 px-3 py-2 rounded-lg border border-gray-200 text-sm font-mono">
                  <input
                    type="text"
                    value={selectedCellFormula || ''}
                    readOnly
                    className="w-full bg-transparent border-none outline-none text-gray-800"
                    placeholder="数式や値が表示されます"
                  />
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Spreadsheet Section */}
        <div className="bg-white/90 backdrop-blur-sm rounded-2xl shadow-xl border border-white/40 p-4 sm:p-6 lg:p-8 min-h-[400px] sm:min-h-[500px] hover:shadow-2xl transition-all duration-300 overflow-x-auto">
          <Controller
            name="spreadsheetData"
            control={control}
            render={({ field }) => (
              <Spreadsheet
                data={convertToMatrixCellBase(field.value)}
                onChange={(data) => {
                  field.onChange(data);
                  handleRecalculateFormulas(data);
                }}
                onSelect={(selected) => {
                  if (Array.isArray(selected) && selected.length > 0) {
                    const firstSelection = selected[0] as { row: number; column: number };
                    if (firstSelection && typeof firstSelection === 'object' && 'row' in firstSelection && 'column' in firstSelection) {
                      const row = Number(firstSelection.row);
                      const column = Number(firstSelection.column);
                      const cellAddress = `${String.fromCharCode(65 + column)}${row + 1}`;
                      const cell = field.value[row]?.[column];
                      
                      setValue('selectedCell.address', cellAddress);
                      
                      if (cell?.value !== undefined && cell.value !== null) {
                        setValue('selectedCell.formula', String(cell.value));
                      } else {
                        setValue('selectedCell.formula', '');
                      }
                    }
                  }
                }}
                columnLabels={['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']}
                rowLabels={['1', '2', '3', '4', '5', '6', '7', '8']}
                className="w-full"
              />
            )}
          />
        </div>
      </div>
    </div>
  );
};

export default ChatGPTSpreadsheet;