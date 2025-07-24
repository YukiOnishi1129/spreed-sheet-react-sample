import { useState, useCallback } from 'react';
import Spreadsheet from 'react-spreadsheet';
import type { Matrix, Selection } from 'react-spreadsheet';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { 
  SpreadsheetFormSchema, 
  type SpreadsheetForm,
  type SpreadsheetData,
  type ProcessSpreadsheetDataInput,
} from '../types/spreadsheet';
import TemplateSelector from './TemplateSelector';
import SyntaxModal from './SyntaxModal';
import type { FunctionTemplate } from '../types/templates';

// Utility imports
import {
  isExcelFunctionResponse,
  isRangeSelection
} from './utils/typeGuards';
import {
  convertToMatrixCellBase,
} from './utils/conversions';
import { recalculateFormulas } from './utils/customFormulaCalculations';
import { useExcelFunctionSearch } from './hooks/useExcelFunctionSearch';

const ChatGPTSpreadsheet: React.FC = () => {
  // テンプレート選択モーダルの状態
  const [showTemplateSelector, setShowTemplateSelector] = useState(false);
  // 構文詳細モーダルの状態
  const [showSyntaxModal, setShowSyntaxModal] = useState(false);
  // ローディングメッセージの状態
  const [loadingMessage, setLoadingMessage] = useState('');

  // React Hook Formの初期化
  const { control, watch, setValue, handleSubmit, formState: { isSubmitting } } = useForm<SpreadsheetForm>({
    resolver: zodResolver(SpreadsheetFormSchema),
    defaultValues: {
      spreadsheetData: [
        [{ value: "Excel関数を作成してください", className: "header-cell" }, null, null, null, null, null, null, null],
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
  const userInput = watch('searchQuery');
  const currentFunction = watch('currentFunction');
  const selectedCellFormula = watch('selectedCell.formula');
  const selectedCellAddress = watch('selectedCell.address');

  // Custom hook for Excel function search
  const { executeSearch } = useExcelFunctionSearch({ 
    isSubmitting, 
    setValue,
    setLoadingMessage 
  });

  // useCallbackでメモ化された関数たち
  const handleRecalculateFormulas = useCallback((data: Matrix<unknown>) => {
    recalculateFormulas(data, currentFunction, setValue);
  }, [currentFunction, setValue]);

  const onSubmit = useCallback(async (data: SpreadsheetForm) => {
    if (!data.searchQuery.trim()) return;
    await executeSearch(data.searchQuery);
  }, [executeSearch]);

  const handleKeyDown = useCallback((e: React.KeyboardEvent) => {
    // Enterキーでの実行を無効化（textareaで改行できるように）
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault(); // 実行を防ぐ
    }
  }, []);

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

  // 数式付きでスプレッドシートをコピーする関数
  const copySpreadsheetWithFormulas = useCallback(() => {
    const spreadsheetData = watch('spreadsheetData');
    if (!spreadsheetData || spreadsheetData.length === 0) return;

    // TSV形式で数式を含むデータを作成
    const tsvData = spreadsheetData.map((row) => {
      return row.map((cell) => {
        if (!cell) return '';
        
        // 数式があるセルの場合は数式をそのまま使用
        if (cell.formula) {
          return cell.formula;
        } else if (cell['data-formula']) {
          return String(cell['data-formula']);
        } else if (cell.value !== undefined && cell.value !== null) {
          return String(cell.value);
        } else {
          return '';
        }
      }).join('\t');
    }).join('\n');

    // クリップボードにコピー
    navigator.clipboard.writeText(tsvData).then(() => {
      alert('数式付きスプレッドシートがクリップボードにコピーされました！\nExcelやGoogleスプレッドシートに貼り付けて関数を実行できます。');
    }).catch(() => {
      // フォールバック: テキストエリアを作成してコピー
      const textArea = document.createElement('textarea');
      textArea.value = tsvData;
      document.body.appendChild(textArea);
      textArea.select();
      try {
        document.execCommand('copy');
      } catch (err) {
        console.error('Copy failed:', err);
      }
      document.body.removeChild(textArea);
      alert('数式付きスプレッドシートがクリップボードにコピーされました！\nExcelやGoogleスプレッドシートに貼り付けて関数を実行できます。');
    });
  }, [watch]);

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

  // セル選択時のハンドラー
  const handleCellSelect = useCallback((selected: Selection, fieldValue: SpreadsheetData) => {
    // Selectionがnullまたは空の場合は何もしない
    if (!selected) return;
    
    // RangeSelectionの場合（型定義に基づく正しい実装）
    if (isRangeSelection(selected)) {
      const row = selected.range.start.row;
      const column = selected.range.start.column;
      const cellAddress = `${String.fromCharCode(65 + column)}${row + 1}`;
      const cell = fieldValue[row]?.[column];
      
      setValue('selectedCell.address', cellAddress);
      
      // まず数式があるかチェック（数式セルの場合）
      if (cell?.formula) {
        setValue('selectedCell.formula', cell.formula);
      } else if (cell?.['data-formula']) {
        // data-formulaプロパティに数式が保存されている場合
        setValue('selectedCell.formula', String(cell['data-formula']));
      } else if (cell?.value !== undefined && cell.value !== null) {
        // 通常の値セル
        setValue('selectedCell.formula', String(cell.value));
      } else {
        setValue('selectedCell.formula', '');
      }
    }
  }, [setValue]);


  return (
    <div className="min-h-screen bg-gradient-to-br from-purple-50 via-blue-50 to-indigo-100 p-4 sm:p-6 lg:p-8 relative">
      {/* 全画面ローディングオーバーレイ */}
      {isSubmitting && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="bg-white rounded-2xl p-8 shadow-2xl flex flex-col items-center gap-6">
            <div className="relative">
              <div className="w-16 h-16 border-4 border-blue-200 rounded-full animate-spin"></div>
              <div className="absolute inset-0 w-16 h-16 border-4 border-transparent border-t-blue-600 rounded-full animate-spin"></div>
            </div>
            <div className="text-center">
              <h3 className="text-lg font-semibold text-gray-800 mb-2">Excel関数を作成中...</h3>
              <p className="text-sm text-gray-600">{loadingMessage || 'AIがあなたのリクエストを分析しています'}</p>
            </div>
          </div>
        </div>
      )}
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
            Excel関数を作成して、実際のスプレッドシートで動作を確認できます
          </p>
        </div>

        {/* Search Section */}
        <div className="bg-white/90 backdrop-blur-sm rounded-2xl shadow-xl border border-white/40 p-4 sm:p-6 lg:p-8 hover:shadow-2xl transition-all duration-300">
          <div className="space-y-4 sm:space-y-6">
            <div className="space-y-4">
              <div className="space-y-2">
                <label className="block text-sm font-semibold text-gray-700">
                  関数名または用途を入力してください
                </label>
                <Controller
                  name="searchQuery"
                  control={control}
                  render={({ field }) => (
                    <textarea
                      {...field}
                      rows={3}
                      placeholder="例: SUM, 合計, データの平均を求める&#10;&#10;複数行で詳しく説明することもできます"
                      className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all duration-200 text-sm sm:text-base resize-y min-h-[80px]"
                      onKeyDown={handleKeyDown}
                    />
                  )}
                />
              </div>
              
              {/* ボタンエリア */}
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                <div>
                  <button
                    type="button"
                    onClick={handleSubmit(onSubmit)}
                    disabled={isSubmitting || !userInput.trim()}
                    className={`px-8 py-3 rounded-xl font-semibold text-white transition-all duration-200 text-base ${
                      isSubmitting || !userInput.trim() 
                        ? 'bg-gray-400 cursor-not-allowed' 
                        : 'bg-gradient-to-r from-blue-500 to-purple-600 hover:from-blue-600 hover:to-purple-700 shadow-lg hover:shadow-xl transform hover:scale-105'
                    }`}
                  >
                    {isSubmitting ? (
                      <span className="flex items-center gap-2">
                        <svg className="animate-spin h-4 w-4" fill="none" viewBox="0 0 24 24">
                          <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                          <path className="opacity-75" fill="currentColor" d="m4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                        </svg>
                        作成中...
                      </span>
                    ) : '⚡ Excel関数を作成'}
                  </button>
                </div>
                
                {/* ヘルプエリア */}
                <div className="flex items-center gap-2 text-sm text-gray-500">
                  <span>迷ったら</span>
                  <button
                    type="button"
                    onClick={() => setShowTemplateSelector(true)}
                    className="px-3 py-1 rounded-lg font-medium text-purple-600 bg-purple-50 hover:bg-purple-100 transition-all duration-200 text-sm border border-purple-200 hover:border-purple-300"
                  >
                    📝 テンプレート
                  </button>
                  <span>をご利用ください</span>
                </div>
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
          {/* スプレッドシートヘッダー */}
          <div className="flex justify-between items-center mb-4">
            <h3 className="text-lg font-bold text-gray-800 flex items-center gap-2">
              <span className="text-xl">📊</span>
              スプレッドシート
            </h3>
            {currentFunction && (
              <button
                onClick={copySpreadsheetWithFormulas}
                className="px-4 py-2 bg-gradient-to-r from-green-500 to-emerald-600 text-white rounded-xl hover:from-green-600 hover:to-emerald-700 transition-all duration-200 font-semibold text-sm shadow-lg hover:shadow-xl transform hover:scale-105 flex items-center gap-2"
                title="数式付きでコピーして、ExcelやGoogleスプレッドシートに貼り付けできます"
              >
                📋 Excel用コピー
              </button>
            )}
          </div>
          
          {/* 色の凡例 */}
          <div className="mb-4 p-3 bg-gray-50 rounded-lg border border-gray-200">
            <p className="text-sm font-semibold text-gray-700 mb-2">セルの色の意味:</p>
            <div className="flex flex-wrap gap-3 text-xs">
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-blue-100 border border-gray-300 rounded"></div>
                <span>ヘッダー</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-yellow-50 border border-yellow-600 rounded"></div>
                <span>💰 財務関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-orange-100 border border-orange-500 rounded"></div>
                <span>📊 数学・三角関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-red-50 border border-red-600 rounded"></div>
                <span>📈 統計関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-blue-100 border border-blue-400 rounded"></div>
                <span>🔍 検索・参照関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-green-100 border border-green-400 rounded"></div>
                <span>⚡ 論理関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-purple-100 border border-purple-400 rounded"></div>
                <span>📅 日付・時刻関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-pink-100 border border-pink-400 rounded"></div>
                <span>📝 文字列関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-teal-50 border border-teal-600 rounded"></div>
                <span>ℹ️ 情報関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-amber-50 border border-amber-700 rounded"></div>
                <span>🗄️ データベース関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-emerald-100 border border-emerald-600 rounded"></div>
                <span>⚙️ エンジニアリング関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-indigo-100 border border-indigo-500 rounded"></div>
                <span>📊 キューブ関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-cyan-50 border border-cyan-600 rounded"></div>
                <span>🌐 Web関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-lime-50 border border-lime-600 rounded"></div>
                <span>📋 Google Sheets関数</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-4 h-4 bg-gray-100 border border-gray-400 rounded"></div>
                <span>🔧 その他</span>
              </div>
            </div>
          </div>
          
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
                onSelect={(selected) => handleCellSelect(selected, field.value)}
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