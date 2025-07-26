import { useState, useEffect } from 'react';
import Spreadsheet, { type Matrix, type CellBase, type Selection } from 'react-spreadsheet';
import { demoSpreadsheetData, type DemoCategory } from '../data/demoSpreadsheetData';
import { allIndividualFunctionTests, type IndividualFunctionTest } from '../data/individualFunctionTests';
import { Link } from 'react-router-dom';
import { recalculateFormulas } from './utils/customFormulaCalculations';
import type { SpreadsheetData, ExcelFunctionResponse } from '../types/spreadsheet';

function DemoSpreadsheet() {
  const [selectedCategory, setSelectedCategory] = useState<DemoCategory>(demoSpreadsheetData[0]);
  const [spreadsheetData, setSpreadsheetData] = useState<Matrix<CellBase>>([]);
  const [selectedCell, setSelectedCell] = useState<{ row: number; col: number; formula?: string; value?: string | number | null } | null>(null);
  const [demoMode, setDemoMode] = useState<'grouped' | 'individual'>('grouped');
  const [selectedFunction, setSelectedFunction] = useState<IndividualFunctionTest | null>(null);
  const [isCalculating, setIsCalculating] = useState<boolean>(false);

  // カテゴリ選択時にデータを初期化と自動計算
  useEffect(() => {
    if (demoMode === 'grouped') {
      initializeAndCalculate();
    }
  }, [selectedCategory, demoMode]); // eslint-disable-line react-hooks/exhaustive-deps
  
  // 個別関数選択時の初期化
  useEffect(() => {
    if (demoMode === 'individual' && selectedFunction) {
      initializeIndividualFunction();
    }
  }, [selectedFunction, demoMode]); // eslint-disable-line react-hooks/exhaustive-deps

  const initializeAndCalculate = () => {
    setIsCalculating(true);
    
    console.log('Selected category:', selectedCategory.name, selectedCategory.id);
    
    const initialData: Matrix<CellBase> = selectedCategory.data.map((row) => 
      row.map((cellValue) => {
        // 数式が直接入っている場合
        if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
          return {
            value: '',
            formula: cellValue,
            'data-formula': cellValue
          };
        }
        return { value: cellValue ?? '' };
      })
    );
    
    // 直接計算を実行
    calculateFormulas(initialData);
    setSelectedCell(null);
  };
  
  const initializeIndividualFunction = () => {
    if (!selectedFunction) return;
    
    setIsCalculating(true);
    
    const initialData: Matrix<CellBase> = selectedFunction.data.map((row) => 
      row.map((cellValue) => {
        if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
          return {
            value: '',
            formula: cellValue,
            'data-formula': cellValue
          };
        }
        return { value: cellValue ?? '' };
      })
    );
    
    // 直接計算を実行
    calculateFormulas(initialData);
    setSelectedCell(null);
  };

  const calculateFormulas = (data: Matrix<CellBase>) => {
    // 直接データ形式で処理
    try {
      const processedData: SpreadsheetData = data.map(row => 
        row.map(cell => {
          if (typeof cell === 'object' && cell !== null && 'value' in cell) {
            const cellWithFormula = cell as { value?: string | number | null; formula?: string; 'data-formula'?: string };
            return {
              value: cellWithFormula.value ?? null,
              formula: cellWithFormula.formula,
              'data-formula': cellWithFormula['data-formula'],
              className: undefined,
              title: cellWithFormula.formula ? `数式: ${cellWithFormula.formula}` : undefined,
              DataEditor: undefined
            };
          }
          return { value: cell ?? '' };
        })
      );
      
      // mock functionを作成（recalculateFormulasが期待する形式）
      const mockFunction: ExcelFunctionResponse = {
        spreadsheet_data: processedData.map(row => 
          row.map(cell => ({
            v: cell ? (cell.value ?? null) : null,
            f: cell ? cell.formula : undefined,
            bg: undefined
          }))
        ),
        function_name: 'DEMO'
      };
      
      // 数式を計算
      recalculateFormulas(
        processedData,
        mockFunction,
        (name: string, value: SpreadsheetData) => {
          if (name === 'spreadsheetData') {
            console.log('計算結果を設定:', value);
            setSpreadsheetData(value as Matrix<CellBase>);
            setIsCalculating(false);
          }
        }
      );
    } catch (error) {
      console.error('スプレッドシートデータ処理エラー:', error);
      setIsCalculating(false);
    }
  };

  const handleCellClick = (row: number, col: number) => {
    const cell = spreadsheetData[row]?.[col];
    console.log('Cell clicked:', { row, col, cell }); // デバッグ用
    
    if (cell && typeof cell === 'object') {
      const cellData = cell as { formula?: string; value?: string | number | null };
      setSelectedCell({ 
        row, 
        col, 
        formula: cellData.formula,
        value: cellData.value
      });
    } else {
      setSelectedCell({ row, col, value: cell !== undefined ? cell as string | number | null : null });
    }
  };


  const getColumnLabel = (col: number): string => {
    return String.fromCharCode(65 + col);
  };

  const getCellAddress = (row: number, col: number): string => {
    return `${getColumnLabel(col)}${row + 1}`;
  };

  return (
    <div className="min-h-screen bg-gray-50">
      {/* ヘッダー */}
      <div className="bg-white shadow-sm border-b">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <h1 className="text-xl font-semibold text-gray-900">
              Excel関数デモモード
            </h1>
            <Link 
              to="/" 
              className="text-sm text-blue-600 hover:text-blue-800"
            >
              ← ChatGPTモードに戻る
            </Link>
          </div>
        </div>
      </div>

      <div className="max-w-7xl mx-auto p-6">
        {/* デモモード選択 */}
        <div className="mb-6">
          <div className="flex gap-4 mb-4">
            <button
              onClick={() => setDemoMode('grouped')}
              className={`px-4 py-2 rounded-lg font-medium ${
                demoMode === 'grouped'
                  ? 'bg-blue-500 text-white'
                  : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
              }`}
            >
              グループデモ
            </button>
            <button
              onClick={() => setDemoMode('individual')}
              className={`px-4 py-2 rounded-lg font-medium ${
                demoMode === 'individual'
                  ? 'bg-blue-500 text-white'
                  : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
              }`}
            >
              個別関数デモ
            </button>
          </div>
          
          {demoMode === 'grouped' ? (
            <>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                デモカテゴリを選択
              </label>
              <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-2">
                {demoSpreadsheetData.map(category => (
                  <button
                    key={category.id}
                    onClick={() => setSelectedCategory(category)}
                    className={`p-3 rounded-lg border-2 transition-all ${
                      selectedCategory.id === category.id
                        ? 'border-blue-500 bg-blue-50'
                        : 'border-gray-200 hover:border-gray-300'
                    }`}
                  >
                    <div className={`text-sm font-medium text-${category.color}-600`}>
                      {category.name}
                    </div>
                  </button>
                ))}
              </div>
            </>
          ) : (
            <>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                関数を選択
              </label>
              <select
                className="w-full p-2 border border-gray-300 rounded-lg"
                value={selectedFunction?.name || ''}
                onChange={(e) => {
                  const func = allIndividualFunctionTests.find(f => f.name === e.target.value);
                  setSelectedFunction(func || null);
                }}
              >
                <option value="">関数を選択してください</option>
                {allIndividualFunctionTests.map(func => (
                  <option key={func.name} value={func.name}>
                    {func.name} - {func.description}
                  </option>
                ))}
              </select>
            </>
          )}
        </div>

        {/* カテゴリ/関数説明 */}
        <div className="mb-4 p-4 bg-blue-50 rounded-lg">
          {demoMode === 'grouped' ? (
            <>
              <h3 className="font-medium text-blue-900 mb-1">{selectedCategory.name}</h3>
              <p className="text-sm text-blue-700">{selectedCategory.description}</p>
            </>
          ) : selectedFunction ? (
            <>
              <h3 className="font-medium text-blue-900 mb-1">{selectedFunction.name}</h3>
              <p className="text-sm text-blue-700">{selectedFunction.description}</p>
              {selectedFunction.expectedValues && (
                <div className="mt-2 text-xs text-blue-600">
                  期待値: {Object.entries(selectedFunction.expectedValues).map(([cell, value]) => (
                    <span key={cell} className="mr-2">
                      {cell}={String(value)}
                    </span>
                  ))}
                </div>
              )}
            </>
          ) : (
            <p className="text-sm text-blue-700">関数を選択してください</p>
          )}
        </div>

        {/* 数式バー */}
        <div className="mb-4 bg-white border rounded-lg shadow-sm">
          <div className="flex items-center border-b">
            <div className="px-4 py-2 border-r bg-gray-50 min-w-[80px]">
              <span className="text-sm font-medium text-gray-700">
                {selectedCell ? getCellAddress(selectedCell.row, selectedCell.col) : 'A1'}
              </span>
            </div>
            <div className="flex-1 px-4 py-2">
              <input
                type="text"
                readOnly
                className="w-full bg-transparent outline-none text-sm font-mono"
                value={
                  selectedCell?.formula ?? 
                  (selectedCell?.value !== null && selectedCell?.value !== undefined ? String(selectedCell.value) : '')
                }
                placeholder="セルを選択してください"
              />
            </div>
          </div>
        </div>

        {/* スプレッドシート */}
        <div className="mb-6 border rounded-lg overflow-hidden bg-white relative">
          {isCalculating ? (
            <div className="absolute inset-0 flex items-center justify-center bg-white bg-opacity-90 z-10">
              <div className="text-center">
                <div className="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-blue-500 mb-2"></div>
                <p className="text-sm text-gray-600">数式を計算中...</p>
              </div>
            </div>
          ) : null}
          <Spreadsheet
            data={spreadsheetData}
            onChange={setSpreadsheetData}
            onSelect={(selection: Selection) => {
              console.log('onSelect called:', selection); // デバッグ用
              
              // selection が null または undefined の場合は何もしない
              if (!selection) return;
              
              // RangeSelectionの場合（ChatGPTSpreadsheetと同じ実装）
              if (selection && 'range' in selection && selection.range) {
                const range = selection.range as { start: { row: number; column: number }; end: { row: number; column: number } };
                const row = range.start.row;
                const column = range.start.column;
                handleCellClick(row, column);
              }
            }}
            darkMode={false}
          />
        </div>

      </div>
    </div>
  );
}

export default DemoSpreadsheet;