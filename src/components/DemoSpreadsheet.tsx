import { useState, useEffect } from 'react';
import Spreadsheet, { type Matrix, type CellBase, type Selection } from 'react-spreadsheet';
import { demoSpreadsheetData, type DemoCategory } from '../data/demoSpreadsheetData';
import type { IndividualFunctionTest } from '../types/spreadsheet';
import { Link } from 'react-router-dom';
import { calculateSingleFormula as calculateFormula } from './utils/customFormulaCalculations';
import { FunctionSelectorModal } from './FunctionSelectorModal';
import type { CellData } from "../utils/spreadsheet/logic";

function DemoSpreadsheet() {
  const [selectedCategory, setSelectedCategory] = useState<DemoCategory>(demoSpreadsheetData[0]);
  const [spreadsheetData, setSpreadsheetData] = useState<Matrix<CellBase>>([]);
  const [selectedCell, setSelectedCell] = useState<{ row: number; col: number; formula?: string; value?: string | number | null } | null>(null);
  const [demoMode, setDemoMode] = useState<'grouped' | 'individual'>('grouped');
  const [selectedFunction, setSelectedFunction] = useState<IndividualFunctionTest | null>(null);
  const [isCalculating, setIsCalculating] = useState<boolean>(false);
  const [isFunctionModalOpen, setIsFunctionModalOpen] = useState<boolean>(false);
  const [originalFormulas, setOriginalFormulas] = useState<Matrix<string | undefined>>([]);


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
    
    // 数式を別途保存
    const formulas: Matrix<string | undefined> = selectedCategory.data.map((row: (string | number | boolean | null)[]) => 
      row.map((cellValue: string | number | boolean | null) => {
        if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
          return cellValue;
        }
        return undefined;
      })
    );
    setOriginalFormulas(formulas);
    
    const initialData: Matrix<CellBase> = selectedCategory.data.map((row: (string | number | boolean | null)[]) => 
      row.map((cellValue: string | number | boolean | null) => {
        // 数式が直接入っている場合
        if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
          return {
            value: '',
            formula: cellValue,
            'data-formula': cellValue
          };
        }
        // 空文字列の場合はそのまま保持（COUNTBLANKが正しく動作するように）
        return { value: cellValue === null || cellValue === undefined ? '' : cellValue };
      })
    );
    
    // 直接計算を実行
    calculateFormulas(initialData);
    setSelectedCell(null);
  };
  
  const initializeIndividualFunction = () => {
    if (!selectedFunction) return;
    
    setIsCalculating(true);
    
    // 数式を別途保存
    const formulas: Matrix<string | undefined> = selectedFunction.data.map((row: (string | number | boolean | null)[]) => 
      row.map((cellValue: string | number | boolean | null) => {
        if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
          return cellValue;
        }
        return undefined;
      })
    );
    setOriginalFormulas(formulas);
    
    const initialData: Matrix<CellBase> = selectedFunction.data.map((row: (string | number | boolean | null)[], rowIdx) => 
      row.map((cellValue: string | number | boolean | null, colIdx) => {
        if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
          // MROUNDのデバッグ
          if (selectedFunction.name === 'MROUND') {
            console.log(`Creating formula cell at [${rowIdx}][${colIdx}]:`, cellValue);
          }
          return {
            value: '',
            formula: cellValue,
            'data-formula': cellValue
          };
        }
        // 空文字列の場合はそのまま保持（COUNTBLANKが正しく動作するように）
        return { value: cellValue === null || cellValue === undefined ? '' : cellValue };
      })
    );
    
    
    // 直接計算を実行
    calculateFormulas(initialData);
    setSelectedCell(null);
  };


  // 全ての数式を計算する関数
  const calculateAllFormulas = (data: CellData[][]): CellData[][] => {
    const result: CellData[][] = data.map(row => row.map(cell => ({ ...cell })));
    const maxIterations = 10; // 循環参照を防ぐための最大反復回数
    
    for (let iteration = 0; iteration < maxIterations; iteration++) {
      let hasChanges = false;
      
      for (let rowIndex = 0; rowIndex < result.length; rowIndex++) {
        for (let colIndex = 0; colIndex < result[rowIndex].length; colIndex++) {
          const cell = result[rowIndex][colIndex];
          
          if (cell.formula) {
            const newValue = calculateFormula(cell.formula, result, rowIndex, colIndex);
            
            
            if (newValue !== cell.value) {
              cell.value = newValue;
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
  };

  const calculateFormulas = (data: Matrix<CellBase>) => {
    // 直接データ形式で処理
    try {
      setIsCalculating(true);
      
      // CellData形式に変換（数式計算用）
      const cellData: CellData[][] = data.map((row, rowIdx) =>
        row.map((cell, colIdx) => {
          if (!cell) return { value: '' };
          
          if (typeof cell === 'object' && 'value' in cell) {
            const cellObj = cell as any;
            let value = cellObj.value ?? '';
            
            // 数値文字列を数値に変換
            if (typeof value === 'string' && value !== '') {
              const num = parseFloat(value);
              if (!isNaN(num) && num.toString() === (value as string).trim()) {
                value = num;
              }
            }
            
            const formula = cellObj.formula || cellObj['data-formula'];
            
            // MROUND関数のデバッグ
            if (selectedFunction?.name === 'MROUND' && formula && formula.includes('MROUND')) {
              console.log(`Converting cell [${rowIdx}][${colIdx}] to CellData:`, {
                originalCell: cell,
                value,
                formula
              });
            }
            
            return {
              value: value,
              formula: formula
            };
          }
          
          // 単純な値の場合も数値変換を試みる
          if (typeof cell === 'string' && cell !== '') {
            const num = parseFloat(cell);
            if (!isNaN(num) && num.toString() === (cell as string).trim()) {
              return { value: num };
            }
          }
          
          return { value: cell };
        })
      );
      
      // 数式を計算
      const calculatedData = calculateAllFormulas(cellData);
      
      // 結果をSpreadsheet形式に変換
      const resultData: Matrix<CellBase> = calculatedData.map((row) =>
        row.map((cell) => {
          if (cell.formula) {
            
            return {
              value: cell.value,
              formula: cell.formula,
              'data-formula': cell.formula,
              title: `数式: ${cell.formula}`
            };
          }
          
          return {
            value: cell.value
          };
        })
      );
      
      
      // MROUND関数の場合、最終的なデータを確認
      if (selectedFunction?.name === 'MROUND') {
        console.log('Setting spreadsheet data for MROUND:');
        resultData.forEach((row, rowIdx) => {
          console.log(`Row ${rowIdx}:`, row.map(cell => 
            typeof cell === 'object' && cell !== null ? 
              `{value: ${cell.value}, formula: ${(cell as any).formula || 'none'}}` : 
              cell
          ));
        });
      }
      
      setSpreadsheetData(resultData);
      setIsCalculating(false);
    } catch (error) {
      console.error('スプレッドシートデータ処理エラー:', error);
      setIsCalculating(false);
    }
  };

  const handleCellClick = (row: number, col: number) => {
    const cell = spreadsheetData[row]?.[col];
    
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
            <div className="space-y-3">
              {/* 関数選択ボタン */}
              <button
                onClick={() => setIsFunctionModalOpen(true)}
                className="w-full p-4 border-2 border-dashed border-gray-300 rounded-lg hover:border-blue-400 hover:bg-blue-50 transition-colors flex items-center justify-center space-x-2"
              >
                <svg className="w-5 h-5 text-gray-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
                </svg>
                <span className="text-gray-700 font-medium">
                  {selectedFunction ? `選択中: ${selectedFunction.name}` : '関数を選択'}
                </span>
              </button>
              
              {/* 選択された関数の情報 */}
              {selectedFunction && (
                <div className="p-4 bg-gray-50 rounded-lg">
                  <div className="flex items-start justify-between">
                    <div className="flex-1">
                      <h4 className="font-medium text-gray-900">{selectedFunction.name}</h4>
                      <p className="text-sm text-gray-600 mt-1">{selectedFunction.description}</p>
                    </div>
                    <button
                      onClick={() => setSelectedFunction(null)}
                      className="ml-4 p-1 hover:bg-gray-200 rounded transition-colors"
                    >
                      <svg className="w-4 h-4 text-gray-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                      </svg>
                    </button>
                  </div>
                </div>
              )}
            </div>
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
            onChange={(newData) => {
              // 数式情報を保持しながらデータを更新
              const updatedData = newData.map((row, rowIndex) =>
                row.map((cell, colIndex) => {
                  const formula = originalFormulas[rowIndex]?.[colIndex];
                  
                  // 数式があるセルの場合
                  if (formula) {
                    return {
                      value: cell?.value ?? '',
                      formula: formula,
                      'data-formula': formula
                    };
                  }
                  return cell;
                })
              );
              
              setSpreadsheetData(updatedData);
              // セル値が変更されたら数式を再計算
              calculateFormulas(updatedData);
            }}
            onSelect={(selection: Selection) => {
              
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

      {/* 関数選択モーダル */}
      <FunctionSelectorModal
        isOpen={isFunctionModalOpen}
        onClose={() => setIsFunctionModalOpen(false)}
        onSelect={(func) => {
          setSelectedFunction(func);
          setIsFunctionModalOpen(false);
        }}
        selectedFunction={selectedFunction}
      />
    </div>
  );
}

export default DemoSpreadsheet;