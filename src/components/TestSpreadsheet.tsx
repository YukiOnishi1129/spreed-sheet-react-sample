import { useState, useEffect } from 'react';
import Spreadsheet, { type Matrix, type CellBase, type Point, type Selection } from 'react-spreadsheet';
import { allIndividualFunctionTests as individualFunctionTests } from '../data/individualFunctionTests';
import type { IndividualFunctionTest } from '../types/spreadsheet';
import { calculateSingleFormula as calculateFormula } from './utils/customFormulaCalculations';
import { isRangeSelection } from './utils/typeGuards';
import type { CellData } from "../utils/spreadsheet/logic";

function TestSpreadsheet() {
  const [selectedCategory, setSelectedCategory] = useState<string>('');
  const [selectedFunction, setSelectedFunction] = useState<IndividualFunctionTest | null>(null);
  const [spreadsheetData, setSpreadsheetData] = useState<Matrix<CellBase>>([]);
  const [isCalculating, setIsCalculating] = useState<boolean>(false);
  const [testResults, setTestResults] = useState<{ [key: string]: { expected: unknown, actual: unknown, passed: boolean } }>({});
  const [selectedCell, setSelectedCell] = useState<Point | null>(null);
  const [selectedCellInfo, setSelectedCellInfo] = useState<{ value: string | number | boolean | null; formula?: string } | null>(null);
  const [editingFormula, setEditingFormula] = useState<string>('');
  const [copySuccess, setCopySuccess] = useState(false);

  // カテゴリ一覧を取得
  const categories = Array.from(new Set(individualFunctionTests.map(f => f.category))).sort();
  
  // 選択されたカテゴリの関数一覧を取得
  const functionsInCategory = selectedCategory 
    ? individualFunctionTests.filter(f => f.category === selectedCategory)
    : [];

  // 関数選択時の処理
  useEffect(() => {
    if (selectedFunction) {
      initializeFunction();
    }
  }, [selectedFunction]); // eslint-disable-line react-hooks/exhaustive-deps

  const initializeFunction = () => {
    if (!selectedFunction) return;
    
    setIsCalculating(true);
    setTestResults({});
    
    const initialData: Matrix<CellBase> = selectedFunction.data.map((row: (string | number | boolean | null)[]) => 
      row.map((cellValue: string | number | boolean | null) => {
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
    
    // 数式を計算
    calculateFormulas(initialData);
  };

  // 全ての数式を計算する関数
  const calculateAllFormulas = (data: CellData[][]): CellData[][] => {
    const result: CellData[][] = data.map(row => row.map(cell => ({ ...cell })));
    const maxIterations = 10;
    
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
    try {
      setIsCalculating(true);
      
      // CellData形式に変換
      const cellData: CellData[][] = data.map((row) =>
        row.map((cell) => {
          if (!cell) return { value: '' };
          
          if (typeof cell === 'object' && 'value' in cell) {
            const cellObj = cell as CellBase & { formula?: string; 'data-formula'?: string };
            const value = String(cellObj.value ?? '');
            const formula = cellObj.formula ?? cellObj['data-formula'];
            
            return {
              value: value as string | number | boolean | null,
              formula: formula
            };
          }
          
          const cellValue = cell as string | number | boolean | null;
          
          if (typeof cellValue === 'string' && cellValue?.startsWith('=')) {
            return {
              value: '',
              formula: cellValue
            };
          }
          
          return { value: cellValue };
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
      
      setSpreadsheetData(resultData);
      
      // テスト結果を検証
      if (selectedFunction?.expectedValues) {
        const results: { [key: string]: { expected: unknown, actual: unknown, passed: boolean } } = {};
        
        Object.entries(selectedFunction.expectedValues).forEach(([cell, expectedValue]) => {
          const col = cell.charCodeAt(0) - 65;
          const row = parseInt(cell.substring(1)) - 1;
          const actualValue = resultData[row]?.[col]?.value as string | number | boolean | null | undefined;
          
          let passed = false;
          if (typeof expectedValue === 'number' && typeof actualValue === 'number') {
            // 数値の場合は近似値で比較
            passed = Math.abs(actualValue - expectedValue) < 0.01;
          } else {
            passed = actualValue === expectedValue;
          }
          
          results[cell] = {
            expected: expectedValue,
            actual: actualValue,
            passed
          };
        });
        
        setTestResults(results);
      }
      
      setIsCalculating(false);
    } catch (error) {
      console.error('スプレッドシートデータ処理エラー:', error);
      setIsCalculating(false);
    }
  };

  // スプレッドシートのデータが変更されたときの処理
  const handleSpreadsheetChange = (newData: Matrix<CellBase>) => {
    setSpreadsheetData(newData);
    
    // 再計算をトリガー
    setTimeout(() => {
      calculateFormulas(newData);
    }, 100);
  };

  // セルが選択されたときの処理
  const handleCellSelect = (selection: Selection) => {
    // selection が null または undefined の場合は何もしない
    if (!selection) return;
    
    // RangeSelectionの場合（ChatGPTSpreadsheetと同じ実装）
    if (isRangeSelection(selection)) {
      const row = selection.range.start.row;
      const column = selection.range.start.column;
      setSelectedCell({ row, column });
    }
  };

  // 選択されたセルの情報を更新
  useEffect(() => {
    console.log('Selected cell:', selectedCell);
    console.log('Spreadsheet data:', spreadsheetData);
    
    if (selectedCell && spreadsheetData[selectedCell.row]) {
      const cell = spreadsheetData[selectedCell.row][selectedCell.column];
      console.log('Cell data:', cell);
      
      if (cell) {
        const cellWithFormula = cell as CellBase & { formula?: string; 'data-formula'?: string };
        const formula = cellWithFormula.formula ?? cellWithFormula['data-formula'];
        
        console.log('Cell value:', cell.value);
        console.log('Cell formula:', formula);
        
        setSelectedCellInfo({
          value: cell.value as string | number | boolean | null,
          formula: formula ?? undefined
        });
        
        // 編集フィールドには数式がある場合は数式を、ない場合は値を表示
        if (formula) {
          setEditingFormula(formula);
        } else {
          setEditingFormula(String(cell.value ?? ''));
        }
      } else {
        setSelectedCellInfo({ value: null });
        setEditingFormula('');
      }
    }
  }, [selectedCell, spreadsheetData]);

  // 数式エディタの値が変更されたときの処理
  const handleFormulaChange = (newFormula: string) => {
    setEditingFormula(newFormula);
    
    if (selectedCell) {
      const newData = [...spreadsheetData];
      const cell = newData[selectedCell.row][selectedCell.column] ?? {};
      
      if (newFormula.startsWith('=')) {
        newData[selectedCell.row][selectedCell.column] = {
          ...cell,
          value: '',
          formula: newFormula
        };
      } else {
        newData[selectedCell.row][selectedCell.column] = {
          ...cell,
          value: newFormula,
          formula: undefined
        };
      }
      
      handleSpreadsheetChange(newData);
    }
  };

  // Excel用コピー機能
  const handleExcelCopy = async () => {
    try {
      const excelData: string[][] = [];
      
      for (let row = 0; row < spreadsheetData.length; row++) {
        const rowData: string[] = [];
        for (let col = 0; col < spreadsheetData[row].length; col++) {
          const cell = spreadsheetData[row][col];
          if (cell) {
            const cellWithFormula = cell as CellBase & { formula?: string; 'data-formula'?: string };
            if (cellWithFormula.formula || cellWithFormula['data-formula']) {
              rowData.push(cellWithFormula.formula ?? cellWithFormula['data-formula'] ?? '');
            } else {
              rowData.push(String(cell.value ?? ''));
            }
          } else {
            rowData.push('');
          }
        }
        excelData.push(rowData);
      }
      
      // タブ区切りテキストとして整形
      const tsvData = excelData.map(row => row.join('\t')).join('\n');
      
      await navigator.clipboard.writeText(tsvData);
      setCopySuccess(true);
      setTimeout(() => setCopySuccess(false), 2000);
    } catch (error) {
      console.error('コピーに失敗しました:', error);
    }
  };


  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto">
        <h1 className="text-2xl font-bold text-gray-900 mb-6">
          Excel関数テストモード
        </h1>

        {/* カテゴリと関数の選択UI */}
        <div className="bg-white rounded-lg shadow-sm p-6 mb-6">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* カテゴリ選択 */}
            <div>
              <label htmlFor="category-select" className="block text-sm font-medium text-gray-700 mb-2">
                カテゴリを選択
              </label>
              <select
                id="category-select"
                value={selectedCategory}
                onChange={(e) => {
                  setSelectedCategory(e.target.value);
                  setSelectedFunction(null);
                }}
                className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
              >
                <option value="">-- カテゴリを選択 --</option>
                {categories.map(category => (
                  <option key={category} value={category}>
                    {category}
                  </option>
                ))}
              </select>
            </div>

            {/* 関数選択 */}
            <div>
              <label htmlFor="function-select" className="block text-sm font-medium text-gray-700 mb-2">
                関数を選択
              </label>
              <select
                id="function-select"
                value={selectedFunction?.name ?? ''}
                onChange={(e) => {
                  const func = functionsInCategory.find(f => f.name === e.target.value);
                  setSelectedFunction(func ?? null);
                }}
                disabled={!selectedCategory}
                className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 disabled:bg-gray-100"
              >
                <option value="">-- 関数を選択 --</option>
                {functionsInCategory.map(func => (
                  <option key={func.name} value={func.name}>
                    {func.name} - {func.description}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* 選択された関数の情報 */}
          {selectedFunction && (
            <div className="mt-6 p-4 bg-blue-50 rounded-lg">
              <h3 className="font-medium text-blue-900 mb-1">{selectedFunction.name}</h3>
              <p className="text-sm text-blue-700">{selectedFunction.description}</p>
            </div>
          )}
        </div>

        {/* テスト結果 */}
        {Object.keys(testResults).length > 0 && (
          <div className="bg-white rounded-lg shadow-sm p-6 mb-6">
            <h3 className="text-lg font-medium text-gray-900 mb-4">テスト結果</h3>
            <div className="space-y-2">
              {Object.entries(testResults).map(([cell, result]) => (
                <div 
                  key={cell}
                  className={`flex items-center justify-between p-3 rounded-lg ${
                    result.passed ? 'bg-green-50' : 'bg-red-50'
                  }`}
                >
                  <div className="flex items-center space-x-4">
                    <span className="font-mono font-medium">{cell}</span>
                    <span className="text-sm text-gray-600">
                      期待値: {JSON.stringify(result.expected)}
                    </span>
                    <span className="text-sm text-gray-600">
                      実際値: {JSON.stringify(result.actual)}
                    </span>
                  </div>
                  <div className={`px-3 py-1 rounded-full text-xs font-medium ${
                    result.passed 
                      ? 'bg-green-100 text-green-800' 
                      : 'bg-red-100 text-red-800'
                  }`}>
                    {result.passed ? '成功' : '失敗'}
                  </div>
                </div>
              ))}
            </div>
            <div className="mt-4 text-right">
              <span className="text-sm text-gray-600">
                合計: {Object.keys(testResults).length} / 
                成功: {Object.values(testResults).filter(r => r.passed).length} / 
                失敗: {Object.values(testResults).filter(r => !r.passed).length}
              </span>
            </div>
          </div>
        )}

        {/* セル編集バー */}
        {selectedFunction && (
          <div className="bg-white rounded-lg shadow-sm p-4 mb-6">
            <div className="space-y-4">
              {/* ChatGPTSpreadsheet.tsxと同じスタイルのセル情報表示 */}
              {console.log('selectedCellInfo:', selectedCellInfo)}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <p className="text-sm font-semibold text-gray-600 mb-2">セルアドレス:</p>
                  <div className="bg-gray-50 px-3 py-2 rounded-lg border border-gray-200 text-sm font-mono">
                    {selectedCell && selectedCell.row !== undefined && selectedCell.column !== undefined 
                      ? `${String.fromCharCode(65 + selectedCell.column)}${selectedCell.row + 1}` 
                      : 'A1'}
                  </div>
                </div>
                
                <div>
                  <p className="text-sm font-semibold text-gray-600 mb-2">数式/値:</p>
                  <div className="bg-gray-50 px-3 py-2 rounded-lg border border-gray-200 text-sm font-mono">
                    <input
                      type="text"
                      value={
                        selectedCellInfo?.formula ?? 
                        (selectedCellInfo?.value !== null && selectedCellInfo?.value !== undefined 
                          ? String(selectedCellInfo.value) 
                          : '')
                      }
                      readOnly
                      className="w-full bg-transparent border-none outline-none text-gray-800"
                      placeholder="数式や値が表示されます"
                    />
                  </div>
                </div>
              </div>
              
              {/* 編集用入力フィールド */}
              <div>
                <p className="text-sm font-semibold text-gray-600 mb-2">編集:</p>
                <input
                  type="text"
                  value={editingFormula}
                  onChange={(e) => handleFormulaChange(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      e.preventDefault();
                      // Enterキーで確定
                    }
                  }}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 font-mono text-sm"
                  placeholder="値または数式を入力"
                />
              </div>
              
              {/* Excel用コピーボタン */}
              <div className="flex justify-between items-center">
                <div className="text-sm text-gray-500">
                  {selectedCellInfo?.formula && (
                    <span className="flex items-center space-x-1">
                      <span className="text-green-600">fx</span>
                      <span>数式モード</span>
                    </span>
                  )}
                </div>
                <button
                  onClick={handleExcelCopy}
                  className="flex items-center space-x-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors duration-200"
                >
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 5H6a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2v-1M8 5a2 2 0 002 2h2a2 2 0 002-2M8 5a2 2 0 012-2h2a2 2 0 012 2m0 0h2a2 2 0 012 2v3m2 4H10m0 0l3-3m-3 3l3 3" />
                  </svg>
                  <span>📋 Excel用コピー</span>
                </button>
              </div>
              
              {/* コピー成功メッセージ */}
              {copySuccess && (
                <div className="absolute top-4 right-4 bg-green-600 text-white px-4 py-2 rounded-lg shadow-lg flex items-center space-x-2 animate-fade-in-out">
                  <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                  </svg>
                  <span>コピーしました！</span>
                </div>
              )}
            </div>
          </div>
        )}

        {/* スプレッドシート */}
        {selectedFunction && (
          <div className="bg-white rounded-lg shadow-sm p-6">
            <div className="relative">
              {isCalculating && (
                <div className="absolute inset-0 flex items-center justify-center bg-white bg-opacity-90 z-10">
                  <div className="text-center">
                    <div className="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-blue-500 mb-2"></div>
                    <p className="text-sm text-gray-600">数式を計算中...</p>
                  </div>
                </div>
              )}
              <div className="border rounded-lg overflow-hidden">
                <Spreadsheet
                  data={spreadsheetData}
                  onChange={handleSpreadsheetChange}
                  onSelect={handleCellSelect}
                  darkMode={false}
                />
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default TestSpreadsheet;