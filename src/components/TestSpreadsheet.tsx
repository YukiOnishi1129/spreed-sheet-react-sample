import { useState, useEffect } from 'react';
import Spreadsheet, { type Matrix, type CellBase } from 'react-spreadsheet';
import { allIndividualFunctionTests as individualFunctionTests } from '../data/individualFunctionTests';
import type { IndividualFunctionTest } from '../types/spreadsheet';
import { calculateSingleFormula as calculateFormula } from './utils/customFormulaCalculations';
import type { CellData } from "../utils/spreadsheet/logic";

function TestSpreadsheet() {
  const [selectedCategory, setSelectedCategory] = useState<string>('');
  const [selectedFunction, setSelectedFunction] = useState<IndividualFunctionTest | null>(null);
  const [spreadsheetData, setSpreadsheetData] = useState<Matrix<CellBase>>([]);
  const [isCalculating, setIsCalculating] = useState<boolean>(false);
  const [testResults, setTestResults] = useState<{ [key: string]: { expected: unknown, actual: unknown, passed: boolean } }>({});

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
                  onChange={() => {}} // 読み取り専用
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