import { useState, useEffect } from 'react';
import Spreadsheet, { type Matrix, type CellBase } from 'react-spreadsheet';
import { testCategories, type TestCategory } from '../data/testData';
import { Link } from 'react-router-dom';
import { recalculateFormulas } from './utils/customFormulaCalculations';
import type { SpreadsheetData, ExcelFunctionResponse } from '../types/spreadsheet';

interface TestResult {
  row: number;
  col: number;
  expected: string | number;
  actual: string | number | null;
  passed: boolean;
}

function TestSpreadsheet() {
  const [selectedCategory, setSelectedCategory] = useState<TestCategory>(testCategories[0]);
  const [spreadsheetData, setSpreadsheetData] = useState<Matrix<CellBase>>([]);
  const [testResults, setTestResults] = useState<TestResult[]>([]);
  const [showResults, setShowResults] = useState(false);
  const [selectedCell, setSelectedCell] = useState<{ row: number; col: number; formula?: string } | null>(null);

  // カテゴリ選択時にデータを初期化と自動計算
  useEffect(() => {
    initializeAndCalculate();
  }, [selectedCategory]); // eslint-disable-line react-hooks/exhaustive-deps

  const initializeAndCalculate = () => {
    const initialData: Matrix<CellBase> = selectedCategory.data.map((row, rowIndex) => 
      row.map((cellValue, colIndex) => {
        // 数式がある場所を探す
        const formula = selectedCategory.formulas.find(
          f => f.row === rowIndex && f.col === colIndex
        );
        
        if (formula) {
          const cell = {
            value: '',
            formula: formula.formula,
            'data-formula': formula.formula
          };
          return cell;
        }
        
        const cell = { value: cellValue ?? '' };
        return cell;
      })
    );
    
    // 自動的に計算を実行
    calculateFormulas(initialData);
    setSelectedCell(null);
  };

  const calculateFormulas = (data: Matrix<CellBase>) => {
    // カスタム関数システムで計算
    const mockFunction: ExcelFunctionResponse = {
      spreadsheet_data: data.map(row => 
        row.map(cell => {
          if (typeof cell === 'object' && cell !== null && 'value' in cell) {
            const cellWithFormula = cell as { value?: string | number | null; formula?: string };
            return {
              v: cellWithFormula.value ?? null,
              f: cellWithFormula.formula
            };
          }
          return { v: String(cell) };
        })
      ),
      function_name: 'TEST'
    };
    
    recalculateFormulas(
      data,
      mockFunction,
      (name: string, value: SpreadsheetData) => {
        if (name === 'spreadsheetData') {
          setSpreadsheetData(value as Matrix<CellBase>);
          
          // 自動的にテスト結果を計算
          if (selectedCategory.expectedResults) {
            const results: TestResult[] = [];
            
            selectedCategory.expectedResults.forEach(expected => {
              const actualCell = value[expected.row]?.[expected.col];
              const actualValue = actualCell?.value ?? null;
              
              results.push({
                row: expected.row,
                col: expected.col,
                expected: expected.value,
                actual: actualValue,
                passed: compareValues(expected.value, actualValue)
              });
            });
            
            setTestResults(results);
            setShowResults(true);
          }
        }
      }
    );
  };

  const handleCellClick = (row: number, col: number) => {
    const cell = spreadsheetData[row]?.[col];
    if (cell && typeof cell === 'object' && 'formula' in cell) {
      const cellWithFormula = cell as { formula?: string };
      setSelectedCell({ row, col, formula: cellWithFormula.formula });
    } else {
      setSelectedCell({ row, col });
    }
  };

  const compareValues = (expected: string | number, actual: string | number | null): boolean => {
    if (actual === null) return false;
    
    // 数値の場合は小数点以下2桁で比較
    if (typeof expected === 'number' && typeof actual === 'number') {
      return Math.abs(expected - actual) < 0.01;
    }
    
    // 文字列の場合は完全一致
    return String(expected) === String(actual);
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
              Excel関数テストモード
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
        {/* カテゴリ選択 */}
        <div className="mb-6">
          <label className="block text-sm font-medium text-gray-700 mb-2">
            テストカテゴリを選択
          </label>
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-2">
            {testCategories.map(category => (
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
        </div>

        {/* カテゴリ説明 */}
        <div className="mb-4 p-4 bg-blue-50 rounded-lg">
          <h3 className="font-medium text-blue-900 mb-1">{selectedCategory.name}</h3>
          <p className="text-sm text-blue-700">{selectedCategory.description}</p>
        </div>

        {/* 選択中のセルの数式表示 */}
        {selectedCell?.formula && (
          <div className="mb-4 p-4 bg-gray-50 rounded-lg">
            <div className="text-sm text-gray-600">
              セル {getCellAddress(selectedCell.row, selectedCell.col)} の数式:
            </div>
            <div className="mt-1 font-mono text-lg text-gray-900">
              {selectedCell.formula}
            </div>
          </div>
        )}

        {/* スプレッドシート */}
        <div className="mb-6 border rounded-lg overflow-hidden bg-white">
          <Spreadsheet
            data={spreadsheetData}
            onChange={setSpreadsheetData}
            onCellCommit={(_prevCell, _nextCell, coords) => {
              if (coords) {
                handleCellClick(coords.row, coords.column);
              }
            }}
            darkMode={false}
          />
        </div>

        {/* テスト結果 */}
        {showResults && testResults.length > 0 && (
          <div className="bg-white rounded-lg shadow p-6">
            <h3 className="text-lg font-semibold mb-4">テスト結果</h3>
            
            {/* サマリー */}
            <div className="mb-4 p-4 bg-gray-50 rounded">
              <div className="text-sm">
                <span className="font-medium">合計: </span>
                {testResults.length} テスト
              </div>
              <div className="text-sm">
                <span className="font-medium text-green-600">成功: </span>
                {testResults.filter(r => r.passed).length}
              </div>
              <div className="text-sm">
                <span className="font-medium text-red-600">失敗: </span>
                {testResults.filter(r => !r.passed).length}
              </div>
            </div>

            {/* 詳細結果 */}
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50">
                  <tr>
                    <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">
                      セル
                    </th>
                    <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">
                      期待値
                    </th>
                    <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">
                      実際の値
                    </th>
                    <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">
                      結果
                    </th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {testResults.map((result, index) => (
                    <tr key={index} className={result.passed ? '' : 'bg-red-50'}>
                      <td className="px-4 py-2 text-sm font-mono">
                        {getCellAddress(result.row, result.col)}
                      </td>
                      <td className="px-4 py-2 text-sm">
                        {result.expected}
                      </td>
                      <td className="px-4 py-2 text-sm">
                        {result.actual ?? '(null)'}
                      </td>
                      <td className="px-4 py-2 text-sm">
                        {result.passed ? (
                          <span className="text-green-600 font-medium">✓ OK</span>
                        ) : (
                          <span className="text-red-600 font-medium">✗ NG</span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default TestSpreadsheet;