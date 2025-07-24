import { useState, useEffect } from 'react';
import Spreadsheet, { type Matrix, type CellBase } from 'react-spreadsheet';
import { testCategories, type TestCategory } from '../data/testData';
import { Link } from 'react-router-dom';
import { recalculateFormulas } from './utils/customFormulaCalculations';
import type { SpreadsheetData, ExcelFunctionResponse } from '../types/spreadsheet';

function TestSpreadsheet() {
  const [selectedCategory, setSelectedCategory] = useState<TestCategory>(testCategories[0]);
  const [spreadsheetData, setSpreadsheetData] = useState<Matrix<CellBase>>([]);
  const [selectedCell, setSelectedCell] = useState<{ row: number; col: number; formula?: string; value?: string | number | null } | null>(null);

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
        }
      }
    );
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
        <div className="mb-6 border rounded-lg overflow-hidden bg-white">
          <Spreadsheet
            data={spreadsheetData}
            onChange={setSpreadsheetData}
            onSelect={(selection) => {
              if (selection && 'row' in selection && 'column' in selection) {
                const sel = selection as { row: number; column: number };
                handleCellClick(sel.row, sel.column);
              }
            }}
            darkMode={false}
          />
        </div>

      </div>
    </div>
  );
}

export default TestSpreadsheet;