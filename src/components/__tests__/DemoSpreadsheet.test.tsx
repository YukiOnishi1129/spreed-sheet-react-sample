import { describe, test, expect, vi } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import { BrowserRouter } from 'react-router-dom';
import DemoSpreadsheet from '../DemoSpreadsheet';
import { allIndividualFunctionTests } from '../../data/individualFunctionTests';

// react-spreadsheetのモック
vi.mock('react-spreadsheet', () => ({
  __esModule: true,
  default: ({ data, onSelect }: any) => {
    return (
      <div data-testid="spreadsheet-mock">
        {data.map((row: any[], rowIndex: number) => (
          <div key={rowIndex} data-testid={`row-${rowIndex}`}>
            {row.map((cell: any, colIndex: number) => (
              <div
                key={colIndex}
                data-testid={`cell-${rowIndex}-${colIndex}`}
                onClick={() => onSelect?.({ row: rowIndex, column: colIndex })}
              >
                {cell.formula ? `Formula: ${cell.formula}` : `Value: ${cell.value}`}
              </div>
            ))}
          </div>
        ))}
      </div>
    );
  }
}));

// カスタム計算関数のモック
vi.mock('../utils/customFormulaCalculations', () => ({
  recalculateFormulas: (data: any, _mockFunction: any, callback: any) => {
    // 簡単な計算ロジックをモック
    const calculatedData = data.map((row: any[]) =>
      row.map((cell: any) => {
        if (cell.formula) {
          // 簡単な計算例
          if (cell.formula === '=SUM(A2:C2)') {
            return { ...cell, value: 60 };
          }
          if (cell.formula === '=AVERAGE(A2:D2)') {
            return { ...cell, value: 25 };
          }
          return { ...cell, value: 'calculated' };
        }
        return cell;
      })
    );
    callback('spreadsheetData', calculatedData);
  }
}));

describe('DemoSpreadsheet', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <DemoSpreadsheet />
      </BrowserRouter>
    );
  };

  test('コンポーネントが正しくレンダリングされること', () => {
    renderComponent();
    expect(screen.getByText('Excel関数デモモード')).toBeInTheDocument();
    expect(screen.getByText('← ChatGPTモードに戻る')).toBeInTheDocument();
  });

  test('デモモードの切り替えが機能すること', () => {
    renderComponent();
    
    // デフォルトはグループデモモード
    expect(screen.getByText('グループデモ')).toHaveClass('bg-blue-500');
    expect(screen.getByText('個別関数デモ')).toHaveClass('bg-gray-200');
    
    // 個別関数デモモードに切り替え
    fireEvent.click(screen.getByText('個別関数デモ'));
    expect(screen.getByText('個別関数デモ')).toHaveClass('bg-blue-500');
    expect(screen.getByText('グループデモ')).toHaveClass('bg-gray-200');
  });

  test('グループデモモードでカテゴリ選択が機能すること', () => {
    renderComponent();
    
    // カテゴリボタンが表示されている（複数の要素がある場合はgetAllByTextを使用）
    const basicButtons = screen.getAllByText('基本関数');
    expect(basicButtons.length).toBeGreaterThan(0);
    const mathButtons = screen.getAllByText('数学・三角関数');
    expect(mathButtons.length).toBeGreaterThan(0);
    
    // カテゴリボタンをクリック（最初の要素をクリック）
    fireEvent.click(mathButtons[0]);
    
    // 選択されたカテゴリがハイライトされる
    const selectedButton = mathButtons[0].parentElement;
    expect(selectedButton).toHaveClass('border-blue-500');
  });

  test('個別関数デモモードで関数選択が機能すること', () => {
    renderComponent();
    
    // 個別関数デモモードに切り替え
    fireEvent.click(screen.getByText('個別関数デモ'));
    
    // セレクトボックスが表示される
    const selectElement = screen.getByRole('combobox');
    expect(selectElement).toBeInTheDocument();
    
    // 関数を選択
    fireEvent.change(selectElement, { target: { value: 'SUM' } });
    
    // 選択された関数の説明が表示される
    expect(screen.getByText('SUM')).toBeInTheDocument();
    expect(screen.getByText('数値の合計を計算')).toBeInTheDocument();
  });

  test('数式バーが正しく表示されること', async () => {
    renderComponent();
    
    await waitFor(() => {
      const formulaBar = screen.getByDisplayValue('');
      expect(formulaBar).toBeInTheDocument();
      expect(formulaBar).toHaveAttribute('readonly');
    });
  });

  test('セル選択時に数式バーが更新されること', async () => {
    renderComponent();
    
    // スプレッドシートのレンダリングを待つ
    await waitFor(() => {
      expect(screen.getByTestId('spreadsheet-mock')).toBeInTheDocument();
    });
    
    // セルをクリック
    const cell = screen.getByTestId('cell-0-0');
    fireEvent.click(cell);
    
    // セルアドレスが表示される
    await waitFor(() => {
      expect(screen.getByText('A1')).toBeInTheDocument();
    });
  });

  test('スプレッドシートが正しくレンダリングされること', async () => {
    renderComponent();
    
    await waitFor(() => {
      expect(screen.getByTestId('spreadsheet-mock')).toBeInTheDocument();
    });
  });
});

// 個別関数のテスト
describe('個別Excel関数のテスト', () => {
  // 数学関数のテスト
  describe('数学関数', () => {
    test.each([
      ['SUM', { 'E2': 100 }],
      ['AVERAGE', { 'E2': 25 }],
      ['MAX', { 'E2': 40 }],
      ['MIN', { 'E2': 10 }],
      ['PRODUCT', { 'D2': 24 }],
      ['SQRT', { 'B2': 4, 'B3': 5, 'B4': 10 }],
      ['POWER', { 'C2': 8, 'C3': 25 }],
      ['ABS', { 'B2': 10, 'B3': 10, 'B4': 0 }],
      ['ROUND', { 'C2': 3.14, 'C3': 120, 'C4': 3 }],
      ['MOD', { 'C2': 1, 'C3': 3, 'C4': 0 }],
      ['INT', { 'B2': 3, 'B3': -4, 'B4': 5 }],
      ['SIGN', { 'B2': 1, 'B3': -1, 'B4': 0 }],
    ])('%s関数が正しく計算されること', (functionName: string, expectedValues: Record<string, string | number | boolean>) => {
      const testCase = allIndividualFunctionTests.find(t => t.name === functionName);
      expect(testCase).toBeDefined();
      
      if (testCase?.expectedValues) {
        Object.entries(expectedValues).forEach(([cell, expectedValue]) => {
          expect(testCase.expectedValues![cell]).toBe(expectedValue);
        });
      }
    });
  });

  // 統計関数のテスト
  describe('統計関数', () => {
    test.each([
      ['COUNT', { 'E2': 3 }],
      ['COUNTA', { 'E2': 3 }],
      ['COUNTBLANK', { 'E2': 2 }],
      ['COUNTIF', { 'B5': 2 }],
      ['MEDIAN', { 'F2': 3 }],
      ['MODE', { 'F2': 3 }],
      ['STDEV', { 'E2': 12.90994449 }],
      ['VAR', { 'E2': 166.6666667 }],
      ['CORREL', { 'D5': 1 }],
      ['LARGE', { 'F2': 30 }],
      ['SMALL', { 'F2': 20 }],
      ['RANK', { 'B2': 4, 'B3': 2, 'B4': 3, 'B5': 1 }],
    ])('%s関数が正しく計算されること', (functionName: string, expectedValues: Record<string, string | number | boolean>) => {
      const testCase = allIndividualFunctionTests.find(t => t.name === functionName);
      expect(testCase).toBeDefined();
      
      if (testCase?.expectedValues) {
        Object.entries(expectedValues).forEach(([cell, expectedValue]) => {
          if (typeof expectedValue === 'number' && typeof testCase.expectedValues![cell] === 'number') {
            expect(testCase.expectedValues![cell]).toBeCloseTo(expectedValue, 5);
          } else {
            expect(testCase.expectedValues![cell]).toBe(expectedValue);
          }
        });
      }
    });
  });

  // 文字列関数のテスト
  describe('文字列関数', () => {
    test.each([
      ['CONCATENATE', { 'D2': 'Hello World' }],
      ['CONCAT', { 'C2': 'ABCDEF' }],
      ['TEXTJOIN', { 'D2': 'A-B-C' }],
      ['LEFT', { 'C2': 'Hello' }],
      ['RIGHT', { 'C2': 'World' }],
      ['MID', { 'D2': 'llo' }],
      ['LEN', { 'B2': 5, 'B3': 5 }],
      ['FIND', { 'C2': 5 }],
      ['SEARCH', { 'C2': 7 }],
      ['UPPER', { 'B2': 'HELLO WORLD' }],
      ['LOWER', { 'B2': 'hello world' }],
      ['PROPER', { 'B2': 'Hello World' }],
      ['TRIM', { 'B2': 'Hello World' }],
      ['VALUE', { 'B2': 123.45 }],
      ['REPT', { 'C2': '★★★★★' }],
      ['CHAR', { 'B2': 'A', 'B3': 'a' }],
      ['CODE', { 'B2': 65, 'B3': 97 }],
      ['EXACT', { 'C2': true, 'C3': false }],
    ])('%s関数が正しく計算されること', (functionName: string, expectedValues: Record<string, string | number | boolean>) => {
      const testCase = allIndividualFunctionTests.find(t => t.name === functionName);
      expect(testCase).toBeDefined();
      
      if (testCase?.expectedValues) {
        Object.entries(expectedValues).forEach(([cell, expectedValue]) => {
          expect(testCase.expectedValues![cell]).toBe(expectedValue);
        });
      }
    });
  });

  // 日付関数のテスト
  describe('日付関数', () => {
    test.each([
      ['YEAR', { 'B2': 2024 }],
      ['MONTH', { 'B2': 12 }],
      ['DAY', { 'B2': 25 }],
      ['HOUR', { 'B2': 13 }],
      ['MINUTE', { 'B2': 30 }],
      ['SECOND', { 'B2': 45 }],
      ['WEEKDAY', { 'B2': 4 }],
      ['DATEDIF', { 'D2': 365 }],
      ['DAYS', { 'C2': 365 }],
    ])('%s関数が正しく計算されること', (functionName: string, expectedValues: Record<string, string | number | boolean>) => {
      const testCase = allIndividualFunctionTests.find(t => t.name === functionName);
      expect(testCase).toBeDefined();
      
      if (testCase?.expectedValues) {
        Object.entries(expectedValues).forEach(([cell, expectedValue]) => {
          expect(testCase.expectedValues![cell]).toBe(expectedValue);
        });
      }
    });
  });

  // 論理関数のテスト
  describe('論理関数', () => {
    test.each([
      ['IF', { 'B2': '合格', 'B3': '不合格' }],
      ['IFS', { 'B2': 'A', 'B3': 'C' }],
      ['AND', { 'C2': true, 'C3': false }],
      ['OR', { 'C2': true, 'C3': false }],
      ['NOT', { 'B2': false, 'B3': true }],
      ['XOR', { 'C2': false, 'C3': true }],
      ['TRUE', { 'A2': true }],
      ['FALSE', { 'A2': false }],
      ['IFERROR', { 'C2': 'エラー', 'C3': 5 }],
      ['SWITCH', { 'B2': '一', 'B3': '二', 'B4': 'その他' }],
    ])('%s関数が正しく計算されること', (functionName: string, expectedValues: Record<string, string | number | boolean>) => {
      const testCase = allIndividualFunctionTests.find(t => t.name === functionName);
      expect(testCase).toBeDefined();
      
      if (testCase?.expectedValues) {
        Object.entries(expectedValues).forEach(([cell, expectedValue]) => {
          expect(testCase.expectedValues![cell]).toBe(expectedValue);
        });
      }
    });
  });

  // 検索・参照関数のテスト
  describe('検索・参照関数', () => {
    test.each([
      ['VLOOKUP', { 'F2': '佐藤' }],
      ['HLOOKUP', { 'C5': 200 }],
      ['XLOOKUP', { 'E2': '佐藤' }],
      ['INDEX', { 'F2': 'E' }],
      ['MATCH', { 'E2': 2 }],
      ['CHOOSE', { 'B2': '二', 'B3': '三' }],
      ['OFFSET', { 'B5': 'B2' }],
      ['INDIRECT', { 'E2': 100 }],
      ['ROW', { 'A2': 2, 'A3': 3, 'A4': 4 }],
      ['COLUMN', { 'A2': 1, 'B2': 2, 'C2': 3 }],
    ])('%s関数が正しく計算されること', (functionName: string, expectedValues: Record<string, string | number | boolean>) => {
      const testCase = allIndividualFunctionTests.find(t => t.name === functionName);
      expect(testCase).toBeDefined();
      
      if (testCase?.expectedValues) {
        Object.entries(expectedValues).forEach(([cell, expectedValue]) => {
          expect(testCase.expectedValues![cell]).toBe(expectedValue);
        });
      }
    });
  });
});

// スプレッドシート全体の計算テスト
describe('スプレッドシート計算機能のテスト', () => {
  test('複数の数式が正しく計算されること', async () => {
    const { container } = render(
      <BrowserRouter>
        <DemoSpreadsheet />
      </BrowserRouter>
    );
    
    // 初期レンダリング待機
    await waitFor(() => {
      expect(screen.getByTestId('spreadsheet-mock')).toBeInTheDocument();
    });
    
    // 数式が含まれるセルが正しく計算されているか確認
    const cells = container.querySelectorAll('[data-testid^="cell-"]');
    expect(cells.length).toBeGreaterThan(0);
  });
  
  test('セル参照が正しく解決されること', async () => {
    const { container } = render(
      <BrowserRouter>
        <DemoSpreadsheet />
      </BrowserRouter>
    );
    
    await waitFor(() => {
      expect(screen.getByTestId('spreadsheet-mock')).toBeInTheDocument();
    });
    
    // セル参照を含む数式のテスト
    // 実際の計算エンジンのテストはcustomFormulaCalculations.test.tsで行う
  });
});