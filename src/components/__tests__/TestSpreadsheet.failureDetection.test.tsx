import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';

// カスタム計算関数をモック
vi.mock('../utils/customFormulaCalculations', () => ({
  calculateSingleFormula: vi.fn((formula: string) => {
    // 特定の数式で意図的に異なる値を返す
    if (formula === '=SUM(A1:A3)') {
      return 999; // 期待値と異なる値
    }
    if (formula === '=AVERAGE(B1:B3)') {
      return '#DIV/0!'; // エラー値
    }
    // その他は元の実装を使う
    return formula;
  })
}));

// react-spreadsheetのモック
vi.mock('react-spreadsheet', () => {
  interface CellType {
    value?: string | number | boolean | null;
    formula?: string;
    title?: string;
    'data-formula'?: string;
  }
  
  const Spreadsheet = ({ data }: { data: CellType[][] }) => {
    return (
      <div data-testid="spreadsheet-mock">
        <table>
          <tbody>
            {data.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, colIndex) => {
                  const cellAddress = `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`;
                  return (
                    <td key={colIndex}>
                      <div
                        data-testid={`cell-${cellAddress}`}
                        data-value={cell?.value?.toString() ?? ''}
                        data-formula={cell?.formula ?? ''}
                      >
                        {cell?.value?.toString() ?? ''}
                      </div>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };
  
  return { 
    default: Spreadsheet,
    Point: class Point {
      row: number;
      column: number;
      constructor(row: number, column: number) {
        this.row = row;
        this.column = column;
      }
    }
  };
});

describe('TestSpreadsheet - 失敗検出テスト', () => {
  let user: ReturnType<typeof userEvent.setup>;

  beforeEach(() => {
    user = userEvent.setup();
  });

  afterEach(() => {
    vi.clearAllMocks();
  });

  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  describe('期待値と異なる結果の検出', () => {
    it('計算結果が期待値と異なる場合、失敗と表示される', async () => {
      renderComponent();
      
      // SUMテスト用のモックデータ
      const mockFunctionTests = vi.fn().mockReturnValue([{
        name: 'SUM',
        category: '数学/三角',
        description: '合計を計算',
        data: [
          [10, null, null],
          [20, null, null],
          [30, null, null],
          ['=SUM(A1:A3)', null, null]
        ],
        expectedValues: { 'A4': 60 } // 期待値は60だが、モックは999を返す
      }]);
      
      // individualFunctionTestsをモック
      vi.doMock('../../data/individualFunctionTests', () => ({
        allIndividualFunctionTests: mockFunctionTests()
      }));
      
      // カテゴリ選択
      const categorySelect = screen.getByLabelText('カテゴリを選択');
      await user.selectOptions(categorySelect, '数学/三角');
      
      // 関数選択
      await waitFor(() => {
        const functionSelect = screen.getByLabelText('関数を選択');
        expect(functionSelect).not.toBeDisabled();
      });
      
      const functionSelect = screen.getByLabelText('関数を選択');
      await user.selectOptions(functionSelect, 'SUM');
      
      // 計算完了を待つ
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // テスト結果セクションが表示される
      await waitFor(() => {
        expect(screen.getByText('テスト結果')).toBeInTheDocument();
      });
      
      // 失敗の表示を確認
      const cellA4 = screen.getByText('A4');
      const resultContainer = cellA4.closest('.flex');
      
      // 失敗時は赤色の背景
      expect(resultContainer).toHaveClass('bg-red-50');
      
      // 失敗バッジの確認
      expect(resultContainer).toHaveTextContent('失敗');
      
      // 期待値と実際値の表示
      expect(resultContainer).toHaveTextContent('期待値: 60');
      expect(resultContainer).toHaveTextContent('実際値: 999');
      
      // サマリーの確認
      const summaryText = screen.getByText(/合計:/);
      expect(summaryText).toHaveTextContent('合計: 1');
      expect(summaryText).toHaveTextContent('成功: 0');
      expect(summaryText).toHaveTextContent('失敗: 1');
    });

    it('エラー値が返された場合も失敗として検出される', async () => {
      renderComponent();
      
      // AVERAGEテスト用のモックデータ
      const mockFunctionTests = vi.fn().mockReturnValue([{
        name: 'AVERAGE',
        category: '統計',
        description: '平均を計算',
        data: [
          [100, null],
          [200, null],
          [300, null],
          ['=AVERAGE(B1:B3)', null]
        ],
        expectedValues: { 'B4': 200 } // 期待値は200だが、モックは#DIV/0!を返す
      }]);
      
      vi.doMock('../../data/individualFunctionTests', () => ({
        allIndividualFunctionTests: mockFunctionTests()
      }));
      
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '統計');
      
      await waitFor(() => {
        const functionSelect = screen.getByLabelText('関数を選択');
        expect(functionSelect).not.toBeDisabled();
      });
      
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'AVERAGE');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // エラー値の場合も失敗として表示
      const cellB4 = screen.getByText('B4');
      const resultContainer = cellB4.closest('.flex');
      
      expect(resultContainer).toHaveClass('bg-red-50');
      expect(resultContainer).toHaveTextContent('失敗');
      expect(resultContainer).toHaveTextContent('実際値: "#DIV/0!"');
    });
  });

  describe('複数の期待値での成功/失敗混在', () => {
    it('一部成功、一部失敗の場合を正しく表示', async () => {
      renderComponent();
      
      const mockFunctionTests = vi.fn().mockReturnValue([{
        name: 'MIXED_TEST',
        category: 'テスト',
        description: '混在テスト',
        data: [
          [10, 20],
          ['=A1*2', '=B1*2'], // A2は20（正解）、B2は999（不正解）
        ],
        expectedValues: { 
          'A2': 20,  // これは成功する
          'B2': 40   // これは失敗する（999が返される）
        }
      }]);
      
      // calculateSingleFormulaのモックを更新
      const { calculateSingleFormula } = await import('../utils/customFormulaCalculations');
      vi.mocked(calculateSingleFormula).mockImplementation((formula: string) => {
        if (formula === '=A1*2') return 20; // 正解
        if (formula === '=B1*2') return 999; // 不正解
        return formula;
      });
      
      vi.doMock('../../data/individualFunctionTests', () => ({
        allIndividualFunctionTests: mockFunctionTests()
      }));
      
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), 'テスト');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'MIXED_TEST');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // A2は成功
      const cellA2 = screen.getByText('A2');
      const containerA2 = cellA2.closest('.flex');
      expect(containerA2).toHaveClass('bg-green-50');
      expect(containerA2).toHaveTextContent('成功');
      
      // B2は失敗
      const cellB2 = screen.getByText('B2');
      const containerB2 = cellB2.closest('.flex');
      expect(containerB2).toHaveClass('bg-red-50');
      expect(containerB2).toHaveTextContent('失敗');
      
      // サマリーの確認
      const summaryText = screen.getByText(/合計:/);
      expect(summaryText).toHaveTextContent('合計: 2');
      expect(summaryText).toHaveTextContent('成功: 1');
      expect(summaryText).toHaveTextContent('失敗: 1');
    });
  });

  describe('浮動小数点の精度問題', () => {
    it('小数の計算で近似値の許容範囲内なら成功とする', async () => {
      renderComponent();
      
      const mockFunctionTests = vi.fn().mockReturnValue([{
        name: 'FLOAT_TEST',
        category: 'テスト',
        description: '浮動小数点テスト',
        data: [
          ['=0.1+0.2'], // JavaScriptでは0.30000000000000004
        ],
        expectedValues: { 'A1': 0.3 }
      }]);
      
      const { calculateSingleFormula } = await import('../utils/customFormulaCalculations');
      vi.mocked(calculateSingleFormula).mockImplementation(() => {
        return 0.30000000000000004; // 浮動小数点誤差あり
      });
      
      vi.doMock('../../data/individualFunctionTests', () => ({
        allIndividualFunctionTests: mockFunctionTests()
      }));
      
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), 'テスト');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'FLOAT_TEST');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // 近似値の範囲内なので成功とする
      const cellA1 = screen.getByText('A1');
      const container = cellA1.closest('.flex');
      expect(container).toHaveClass('bg-green-50');
      expect(container).toHaveTextContent('成功');
    });

    it('許容範囲を超える差があれば失敗とする', async () => {
      renderComponent();
      
      const mockFunctionTests = vi.fn().mockReturnValue([{
        name: 'FLOAT_FAIL_TEST',
        category: 'テスト',
        description: '浮動小数点失敗テスト',
        data: [
          ['=PI()'],
        ],
        expectedValues: { 'A1': 3.14 } // 3.14159...を期待
      }]);
      
      const { calculateSingleFormula } = await import('../utils/customFormulaCalculations');
      vi.mocked(calculateSingleFormula).mockImplementation(() => {
        return 3.14159265359; // より精密な値
      });
      
      vi.doMock('../../data/individualFunctionTests', () => ({
        allIndividualFunctionTests: mockFunctionTests()
      }));
      
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), 'テスト');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'FLOAT_FAIL_TEST');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // 0.01を超える差があるので失敗
      const cellA1 = screen.getByText('A1');
      const container = cellA1.closest('.flex');
      expect(container).toHaveClass('bg-red-50');
      expect(container).toHaveTextContent('失敗');
    });
  });
});