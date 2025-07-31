import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';

// react-spreadsheetのモックを最初に定義
vi.mock('react-spreadsheet', () => {
  return {
    default: vi.fn(() => <div data-testid="spreadsheet-mock">Spreadsheet</div>),
    Selection: {},
    Point: class Point {}
  };
});

import TestSpreadsheet from '../TestSpreadsheet';

describe('TestSpreadsheet - インタラクティブテスト', () => {
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

  describe('値変更による再計算', () => {
    it('SUM関数: セルの値を変更すると合計が更新される', async () => {
      renderComponent();
      
      // SUM関数を選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '数学/三角');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');
      
      // 初期計算完了を待つ
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // 初期の合計値を取得
      const sumCell = screen.getByTestId('cell-A4');
      const initialSum = sumCell.getAttribute('value');
      
      // A1セルの値を変更
      const cellA1 = screen.getByTestId('cell-A1');
      await user.clear(cellA1);
      await user.type(cellA1, '100');
      
      // 再計算をトリガー（実際の実装では onChange イベントで自動的に発生）
      updateCell(0, 0, '100');
      
      // 合計値が更新されることを確認
      await waitFor(() => {
        const newSum = screen.getByTestId('cell-A4').getAttribute('value');
        expect(newSum).not.toBe(initialSum);
      });
    });

    it('AVERAGE関数: 複数セルの値を変更すると平均が更新される', async () => {
      renderComponent();
      
      // AVERAGE関数を選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '統計');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'AVERAGE');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // 複数のセルを変更
      const cellA1 = screen.getByTestId('cell-A1');
      const cellA2 = screen.getByTestId('cell-A2');
      
      await user.clear(cellA1);
      await user.type(cellA1, '50');
      updateCell(0, 0, '50');
      
      await user.clear(cellA2);
      await user.type(cellA2, '100');
      updateCell(1, 0, '100');
      
      // 平均値が75になることを確認
      await waitFor(() => {
        const avgCell = screen.getByTestId('cell-A4');
        expect(avgCell.getAttribute('value')).toBe('75');
      });
    });

    it('IF関数: 条件値を変更すると結果が切り替わる', async () => {
      renderComponent();
      
      // IF関数を選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '論理');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'IF');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // 条件セルの値を変更してTRUE/FALSEを切り替え
      const conditionCell = screen.getByTestId('cell-A1');
      
      // FALSEに変更
      await user.clear(conditionCell);
      await user.type(conditionCell, '0');
      updateCell(0, 0, '0');
      
      // 結果が"FALSE結果"になることを確認
      await waitFor(() => {
        const resultCell = screen.getByTestId('cell-B1');
        expect(resultCell.getAttribute('value')).toContain('FALSE');
      });
      
      // TRUEに戻す
      await user.clear(conditionCell);
      await user.type(conditionCell, '1');
      updateCell(0, 0, '1');
      
      // 結果が"TRUE結果"になることを確認
      await waitFor(() => {
        const resultCell = screen.getByTestId('cell-B1');
        expect(resultCell.getAttribute('value')).toContain('TRUE');
      });
    });

    it('VLOOKUP関数: 検索値を変更すると結果が更新される', async () => {
      renderComponent();
      
      // VLOOKUP関数を選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '検索/参照');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'VLOOKUP');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // 検索値を変更
      const lookupCell = screen.getByTestId('cell-A1');
      await user.clear(lookupCell);
      await user.type(lookupCell, '2');
      updateCell(0, 0, '2');
      
      // 対応する値が返されることを確認
      await waitFor(() => {
        const resultCell = screen.getByTestId('cell-B5');
        expect(resultCell.getAttribute('value')).toBeTruthy();
      });
    });

    it('連鎖的な再計算: 依存関係のあるセルが順次更新される', async () => {
      renderComponent();
      
      // 複数の数式が連鎖する関数を選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '数学/三角');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // 基本値を変更
      const baseCell = screen.getByTestId('cell-A1');
      await user.clear(baseCell);
      await user.type(baseCell, '20');
      updateCell(0, 0, '20');
      
      // 連鎖的に更新されることを確認
      await waitFor(() => {
        // 直接依存するセル
        const sumCell = screen.getByTestId('cell-A4');
        expect(sumCell.getAttribute('value')).not.toBe('');
        
        // さらに依存するセル（もしあれば）
        const dependentCells = screen.getAllByRole('textbox').filter(
          input => input.getAttribute('data-formula')?.includes('A4')
        );
        
        dependentCells.forEach(cell => {
          expect(cell.getAttribute('value')).not.toBe('');
        });
      });
    });
  });

  describe('エラー処理と検証', () => {
    it('無効な値を入力した場合、エラーが表示される', async () => {
      renderComponent();
      
      // 数値を期待する関数を選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '数学/三角');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // 文字列を入力
      const cellA1 = screen.getByTestId('cell-A1');
      await user.clear(cellA1);
      await user.type(cellA1, 'invalid');
      updateCell(0, 0, 'invalid');
      
      // エラーまたは0が表示されることを確認
      await waitFor(() => {
        const sumCell = screen.getByTestId('cell-A4');
        const value = sumCell.getAttribute('value');
        expect(value === '0' || value?.includes('#')).toBeTruthy();
      });
    });

    it('循環参照がある場合、適切に処理される', async () => {
      renderComponent();
      
      // 任意の関数を選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '数学/三角');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // 循環参照を作成（A1が=A1を参照）
      const cellA1 = screen.getByTestId('cell-A1');
      await user.clear(cellA1);
      await user.type(cellA1, '=A1');
      updateCell(0, 0, '=A1');
      
      // 無限ループに陥らず、エラーが表示されることを確認
      await waitFor(() => {
        const value = cellA1.getAttribute('value');
        expect(value?.includes('#') ?? value === '').toBeTruthy();
      });
    });
  });

  describe('パフォーマンステスト', () => {
    it('大量のセル変更でも適切に再計算される', async () => {
      renderComponent();
      
      // 配列関数を選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '配列');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      
      const arrayFunctions = screen.getByLabelText('関数を選択').querySelectorAll('option');
      if (arrayFunctions.length > 1) {
        await user.selectOptions(screen.getByLabelText('関数を選択'), arrayFunctions[1].textContent ?? '');
        
        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        });
        
        // 複数のセルを連続して変更
        const startTime = Date.now();
        
        for (let i = 0; i < 5; i++) {
          const cell = screen.queryByTestId(`cell-A${i + 1}`);
          if (cell) {
            await user.clear(cell);
            await user.type(cell, String(i * 10));
            updateCell(i, 0, String(i * 10));
          }
        }
        
        // 再計算が5秒以内に完了することを確認
        await waitFor(() => {
          const endTime = Date.now();
          expect(endTime - startTime).toBeLessThan(5000);
        });
      }
    });
  });

  describe('テスト結果の検証', () => {
    it('期待値と実際値が一致した場合、成功と表示される', async () => {
      renderComponent();
      
      // 期待値が定義されている関数を選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '数学/三角');
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });
      
      // テスト結果セクションが表示される
      const testResultsSection = screen.queryByText('テスト結果');
      if (testResultsSection) {
        // 成功の表示を確認
        const successBadges = screen.getAllByText('成功');
        expect(successBadges.length).toBeGreaterThan(0);
        
        // 成功したテストは緑色の背景
        successBadges.forEach(badge => {
          const container = badge.closest('.flex');
          expect(container).toHaveClass('bg-green-50');
        });
      }
    });

  });
});