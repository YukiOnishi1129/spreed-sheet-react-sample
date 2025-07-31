import { describe, it, expect, vi } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';

// react-spreadsheetのモック（シンプル版）
vi.mock('react-spreadsheet', () => ({
  default: vi.fn(() => <div data-testid="spreadsheet-mock">Spreadsheet</div>),
  Selection: {},
  Point: class Point {}
}));

describe('EXPAND関数の動作確認', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('EXPAND関数の詳細確認', async () => {
    const user = userEvent.setup();
    renderComponent();

    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '12. 動的配列');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'EXPAND');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    // テスト結果の詳細を出力
    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      const resultContainer = testResultsSection.parentElement;
      console.log('EXPAND テスト結果:', resultContainer?.textContent);
      
      // 失敗数を確認
      const summaryText = screen.queryByText(/合計:/);
      if (summaryText) {
        console.log('サマリー:', summaryText.textContent);
      }
      
      // 特定のセルの結果を確認
      const cells = ['C2', 'D2', 'E2', 'C3', 'D3', 'E3', 'C4', 'D4', 'E4', 'C5', 'D5', 'E5'];
      for (const cell of cells) {
        const cellElement = screen.queryByText(new RegExp(`${cell}:`));
        if (cellElement) {
          const container = cellElement.closest('.flex');
          console.log(`${cell}の詳細:`, container?.textContent);
        }
      }
    }
  });
});