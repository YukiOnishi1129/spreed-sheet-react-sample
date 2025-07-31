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

describe('特定の関数の動作確認', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('SERIESSUM関数の詳細確認', async () => {
    const user = userEvent.setup();
    renderComponent();

    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '01. 数学');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'SERIESSUM');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    // テスト結果の詳細を出力
    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      const resultContainer = testResultsSection.parentElement;
      console.log('SERIESSUM テスト結果:', resultContainer?.textContent);
      
      // G2セルの結果を探す
      const cellG2 = screen.queryByText(/G2:/);
      if (cellG2) {
        const container = cellG2.closest('.flex');
        console.log('G2セルの詳細:', container?.textContent);
      }
    }
  });

  it('TEXT関数の詳細確認', async () => {
    const user = userEvent.setup();
    renderComponent();

    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '03. 文字列');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'TEXT');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    // テスト結果の詳細を出力
    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      const resultContainer = testResultsSection.parentElement;
      console.log('TEXT テスト結果:', resultContainer?.textContent);
      
      // C2セルの結果を探す
      const cellC2 = screen.queryByText(/C2:/);
      if (cellC2) {
        const container = cellC2.closest('.flex');
        console.log('C2セルの詳細:', container?.textContent);
      }
    }
  });

  it('EDATE関数の詳細確認', async () => {
    const user = userEvent.setup();
    renderComponent();

    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '04. 日付');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'EDATE');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    // テスト結果の詳細を出力
    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      const resultContainer = testResultsSection.parentElement;
      console.log('EDATE テスト結果:', resultContainer?.textContent);
      
      // C2セルの結果を探す
      const cellC2 = screen.queryByText(/C2:/);
      if (cellC2) {
        const container = cellC2.closest('.flex');
        console.log('C2セルの詳細:', container?.textContent);
      }
    }
  });
});