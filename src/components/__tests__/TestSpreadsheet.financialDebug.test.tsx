import { describe, it, expect } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';

describe('財務関数のデバッグ', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('PMT関数のテスト結果を確認', async () => {
    const user = userEvent.setup();
    const { container } = renderComponent();
    
    // 財務カテゴリを選択
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '07. 財務');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    // PMT関数を選択
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'PMT');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });
    
    // セルの値を確認
    const cells = container.querySelectorAll('td');
    const results: Record<string, string> = {};
    
    cells.forEach((cell, index) => {
      const row = Math.floor(index / 10);
      const col = index % 10;
      const cellContent = cell.textContent || '';
      if (cellContent) {
        const cellAddr = `${String.fromCharCode(65 + col)}${row + 1}`;
        results[cellAddr] = cellContent;
      }
    });
    
    console.log('PMT関数の結果:');
    console.log('D2セルの値:', results['D2']);
    console.log('期待値: 8606.64');
    
    // エラーメッセージの確認
    const errorMessages = container.querySelectorAll('.text-red-600');
    errorMessages.forEach((error) => {
      console.log('エラー:', error.textContent);
    });
    
    // テスト結果セクションを確認
    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      const resultItems = container.querySelectorAll('.bg-red-50, .bg-green-50');
      resultItems.forEach((item) => {
        console.log('テスト結果アイテム:', item.textContent);
      });
    }
    
    expect(true).toBe(true); // 情報収集のため常にパス
  });

  it('FV関数のテスト結果を確認', async () => {
    const user = userEvent.setup();
    const { container } = renderComponent();
    
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '07. 財務');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'FV');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });
    
    const cells = container.querySelectorAll('td');
    const results: Record<string, string> = {};
    
    cells.forEach((cell, index) => {
      const row = Math.floor(index / 10);
      const col = index % 10;
      const cellContent = cell.textContent || '';
      if (cellContent) {
        const cellAddr = `${String.fromCharCode(65 + col)}${row + 1}`;
        results[cellAddr] = cellContent;
      }
    });
    
    console.log('\\nFV関数の結果:');
    console.log('D2セルの値:', results['D2']);
    console.log('期待値: 10189.60');
    
    expect(true).toBe(true);
  });
}, 30000);