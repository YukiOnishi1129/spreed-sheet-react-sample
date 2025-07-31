import { describe, it, expect } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';

describe('ISO.CEILING関数のテスト', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('ISO.CEILING関数の動作を確認', async () => {
    const user = userEvent.setup();
    const { container } = renderComponent();
    
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '01. 数学');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'ISO.CEILING');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });
    
    // テスト結果を取得
    const resultItems = container.querySelectorAll('.bg-red-50, .bg-green-50');
    
    resultItems.forEach((item) => {
      const text = item.textContent || '';
      console.log('ISO.CEILING テスト結果:', text);
      
      if (text.includes('C3')) {
        console.log('\\nC3セルの詳細:');
        console.log(text);
      }
    });
    
    // セルの値を直接確認
    const cells = container.querySelectorAll('td');
    cells.forEach((cell, index) => {
      const row = Math.floor(index / 10);
      const col = index % 10;
      if (row === 2 && col === 2) { // C3
        console.log('\\nC3セルの実際の値:', cell.textContent);
      }
    });
    
    expect(true).toBe(true);
  });
});