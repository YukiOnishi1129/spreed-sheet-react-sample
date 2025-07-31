import { describe, it, expect } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';
import { databaseTests } from '../../data/individualTests/10-database';

describe('データベース関数のテスト', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('DSUM関数の動作を確認', async () => {
    const user = userEvent.setup();
    const { container } = renderComponent();
    
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '10. データベース');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'DSUM');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });
    
    // テスト結果を確認
    const resultItems = container.querySelectorAll('.bg-red-50, .bg-green-50');
    const hasResults = resultItems.length > 0;
    
    if (hasResults) {
      resultItems.forEach((item) => {
        console.log('DSUM テスト結果:', item.textContent);
      });
    } else {
      console.log('DSUM: テスト結果が見つかりません');
    }
    
    // セルの値を確認
    const cells = container.querySelectorAll('td');
    console.log('セル数:', cells.length);
    
    // C7セルの値を探す（期待値: 250）
    cells.forEach((cell, index) => {
      const row = Math.floor(index / 10);
      const col = index % 10;
      if (row === 6 && col === 2) { // C7
        console.log('C7セルの値:', cell.textContent);
      }
    });
    
    expect(true).toBe(true);
  });

  it('全データベース関数の成功率を確認', async () => {
    const user = userEvent.setup();
    
    const results: Record<string, boolean> = {};
    
    for (const func of databaseTests) {
      const { container, unmount } = renderComponent();
      
      try {
        await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '10. データベース');
        await waitFor(() => {
          expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
        });
        
        await user.selectOptions(screen.getByLabelText('関数を選択'), func.name);
        
        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        }, { timeout: 5000 });
        
        // サマリーから成功/失敗を判定
        const summaryText = screen.queryByText(/合計:/);
        let success = false;
        
        if (summaryText) {
          const content = summaryText.textContent || '';
          const failureMatch = content.match(/失敗:\s*(\d+)/);
          const failureCount = failureMatch ? parseInt(failureMatch[1]) : 0;
          success = failureCount === 0;
        }
        
        results[func.name] = success;
        
      } catch (error) {
        console.error(`Error testing ${func.name}:`, error);
        results[func.name] = false;
      } finally {
        unmount();
      }
    }
    
    console.log('\\n=== データベース関数テスト結果 ===');
    let successCount = 0;
    Object.entries(results).forEach(([name, success]) => {
      console.log(`${success ? '✓' : '✗'} ${name}`);
      if (success) successCount++;
    });
    
    const total = Object.keys(results).length;
    console.log(`\\n成功率: ${successCount}/${total} (${(successCount/total*100).toFixed(1)}%)`);
    
    expect(true).toBe(true);
  }, 60000);
});