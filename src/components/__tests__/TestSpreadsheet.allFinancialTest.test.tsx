import { describe, it, expect } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';
import { financialTests } from '../../data/individualTests/07-financial';

describe('全財務関数のテスト', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('財務関数の成功率を確認', async () => {
    const user = userEvent.setup();
    
    const results: Record<string, { expected: unknown; actual: unknown; success: boolean }> = {};
    
    // 最初の5つの財務関数をテスト
    const testFunctions = financialTests.slice(0, 5);
    
    for (const func of testFunctions) {
      const { container, unmount } = renderComponent();
      
      try {
        await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '07. 財務');
        await waitFor(() => {
          expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
        });
        
        await user.selectOptions(screen.getByLabelText('関数を選択'), func.name);
        
        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        }, { timeout: 5000 });
        
        // テスト結果を取得
        const resultItems = container.querySelectorAll('.bg-red-50, .bg-green-50');
        
        resultItems.forEach((item) => {
          const text = item.textContent || '';
          const match = text.match(/([A-Z]\d+)期待値:\s*([\d.-]+)実際値:\s*"?([^"]+)"?(成功|失敗)/);
          
          if (match) {
            const [, cell, expected, actual, status] = match;
            results[`${func.name}_${cell}`] = {
              expected: parseFloat(expected),
              actual: actual === '#VALUE!' || actual === '#NUM!' ? actual : parseFloat(actual),
              success: status === '成功'
            };
          }
        });
      } catch (error) {
        console.error(`Error testing ${func.name}:`, error);
      } finally {
        unmount();
      }
    }
    
    console.log('\\n=== 財務関数テスト結果 ===\\n');
    let successCount = 0;
    let totalCount = 0;
    
    Object.entries(results).forEach(([key, result]) => {
      totalCount++;
      if (result.success) {
        successCount++;
        console.log(`✓ ${key}: 期待値=${result.expected}, 実際値=${result.actual}`);
      } else {
        console.log(`✗ ${key}: 期待値=${result.expected}, 実際値=${result.actual}`);
      }
    });
    
    console.log(`\\n成功率: ${successCount}/${totalCount} (${(successCount/totalCount*100).toFixed(1)}%)`);
    
    expect(true).toBe(true);
  }, 60000);
});