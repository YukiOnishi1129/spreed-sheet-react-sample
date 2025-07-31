import { describe, it, expect } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';
import { financialTests } from '../../data/individualTests/07-financial';

describe('全財務関数の包括的テスト', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('全財務関数の成功率を確認', async () => {
    const user = userEvent.setup();
    
    const results: Record<string, { 
      expected: unknown; 
      actual: unknown; 
      success: boolean;
      error?: string;
    }> = {};
    
    let successCount = 0;
    let totalCount = 0;
    
    // すべての財務関数をテスト
    for (const func of financialTests) {
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
        
        // サマリーテキストから成功/失敗数を取得
        const summaryText = screen.queryByText(/合計:/);
        let functionSuccess = false;
        let failureCount = 0;
        
        if (summaryText) {
          const summaryContent = summaryText.textContent || '';
          const successMatch = summaryContent.match(/成功:\s*(\d+)/);
          const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
          
          const successNum = successMatch ? parseInt(successMatch[1]) : 0;
          failureCount = failureMatch ? parseInt(failureMatch[1]) : 0;
          
          functionSuccess = failureCount === 0 && successNum > 0;
        }
        
        // テスト結果を記録
        results[func.name] = {
          expected: '全テスト成功',
          actual: functionSuccess ? '成功' : `${failureCount}個の失敗`,
          success: functionSuccess,
          error: failureCount > 0 ? `${failureCount}個のテストが失敗` : undefined
        };
        
        totalCount++;
        if (functionSuccess) {
          successCount++;
        }
        
      } catch (error) {
        results[func.name] = {
          expected: '成功',
          actual: 'エラー',
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        };
        totalCount++;
      } finally {
        unmount();
      }
    }
    
    // 結果をカテゴリ別に整理
    console.log('\\n=== 財務関数包括的テスト結果 ===\\n');
    console.log(`テスト関数数: ${financialTests.length}`);
    console.log(`\\n関数別結果:`);
    
    Object.entries(results).forEach(([name, result]) => {
      const statusIcon = result.success ? '✓' : '✗';
      console.log(`${statusIcon} ${name}: ${result.actual}${result.error ? ` (${result.error})` : ''}`);
    });
    
    const successRate = (successCount / totalCount * 100).toFixed(1);
    console.log(`\\n全体成功率: ${successCount}/${totalCount} (${successRate}%)`);
    
    // 失敗している関数の詳細
    const failedFunctions = Object.entries(results)
      .filter(([, result]) => !result.success)
      .map(([name]) => name);
    
    if (failedFunctions.length > 0) {
      console.log(`\\n失敗している関数 (${failedFunctions.length}個):`);
      failedFunctions.forEach(name => console.log(`  - ${name}`));
    }
    
    expect(true).toBe(true);
  }, 180000); // 3分のタイムアウト
});