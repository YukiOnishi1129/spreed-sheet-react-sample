import { describe, it, expect } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';
import { allIndividualFunctionTests } from '../../data/individualFunctionTests';

describe('全カテゴリの失敗チェック', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('カテゴリ別の失敗を集計', async () => {
    const user = userEvent.setup();
    
    // 期待値を持つ関数のみ
    const functionsWithExpectedValues = allIndividualFunctionTests.filter(
      func => func.expectedValues && Object.keys(func.expectedValues).length > 0
    );
    
    // カテゴリ別に集計
    const categorySummary: Record<string, {
      total: number;
      failed: { name: string; failures: number }[];
    }> = {};
    
    // カテゴリ一覧を取得
    const categories = Array.from(new Set(functionsWithExpectedValues.map(f => f.category))).sort();
    
    console.log('\n=== 全カテゴリの失敗チェック開始 ===\n');
    
    // 各カテゴリをテスト（サンプリング）
    for (const category of categories) {
      const categoryFunctions = functionsWithExpectedValues.filter(f => f.category === category);
      categorySummary[category] = { total: categoryFunctions.length, failed: [] };
      
      // カテゴリから最大5個の関数をサンプリング
      const sampled = categoryFunctions.slice(0, 5);
      
      for (const func of sampled) {
        const { unmount } = renderComponent();
        
        try {
          await user.selectOptions(screen.getByLabelText('カテゴリを選択'), category);
          await waitFor(() => {
            expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
          });
          
          await user.selectOptions(screen.getByLabelText('関数を選択'), func.name);
          
          await waitFor(() => {
            expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
          }, { timeout: 5000 });
          
          // サマリーテキストから失敗数を取得
          const summaryText = screen.queryByText(/合計:/);
          if (summaryText) {
            const summaryContent = summaryText.textContent ?? '';
            const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
            const failureCount = failureMatch ? parseInt(failureMatch[1]) : 0;
            
            if (failureCount > 0) {
              categorySummary[category].failed.push({
                name: func.name,
                failures: failureCount
              });
            }
          }
        } catch (error) {
          console.error(`Error testing ${func.name}:`, error);
        } finally {
          unmount();
        }
      }
    }
    
    // 結果を出力
    console.log('\n=== カテゴリ別失敗サマリー ===\n');
    let totalFailed = 0;
    
    Object.entries(categorySummary).forEach(([category, summary]) => {
      const failedCount = summary.failed.length;
      const estimatedFailed = Math.round((failedCount / Math.min(5, summary.total)) * summary.total);
      
      console.log(`${category}:`);
      console.log(`  総関数数: ${summary.total}`);
      console.log(`  サンプル失敗数: ${failedCount}/5`);
      console.log(`  推定失敗数: ${estimatedFailed}`);
      
      if (summary.failed.length > 0) {
        console.log(`  失敗した関数（サンプル）:`);
        summary.failed.forEach(({ name, failures }) => {
          console.log(`    - ${name}: ${failures}個の失敗`);
        });
      }
      console.log('');
      
      totalFailed += estimatedFailed;
    });
    
    console.log(`推定総失敗数: ${totalFailed}個の関数\n`);
    
    // テストは情報収集のため常にパス
    expect(true).toBe(true);
  }, 60000); // 1分のタイムアウト
});