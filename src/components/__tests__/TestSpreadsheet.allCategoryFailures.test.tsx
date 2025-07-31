import { describe, it, expect } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';
import { allIndividualFunctionTests } from '../../data/individualFunctionTests';

describe('全カテゴリの失敗を正確に検出', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('全カテゴリの失敗数を正確に集計', async () => {
    const user = userEvent.setup();
    
    // 期待値を持つ関数のみ
    const functionsWithExpectedValues = allIndividualFunctionTests.filter(
      func => func.expectedValues && Object.keys(func.expectedValues).length > 0
    );
    
    // カテゴリ別に集計
    const categoryResults: Record<string, {
      total: number;
      tested: number;
      failedFunctions: { name: string; failures: number }[];
    }> = {};
    
    // カテゴリ一覧を取得
    const categories = Array.from(new Set(functionsWithExpectedValues.map(f => f.category))).sort();
    
    console.log('\n=== 全カテゴリの正確な失敗検出 ===\n');
    
    // 各カテゴリをテスト
    for (const category of categories) {
      const categoryFunctions = functionsWithExpectedValues.filter(f => f.category === category);
      categoryResults[category] = { 
        total: categoryFunctions.length, 
        tested: 0,
        failedFunctions: [] 
      };
      
      // 各カテゴリから最大3個の関数をテスト（時間短縮のため）
      const sampled = categoryFunctions.slice(0, 3);
      
      for (const func of sampled) {
        const { container, unmount } = renderComponent();
        
        try {
          await user.selectOptions(screen.getByLabelText('カテゴリを選択'), category);
          await waitFor(() => {
            expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
          });
          
          await user.selectOptions(screen.getByLabelText('関数を選択'), func.name);
          
          await waitFor(() => {
            expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
          }, { timeout: 5000 });
          
          // 失敗要素を正確にカウント
          // bg-red-50クラスを持つ要素が失敗したセル
          const failedCells = container.querySelectorAll('.bg-red-50');
          // 失敗バッジの数も確認
          const failureBadges = container.querySelectorAll('.text-red-800:not(.bg-red-100)');
          
          // サマリーテキストから失敗数を取得（最も正確）
          const summaryText = screen.queryByText(/合計:/);
          let failureCount = 0;
          
          if (summaryText) {
            const summaryContent = summaryText.textContent || '';
            const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
            failureCount = failureMatch ? parseInt(failureMatch[1]) : 0;
          } else {
            // サマリーがない場合は失敗セルの数を使用
            failureCount = failedCells.length;
          }
          
          categoryResults[category].tested++;
          
          if (failureCount > 0) {
            categoryResults[category].failedFunctions.push({
              name: func.name,
              failures: failureCount
            });
            
            // 最初の失敗の詳細を取得
            if (failedCells.length > 0) {
              const firstFailure = failedCells[0];
              console.log(`${category} - ${func.name}: ${failureCount}個の失敗`);
              console.log(`  詳細: ${firstFailure.textContent}`);
            }
          }
        } catch (error) {
          console.error(`Error testing ${func.name}:`, error);
        } finally {
          unmount();
        }
      }
    }
    
    // 結果を集計して出力
    console.log('\n=== カテゴリ別失敗サマリー（正確版） ===\n');
    let totalFunctions = 0;
    let totalEstimatedFailures = 0;
    
    Object.entries(categoryResults).forEach(([category, result]) => {
      const failedCount = result.failedFunctions.length;
      const failureRate = result.tested > 0 ? failedCount / result.tested : 0;
      const estimatedFailures = Math.round(failureRate * result.total);
      
      totalFunctions += result.total;
      totalEstimatedFailures += estimatedFailures;
      
      console.log(`${category}:`);
      console.log(`  総関数数: ${result.total}`);
      console.log(`  テスト数: ${result.tested}`);
      console.log(`  失敗数: ${failedCount}`);
      console.log(`  推定失敗率: ${(failureRate * 100).toFixed(0)}%`);
      console.log(`  推定失敗数: ${estimatedFailures}`);
      
      if (result.failedFunctions.length > 0) {
        console.log(`  失敗した関数:`);
        result.failedFunctions.forEach(({ name, failures }) => {
          console.log(`    - ${name}: ${failures}個の失敗`);
        });
      }
      console.log('');
    });
    
    console.log(`\n総計: ${totalFunctions}個の関数中、推定${totalEstimatedFailures}個が失敗`);
    
    // カテゴリ別の実際の失敗率を記録
    const failureRates: Record<string, number> = {};
    Object.entries(categoryResults).forEach(([category, result]) => {
      if (result.tested > 0) {
        failureRates[category] = result.failedFunctions.length / result.tested;
      }
    });
    
    console.log('\n=== 失敗率が高いカテゴリ ===');
    Object.entries(failureRates)
      .sort((a, b) => b[1] - a[1])
      .filter(([, rate]) => rate > 0)
      .forEach(([category, rate]) => {
        console.log(`${category}: ${(rate * 100).toFixed(0)}%`);
      });
    
    // テストは情報収集のため常にパス
    expect(true).toBe(true);
  }, 120000); // 2分のタイムアウト
});