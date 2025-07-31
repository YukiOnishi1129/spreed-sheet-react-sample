import { describe, it, expect } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';
import { allIndividualFunctionTests } from '../../data/individualFunctionTests';

describe('完全な失敗検出（テスト結果なしも含む）', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('テスト結果がない関数も失敗としてカウント', async () => {
    const user = userEvent.setup();
    
    // すべての関数を対象とする
    const allFunctions = allIndividualFunctionTests;
    
    // カテゴリ別に集計
    const categoryResults: Record<string, {
      total: number;
      withExpectedValues: number;
      withoutExpectedValues: number;
      tested: number;
      failedWithTests: { name: string; failures: number }[];
      noTestResults: string[];
    }> = {};
    
    // カテゴリ一覧を取得
    const categories = Array.from(new Set(allFunctions.map(f => f.category))).sort();
    
    console.log('\n=== 完全な失敗検出（テスト結果なしも含む） ===\n');
    
    // 各カテゴリをテスト
    for (const category of categories) {
      const categoryFunctions = allFunctions.filter(f => f.category === category);
      const withExpected = categoryFunctions.filter(f => f.expectedValues && Object.keys(f.expectedValues).length > 0);
      const withoutExpected = categoryFunctions.filter(f => !f.expectedValues || Object.keys(f.expectedValues).length === 0);
      
      categoryResults[category] = { 
        total: categoryFunctions.length,
        withExpectedValues: withExpected.length,
        withoutExpectedValues: withoutExpected.length,
        tested: 0,
        failedWithTests: [],
        noTestResults: []
      };
      
      // 期待値がない関数のリスト
      categoryResults[category].noTestResults = withoutExpected.map(f => f.name);
      
      // 期待値がある関数から最大3個をテスト
      const sampled = withExpected.slice(0, 3);
      
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
          
          categoryResults[category].tested++;
          
          // テスト結果セクションの存在を確認
          const testResultsSection = screen.queryByText('テスト結果');
          
          if (!testResultsSection) {
            // テスト結果セクションがない場合も失敗としてカウント
            console.log(`${category} - ${func.name}: テスト結果セクションなし（失敗）`);
            categoryResults[category].failedWithTests.push({
              name: func.name,
              failures: -1 // テスト結果なしを示す特別な値
            });
          } else {
            // サマリーテキストから失敗数を取得
            const summaryText = screen.queryByText(/合計:/);
            let failureCount = 0;
            
            if (summaryText) {
              const summaryContent = summaryText.textContent || '';
              const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
              failureCount = failureMatch ? parseInt(failureMatch[1]) : 0;
            } else {
              // サマリーがない場合は失敗セルの数を使用
              const failedCells = container.querySelectorAll('.bg-red-50');
              failureCount = failedCells.length;
            }
            
            if (failureCount > 0) {
              categoryResults[category].failedWithTests.push({
                name: func.name,
                failures: failureCount
              });
              console.log(`${category} - ${func.name}: ${failureCount}個の失敗`);
            }
          }
        } catch (error) {
          console.error(`Error testing ${func.name}:`, error);
          // エラーが発生した場合も失敗としてカウント
          categoryResults[category].failedWithTests.push({
            name: func.name,
            failures: -2 // エラーを示す特別な値
          });
        } finally {
          unmount();
        }
      }
    }
    
    // 結果を集計して出力
    console.log('\n=== カテゴリ別完全失敗サマリー ===\n');
    let totalFunctions = 0;
    let totalFailures = 0;
    let totalNoTests = 0;
    
    Object.entries(categoryResults).forEach(([category, result]) => {
      const testedFailures = result.failedWithTests.length;
      const failureRate = result.tested > 0 ? testedFailures / result.tested : 0;
      const estimatedFailuresWithTests = Math.round(failureRate * result.withExpectedValues);
      const totalCategoryFailures = estimatedFailuresWithTests + result.withoutExpectedValues;
      
      totalFunctions += result.total;
      totalFailures += totalCategoryFailures;
      totalNoTests += result.withoutExpectedValues;
      
      console.log(`${category}:`);
      console.log(`  総関数数: ${result.total}`);
      console.log(`  期待値あり: ${result.withExpectedValues}`);
      console.log(`  期待値なし（テスト結果なし）: ${result.withoutExpectedValues}`);
      
      if (result.withExpectedValues > 0) {
        console.log(`  テスト実行数: ${result.tested}`);
        console.log(`  テスト失敗数: ${testedFailures}`);
        console.log(`  推定失敗率: ${(failureRate * 100).toFixed(0)}%`);
      }
      
      console.log(`  推定総失敗数: ${totalCategoryFailures}（テスト結果なし含む）`);
      
      if (result.failedWithTests.length > 0) {
        console.log(`  失敗した関数（テスト済み）:`);
        result.failedWithTests.forEach(({ name, failures }) => {
          if (failures === -1) {
            console.log(`    - ${name}: テスト結果セクションなし`);
          } else if (failures === -2) {
            console.log(`    - ${name}: エラー発生`);
          } else {
            console.log(`    - ${name}: ${failures}個の失敗`);
          }
        });
      }
      
      if (result.noTestResults.length > 0 && result.noTestResults.length <= 10) {
        console.log(`  テスト結果がない関数:`);
        result.noTestResults.forEach(name => {
          console.log(`    - ${name}`);
        });
      } else if (result.noTestResults.length > 10) {
        console.log(`  テスト結果がない関数: ${result.noTestResults.length}個`);
        console.log(`    最初の5個: ${result.noTestResults.slice(0, 5).join(', ')}...`);
      }
      
      console.log('');
    });
    
    console.log(`\n=== 総計 ===`);
    console.log(`総関数数: ${totalFunctions}`);
    console.log(`テスト結果なし: ${totalNoTests}`);
    console.log(`推定総失敗数: ${totalFailures}（テスト結果なし含む）`);
    console.log(`失敗率: ${((totalFailures / totalFunctions) * 100).toFixed(1)}%`);
    
    // カテゴリ別の失敗率を記録（テスト結果なし含む）
    const failureRates: Record<string, number> = {};
    Object.entries(categoryResults).forEach(([category, result]) => {
      const totalFailures = result.failedWithTests.length + result.withoutExpectedValues;
      failureRates[category] = totalFailures / result.total;
    });
    
    console.log('\n=== 失敗率が高いカテゴリ（テスト結果なし含む） ===');
    Object.entries(failureRates)
      .sort((a, b) => b[1] - a[1])
      .filter(([, rate]) => rate > 0)
      .slice(0, 10)
      .forEach(([category, rate]) => {
        const result = categoryResults[category];
        console.log(`${category}: ${(rate * 100).toFixed(0)}% (${result.withoutExpectedValues}個はテスト結果なし)`);
      });
    
    // テストは情報収集のため常にパス
    expect(true).toBe(true);
  }, 120000); // 2分のタイムアウト
});