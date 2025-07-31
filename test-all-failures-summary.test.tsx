import { describe, it, expect } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from './src/components/TestSpreadsheet';
import { allIndividualFunctionTests } from './src/data/individualFunctionTests';

describe('全関数の失敗サマリー', () => {
  it('カテゴリ別の失敗をまとめる', async () => {
    const user = userEvent.setup();
    const categorySummary: Record<string, { total: number, failed: string[] }> = {};
    
    const functionsWithExpectedValues = allIndividualFunctionTests.filter(
      func => func.expectedValues && Object.keys(func.expectedValues).length > 0
    );
    
    // カテゴリ別に集計
    const categories = Array.from(new Set(functionsWithExpectedValues.map(f => f.category))).sort();
    
    for (const category of categories) {
      const categoryFunctions = functionsWithExpectedValues.filter(f => f.category === category);
      const failedFunctions: string[] = [];
      
      for (const func of categoryFunctions) {
        const { container, unmount } = render(
          <BrowserRouter>
            <TestSpreadsheet />
          </BrowserRouter>
        );
        
        // カテゴリと関数を選択
        await user.selectOptions(screen.getByLabelText('カテゴリを選択'), category);
        await waitFor(() => {
          expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
        });
        await user.selectOptions(screen.getByLabelText('関数を選択'), func.name);
        
        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        }, { timeout: 5000 });
        
        // 失敗をチェック
        const failureBadges = container.querySelectorAll('.bg-red-100.text-red-800');
        if (failureBadges.length > 0) {
          failedFunctions.push(`${func.name} (${failureBadges.length}個の失敗)`);
        }
        
        unmount();
      }
      
      categorySummary[category] = {
        total: categoryFunctions.length,
        failed: failedFunctions
      };
    }
    
    // 結果を出力
    console.log('\n=== カテゴリ別の失敗サマリー ===\n');
    let totalFunctions = 0;
    let totalFailed = 0;
    
    Object.entries(categorySummary).forEach(([category, summary]) => {
      console.log(`${category}:`);
      console.log(`  総関数数: ${summary.total}`);
      console.log(`  失敗数: ${summary.failed.length}`);
      if (summary.failed.length > 0) {
        console.log(`  失敗した関数:`);
        summary.failed.forEach(func => console.log(`    - ${func}`));
      }
      console.log('');
      
      totalFunctions += summary.total;
      totalFailed += summary.failed.length;
    });
    
    console.log(`総計: ${totalFunctions}個の関数中、${totalFailed}個が失敗\n`);
    
    // テストは情報収集のために常にパス
    expect(true).toBe(true);
  });
});