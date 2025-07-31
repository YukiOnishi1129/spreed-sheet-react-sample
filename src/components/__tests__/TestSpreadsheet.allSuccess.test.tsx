import { describe, it, expect, vi } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';
import { allIndividualFunctionTests } from '../../data/individualFunctionTests';

// react-spreadsheetのモック（シンプル版）
vi.mock('react-spreadsheet', () => ({
  default: vi.fn(() => <div data-testid="spreadsheet-mock">Spreadsheet</div>),
  Selection: {},
  Point: class Point {}
}));

describe('TestSpreadsheet - 全関数の成功チェック', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  // 期待値を持つ関数のみをフィルタリング
  const functionsWithExpectedValues = allIndividualFunctionTests.filter(
    func => func.expectedValues && Object.keys(func.expectedValues).length > 0
  );

  it('すべての関数が成功しているかチェック', async () => {
    const user = userEvent.setup();
    const failedFunctions: { function: string, category: string, failures: string[] }[] = [];
    
    for (const functionTest of functionsWithExpectedValues) {
      // 各関数を個別にテスト
      const { unmount } = renderComponent();
      
      // カテゴリを選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), functionTest.category);
      
      // 関数を選択
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), functionTest.name);
      
      // 計算完了を待つ
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      }, { timeout: 10000 });
      
      // テスト結果セクションの存在を確認
      const testResultsSection = screen.queryByText('テスト結果');
      if (testResultsSection && functionTest.expectedValues) {
        const failures: string[] = [];
        
        // 各セルの結果をチェック
        for (const cellAddress of Object.keys(functionTest.expectedValues)) {
          const cellElement = screen.queryByText(cellAddress);
          if (cellElement) {
            const resultContainer = cellElement.closest('.flex');
            if (resultContainer) {
              // 失敗かどうかをチェック
              const isFailure = resultContainer.classList.contains('bg-red-50');
              const hasFailureBadge = resultContainer.textContent?.includes('失敗');
              
              if (isFailure || hasFailureBadge) {
                const textContent = resultContainer.textContent || '';
                failures.push(`${cellAddress}: ${textContent}`);
              }
            }
          }
        }
        
        // サマリーから失敗数を取得
        const summaryText = screen.queryByText(/合計:/);
        if (summaryText) {
          const summaryContent = summaryText.textContent || '';
          const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
          const failureCount = failureMatch ? parseInt(failureMatch[1]) : 0;
          
          if (failureCount > 0 || failures.length > 0) {
            failedFunctions.push({
              function: functionTest.name,
              category: functionTest.category,
              failures
            });
          }
        }
      }
      
      unmount();
    }
    
    // 結果を出力
    console.log('\n=== 全関数の成功/失敗チェック結果 ===');
    console.log(`総関数数: ${functionsWithExpectedValues.length}`);
    console.log(`失敗した関数数: ${failedFunctions.length}`);
    
    if (failedFunctions.length > 0) {
      console.log('\n失敗した関数の詳細:');
      failedFunctions.forEach(({ function: funcName, category, failures }) => {
        console.log(`\n${category} - ${funcName}:`);
        failures.forEach(failure => console.log(`  ${failure}`));
      });
    }
    
    // すべての関数が成功していることを期待
    expect(failedFunctions).toEqual([]);
  }, 300000); // 5分のタイムアウト
  
  // 特定のカテゴリだけをテストする（デバッグ用）
  describe('カテゴリ別テスト', () => {
    it('01. 数学カテゴリの関数をチェック', async () => {
      const user = userEvent.setup();
      renderComponent();
      
      const mathFunctions = functionsWithExpectedValues.filter(f => f.category === '01. 数学');
      const failures: string[] = [];
      
      for (const func of mathFunctions) {
        await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '01. 数学');
        await waitFor(() => {
          expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
        });
        await user.selectOptions(screen.getByLabelText('関数を選択'), func.name);
        
        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        });
        
        // 失敗をチェック
        const failureBadges = screen.queryAllByText('失敗');
        if (failureBadges.length > 0) {
          failures.push(`${func.name}: ${failureBadges.length}個の失敗`);
        }
      }
      
      console.log('\n01. 数学カテゴリの結果:');
      console.log('総関数数:', mathFunctions.length);
      console.log('失敗:', failures);
      
      expect(failures).toEqual([]);
    });
    
    it('12. 動的配列カテゴリの関数をチェック', async () => {
      const user = userEvent.setup();
      renderComponent();
      
      const dynamicArrayFunctions = functionsWithExpectedValues.filter(f => f.category === '12. 動的配列');
      const failures: string[] = [];
      
      for (const func of dynamicArrayFunctions) {
        await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '12. 動的配列');
        await waitFor(() => {
          expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
        });
        await user.selectOptions(screen.getByLabelText('関数を選択'), func.name);
        
        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        });
        
        // 失敗をチェック
        const summaryText = screen.queryByText(/合計:/);
        if (summaryText) {
          const summaryContent = summaryText.textContent || '';
          const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
          const failureCount = failureMatch ? parseInt(failureMatch[1]) : 0;
          
          if (failureCount > 0) {
            failures.push(`${func.name}: ${failureCount}個の失敗`);
          }
        }
      }
      
      console.log('\n12. 動的配列カテゴリの結果:');
      console.log('総関数数:', dynamicArrayFunctions.length);
      console.log('失敗:', failures);
      
      expect(failures).toEqual([]);
    });
  });
});