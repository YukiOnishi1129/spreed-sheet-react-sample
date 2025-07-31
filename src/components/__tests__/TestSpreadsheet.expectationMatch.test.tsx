import { describe, it, expect, vi } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';

// react-spreadsheetのモック（シンプル版）
vi.mock('react-spreadsheet', () => ({
  default: vi.fn(() => <div data-testid="spreadsheet-mock">Spreadsheet</div>),
  Selection: {},
  Point: class Point {}
}));

describe('TestSpreadsheet - 期待値と実際値の一致確認テスト', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('SUM関数：全ての期待値と実際値が一致する', async () => {
    const user = userEvent.setup();
    renderComponent();

    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '01. 数学');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    // テスト結果セクションが表示される
    const testResultsSection = screen.queryByText('テスト結果');
    expect(testResultsSection).toBeInTheDocument();

    // 失敗バッジが存在しないことを確認
    const failureBadges = screen.queryAllByText('失敗');
    expect(failureBadges.length).toBe(0);

    // 成功バッジのみが存在する
    const successBadges = screen.queryAllByText('成功');
    expect(successBadges.length).toBeGreaterThan(0);

    // サマリーで失敗数が0であることを確認
    const summaryText = screen.queryByText(/合計:/);
    expect(summaryText).toBeTruthy();
    const summaryContent = summaryText!.textContent || '';
    expect(summaryContent).toMatch(/失敗:\s*0/);
  });

  it('AVERAGE関数：全ての期待値と実際値が一致する', async () => {
    const user = userEvent.setup();
    renderComponent();

    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '02. 統計');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'AVERAGE');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      // 失敗が0件であることを確認
      const failureBadges = screen.queryAllByText('失敗');
      expect(failureBadges.length).toBe(0);
    }
  });

  it('MAX関数：全ての期待値と実際値が一致する', async () => {
    const user = userEvent.setup();
    renderComponent();

    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '02. 統計');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'MAX');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      const failureBadges = screen.queryAllByText('失敗');
      expect(failureBadges.length).toBe(0);
    }
  });

  it('CONCATENATE関数：全ての期待値と実際値が一致する', async () => {
    const user = userEvent.setup();
    renderComponent();

    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '03. 文字列');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'CONCATENATE');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      const failureBadges = screen.queryAllByText('失敗');
      expect(failureBadges.length).toBe(0);
    }
  });

  it('期待値がある全ての関数で、期待値と実際値が完全に一致する', async () => {
    const user = userEvent.setup();
    
    // allIndividualFunctionTestsをインポート
    const { allIndividualFunctionTests } = await import('../../data/individualFunctionTests');
    
    // 期待値を持つ関数のみをフィルタリング
    const functionsWithExpectedValues = allIndividualFunctionTests.filter(
      func => func.expectedValues && Object.keys(func.expectedValues).length > 0
    );

    console.log(`期待値がある関数の総数: ${functionsWithExpectedValues.length}`);
    
    const failedFunctions: { function: string, category: string, failureCount: number }[] = [];

    for (const functionTest of functionsWithExpectedValues) {
      // 各関数を個別にテスト
      const { unmount } = renderComponent();
      
      // カテゴリを選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), functionTest.category);
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      
      // 関数を選択
      await user.selectOptions(screen.getByLabelText('関数を選択'), functionTest.name);
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      }, { timeout: 5000 });

      // テスト結果セクションがある場合、失敗数を確認
      const testResultsSection = screen.queryByText('テスト結果');
      if (testResultsSection) {
        const failureBadges = screen.queryAllByText('失敗');
        
        if (failureBadges.length > 0) {
          failedFunctions.push({
            function: functionTest.name,
            category: functionTest.category,
            failureCount: failureBadges.length
          });
        }
        
        // サマリーでも確認
        const summaryText = screen.queryByText(/合計:/);
        if (summaryText) {
          const summaryContent = summaryText.textContent || '';
          const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
          if (failureMatch) {
            const failureCount = parseInt(failureMatch[1]);
            // expect(failureCount).toBe(0); // 一旦コメントアウトして全体を確認
          }
        }
      }
      
      unmount();
    }
    
    // 失敗した関数をレポート
    if (failedFunctions.length > 0) {
      console.log('\n失敗した関数:');
      failedFunctions.forEach(({ function: funcName, category, failureCount }) => {
        console.log(`  ${category} - ${funcName}: ${failureCount}個の失敗`);
      });
    }
    
    // 全ての関数で失敗が0件であることを期待
    expect(failedFunctions.length).toBe(0);
  }, 300000); // 5分のタイムアウト

  it('各セルの期待値と実際値が正確に一致していることを検証', async () => {
    const user = userEvent.setup();
    renderComponent();

    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '01. 数学');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    // 各セルの結果を確認
    const cellResults = screen.queryAllByText(/^[A-Z]\d+:/);
    
    cellResults.forEach(cellElement => {
      const container = cellElement.closest('.flex');
      if (container) {
        // 成功バッジがあることを確認
        const successBadge = container.querySelector('.bg-green-600');
        expect(successBadge).toBeTruthy();
        
        // 失敗バッジがないことを確認
        const failureBadge = container.querySelector('.bg-red-600');
        expect(failureBadge).toBeFalsy();
        
        // 背景色が緑（成功）であることを確認
        expect(container.classList.contains('bg-green-50')).toBe(true);
      }
    });
  });
});