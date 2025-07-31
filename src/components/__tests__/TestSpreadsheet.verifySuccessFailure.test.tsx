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

describe('TestSpreadsheet - 期待値と実際値の検証テスト', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('期待値と実際値が一致する場合、成功バッジが表示される', async () => {
    const user = userEvent.setup();
    renderComponent();

    // SUMなど、実装済みで期待値が正しく動作する関数を選択
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '01. 数学');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    // SUM関数を選択
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');
    
    // 計算完了を待つ
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    // テスト結果セクションが表示されることを確認
    const testResultsSection = screen.queryByText('テスト結果');
    expect(testResultsSection).toBeInTheDocument();

    // 成功バッジが少なくとも1つ表示されることを確認
    const successBadges = screen.queryAllByText('成功');
    expect(successBadges.length).toBeGreaterThan(0);

    // 各成功バッジの詳細を確認
    successBadges.forEach(badge => {
      const container = badge.closest('.flex');
      expect(container).toHaveClass('bg-green-50');
      
      // 期待値と実際値が表示されていることを確認
      const containerText = container?.textContent || '';
      expect(containerText).toMatch(/期待値:/);
      expect(containerText).toMatch(/実際値:/);
    });

    // サマリーに成功が含まれることを確認
    const summaryText = screen.queryByText(/合計:/);
    if (summaryText) {
      const summaryContent = summaryText.textContent || '';
      const successMatch = summaryContent.match(/成功:\s*(\d+)/);
      expect(successMatch).toBeTruthy();
      const successCount = parseInt(successMatch![1]);
      expect(successCount).toBeGreaterThan(0);
    }
  });

  it('期待値と実際値が異なる場合、失敗バッジが表示される（意図的に失敗するケース）', async () => {
    const user = userEvent.setup();
    renderComponent();

    // 動的配列カテゴリから未実装の関数を選択
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '12. 動的配列');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });

    // TRANSPOSE関数を選択（未実装で失敗することが予想される）
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'TRANSPOSE');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    // テスト結果セクションが表示される場合
    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      // 失敗バッジを探す
      const failureBadges = screen.queryAllByText('失敗');
      
      // TRANSPOSEは未実装なので失敗することを期待
      expect(failureBadges.length).toBeGreaterThan(0);

      // 各失敗バッジの詳細を確認
      failureBadges.forEach(badge => {
        const container = badge.closest('.flex');
        expect(container).toHaveClass('bg-red-50');
        
        // 期待値と実際値が表示されていることを確認
        const containerText = container?.textContent || '';
        expect(containerText).toMatch(/期待値:/);
        expect(containerText).toMatch(/実際値:/);
      });

      // サマリーに失敗が含まれることを確認
      const summaryText = screen.queryByText(/合計:/);
      if (summaryText) {
        const summaryContent = summaryText.textContent || '';
        const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
        expect(failureMatch).toBeTruthy();
        const failureCount = parseInt(failureMatch![1]);
        expect(failureCount).toBeGreaterThan(0);
      }
    }
  });

  it('期待値通りに動作する関数は全てのテストケースが成功する', async () => {
    const user = userEvent.setup();
    renderComponent();

    // 基本的な数学関数を選択
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '01. 数学');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    // ABS関数を選択（シンプルで確実に動作する関数）
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'ABS');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    // テスト結果セクションの確認
    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      // 失敗バッジが存在しないことを確認
      const failureBadges = screen.queryAllByText('失敗');
      expect(failureBadges.length).toBe(0);

      // 成功バッジのみが存在することを確認
      const successBadges = screen.queryAllByText('成功');
      expect(successBadges.length).toBeGreaterThan(0);

      // サマリーで失敗数が0であることを確認
      const summaryText = screen.queryByText(/合計:/);
      if (summaryText) {
        const summaryContent = summaryText.textContent || '';
        const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
        if (failureMatch) {
          const failureCount = parseInt(failureMatch[1]);
          expect(failureCount).toBe(0);
        }
      }
    }
  });

  it('特定のセルで期待値と実際値を比較して正しく判定される', async () => {
    const user = userEvent.setup();
    renderComponent();

    // SUM関数を選択
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '01. 数学');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    // 特定のセルアドレスを含む要素を探す
    const cellResults = screen.queryAllByText(/^[A-Z]\d+:/);
    
    cellResults.forEach(cellElement => {
      const container = cellElement.closest('.flex');
      if (container) {
        const containerText = container.textContent || '';
        
        // 期待値と実際値を抽出
        const expectedMatch = containerText.match(/期待値:\s*([^,]+)/);
        const actualMatch = containerText.match(/実際値:\s*([^)]+)/);
        
        if (expectedMatch && actualMatch) {
          const expectedValue = expectedMatch[1].trim();
          const actualValue = actualMatch[1].trim();
          
          // 成功/失敗のバッジを確認
          const successBadge = container.querySelector('.bg-green-600');
          const failureBadge = container.querySelector('.bg-red-600');
          
          if (expectedValue === actualValue) {
            // 値が一致する場合は成功バッジが表示される
            expect(successBadge).toBeTruthy();
            expect(failureBadge).toBeFalsy();
          } else {
            // 値が異なる場合は失敗バッジが表示される
            expect(failureBadge).toBeTruthy();
            expect(successBadge).toBeFalsy();
          }
        }
      }
    });
  });
});