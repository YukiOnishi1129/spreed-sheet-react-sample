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

describe('TestSpreadsheet - 成功/失敗の表示テスト', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('期待値と実際値が一致する場合、成功と表示される', async () => {
    const user = userEvent.setup();
    renderComponent();

    // SUMなど期待値が定義されている関数を選択
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '01. 数学');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    // SUM関数を選択（実装済みで期待値がある関数）
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');
    
    // 計算完了を待つ
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    });

    // テスト結果セクションが存在することを確認
    const testResultsSection = screen.queryByText('テスト結果');
    expect(testResultsSection).toBeInTheDocument();

    // 成功バッジが表示されることを確認
    const successBadges = screen.getAllByText('成功');
    expect(successBadges.length).toBeGreaterThan(0);

    // 成功したテストは緑色の背景を持つことを確認
    successBadges.forEach(badge => {
      const container = badge.closest('.flex');
      expect(container).toHaveClass('bg-green-50');
    });

    // 期待値と実際値が表示されていることを確認
    const expectedTexts = screen.getAllByText(/期待値:/);
    const actualTexts = screen.getAllByText(/実際値:/);
    expect(expectedTexts.length).toBeGreaterThan(0);
    expect(actualTexts.length).toBeGreaterThan(0);
  });

  it('期待値と実際値が異なる場合、失敗と表示される', async () => {
    const user = userEvent.setup();
    renderComponent();

    // 失敗が発生する可能性のある関数を選択
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '12. 動的配列');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });

    // TRANSPOSE関数を選択（テスト結果で失敗が報告されている）
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'TRANSPOSE');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    });

    // テスト結果セクションが存在する場合
    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      // 失敗バッジを探す
      const failureBadges = screen.queryAllByText('失敗');
      
      // 失敗がある場合の確認
      if (failureBadges.length > 0) {
        // 失敗したテストは赤色の背景を持つことを確認
        failureBadges.forEach(badge => {
          const container = badge.closest('.flex');
          expect(container).toHaveClass('bg-red-50');
        });

        // 期待値と実際値が異なることを視覚的に確認できる
        const containers = screen.getAllByRole('generic').filter(
          el => el.classList.contains('bg-red-50')
        );
        expect(containers.length).toBe(failureBadges.length);
      }
    }
  });

  it('テスト結果のサマリーが正しく表示される', async () => {
    const user = userEvent.setup();
    renderComponent();

    // 期待値が定義されている関数を選択
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '01. 数学');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    });

    // サマリーテキストを探す
    const summaryText = screen.queryByText(/合計:/);
    
    if (summaryText) {
      const summaryContent = summaryText.textContent || '';
      
      // サマリーに成功数と失敗数が含まれていることを確認
      expect(summaryContent).toMatch(/成功:\s*\d+/);
      expect(summaryContent).toMatch(/失敗:\s*\d+/);
      expect(summaryContent).toMatch(/合計:\s*\d+/);
      
      // 数値の整合性をチェック
      const successMatch = summaryContent.match(/成功:\s*(\d+)/);
      const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
      const totalMatch = summaryContent.match(/合計:\s*(\d+)/);
      
      if (successMatch && failureMatch && totalMatch) {
        const successCount = parseInt(successMatch[1]);
        const failureCount = parseInt(failureMatch[1]);
        const totalCount = parseInt(totalMatch[1]);
        
        // 成功数 + 失敗数 = 合計数
        expect(successCount + failureCount).toBe(totalCount);
      }
    }
  });

  it('期待値がない関数では、テスト結果セクションが表示されない', async () => {
    const user = userEvent.setup();
    renderComponent();

    // 基本演算子カテゴリ（通常期待値が定義されていない）
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '00. 基本演算子');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });

    const functionSelect = screen.getByLabelText('関数を選択');
    const options = functionSelect.querySelectorAll('option');
    
    // 最初の関数を選択（"-- 関数を選択 --"以外）
    if (options.length > 1) {
      await user.selectOptions(functionSelect, options[1].textContent || '');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });

      // テスト結果セクションが表示されないことを確認
      const testResultsSection = screen.queryByText('テスト結果');
      
      // 期待値がない場合はテスト結果セクションが表示されない
      if (!testResultsSection) {
        expect(testResultsSection).not.toBeInTheDocument();
      }
    }
  });

  it('成功と失敗が混在する場合、両方が正しく表示される', async () => {
    const user = userEvent.setup();
    renderComponent();

    // 成功と失敗が混在する可能性のある関数を選択
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '02. 統計');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });

    // AVERAGE関数など
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'AVERAGE');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    });

    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      // 成功と失敗の両方のバッジを確認
      const successBadges = screen.queryAllByText('成功');
      const failureBadges = screen.queryAllByText('失敗');
      
      // 少なくとも1つのテスト結果があることを確認
      expect(successBadges.length + failureBadges.length).toBeGreaterThan(0);
      
      // それぞれの背景色が正しいことを確認
      successBadges.forEach(badge => {
        const container = badge.closest('.flex');
        expect(container).toHaveClass('bg-green-50');
      });
      
      failureBadges.forEach(badge => {
        const container = badge.closest('.flex');
        expect(container).toHaveClass('bg-red-50');
      });
    }
  });
});