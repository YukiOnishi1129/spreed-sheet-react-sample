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

describe('SKEW関数の動作確認', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('SKEW関数の計算確認', async () => {
    const user = userEvent.setup();
    renderComponent();

    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '02. 統計');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'SKEW');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });

    // テスト結果の詳細を出力
    const testResultsSection = screen.queryByText('テスト結果');
    if (testResultsSection) {
      const cellF2 = screen.queryByText(/F2:/);
      if (cellF2) {
        const container = cellF2.closest('.flex');
        console.log('F2セルの詳細:', container?.textContent);
      }
    }
    
    // 手動計算で確認
    const values = [3, 4, 5, 2, 1];
    const n = values.length;
    const mean = values.reduce((sum, val) => sum + val, 0) / n;
    const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / (n - 1);
    const stdDev = Math.sqrt(variance);
    const thirdMoment = values.reduce((sum, val) => sum + Math.pow((val - mean) / stdDev, 3), 0);
    const skewness = (n / ((n - 1) * (n - 2))) * thirdMoment;
    
    console.log('手動計算結果:');
    console.log('平均:', mean);
    console.log('分散:', variance);
    console.log('標準偏差:', stdDev);
    console.log('第3モーメント:', thirdMoment);
    console.log('歪度:', skewness);
    console.log('期待値: 0.406');
    console.log('丸めた値:', Math.round(skewness * 1000) / 1000);
  });
});