import { describe, it, expect, vi } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';

interface MockCellData {
  value?: unknown;
  formula?: string;
}

// react-spreadsheetのモック
vi.mock('react-spreadsheet', () => ({
  default: ({ data }: { data: MockCellData[][] }) => (
    <div data-testid="spreadsheet">
      <table>
        <tbody>
          {data.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, colIndex) => (
                <td key={colIndex} data-testid={`cell-${rowIndex}-${colIndex}`}>
                  {String(cell?.value ?? '')}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  ),
}));

const renderWithRouter = (component: React.ReactElement) => {
  return render(
    <BrowserRouter>
      {component}
    </BrowserRouter>
  );
};

describe('TestSpreadsheet', () => {
  it('コンポーネントが正しくレンダリングされる', () => {
    renderWithRouter(<TestSpreadsheet />);
    
    expect(screen.getByText('Excel関数テストモード')).toBeInTheDocument();
    expect(screen.getByText('カテゴリを選択')).toBeInTheDocument();
    expect(screen.getByText('関数を選択')).toBeInTheDocument();
  });

  it('カテゴリを選択すると関数一覧が表示される', async () => {
    renderWithRouter(<TestSpreadsheet />);
    
    const categorySelect = screen.getByLabelText('カテゴリを選択');
    fireEvent.change(categorySelect, { target: { value: '02. 統計' } });
    
    const functionSelect = screen.getByLabelText('関数を選択');
    expect(functionSelect).not.toBeDisabled();
  });

  it('関数を選択するとスプレッドシートが表示される', async () => {
    renderWithRouter(<TestSpreadsheet />);
    
    // カテゴリを選択
    const categorySelect = screen.getByLabelText('カテゴリを選択');
    fireEvent.change(categorySelect, { target: { value: '02. 統計' } });
    
    // 関数を選択
    const functionSelect = screen.getByLabelText('関数を選択');
    fireEvent.change(functionSelect, { target: { value: 'AVERAGE' } });
    
    // スプレッドシートが表示されるまで待つ
    await waitFor(() => {
      expect(screen.getByTestId('spreadsheet')).toBeInTheDocument();
    });
  });

  it('統計関数GAMMALN(2)が正しく計算される', async () => {
    renderWithRouter(<TestSpreadsheet />);
    
    // カテゴリを選択
    const categorySelect = screen.getByLabelText('カテゴリを選択');
    fireEvent.change(categorySelect, { target: { value: '02. 統計' } });
    
    // GAMMALN関数を選択
    const functionSelect = screen.getByLabelText('関数を選択');
    fireEvent.change(functionSelect, { target: { value: 'GAMMALN' } });
    
    // 計算が完了するまで待つ
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });
    
    // B2セルの値を確認（GAMMALN(2) = 0）
    const b2Cell = screen.getByTestId('cell-1-1');
    expect(b2Cell.textContent).toBe('0');
  });

  it('テスト結果が正しく表示される', async () => {
    renderWithRouter(<TestSpreadsheet />);
    
    // カテゴリを選択
    const categorySelect = screen.getByLabelText('カテゴリを選択');
    fireEvent.change(categorySelect, { target: { value: '02. 統計' } });
    
    // AVERAGE関数を選択（期待値が設定されているもの）
    const functionSelect = screen.getByLabelText('関数を選択');
    fireEvent.change(functionSelect, { target: { value: 'AVERAGE' } });
    
    // テスト結果が表示されるまで待つ
    await waitFor(() => {
      expect(screen.getByText('テスト結果')).toBeInTheDocument();
    }, { timeout: 5000 });
    
    // 成功/失敗の表示を確認
    const successBadges = screen.getAllByText('成功');
    const failBadges = screen.queryAllByText('失敗');
    
    expect(successBadges.length + failBadges.length).toBeGreaterThan(0);
  });

  it('複数のカテゴリが正しく表示される', () => {
    renderWithRouter(<TestSpreadsheet />);
    
    const categorySelect = screen.getByLabelText('カテゴリを選択');
    const options = categorySelect.querySelectorAll('option');
    
    // デフォルトオプション + 実際のカテゴリ
    expect(options.length).toBeGreaterThan(5);
    
    // いくつかの主要なカテゴリが存在することを確認
    const categoryTexts = Array.from(options).map(opt => opt.textContent ?? '');
    expect(categoryTexts).toContain('01. 数学・三角');
    expect(categoryTexts).toContain('02. 統計');
    expect(categoryTexts).toContain('03. 文字列');
  });
});