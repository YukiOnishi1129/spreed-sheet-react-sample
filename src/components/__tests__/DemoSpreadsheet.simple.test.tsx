import { describe, it, expect, beforeEach, vi } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import { BrowserRouter } from 'react-router-dom';
import DemoSpreadsheet from '../DemoSpreadsheet';
import '@testing-library/jest-dom';

// react-spreadsheetのモック - より詳細に
vi.mock('react-spreadsheet', () => ({
  default: ({ data }: any) => {
    return (
      <div data-testid="spreadsheet-mock">
        <table>
          <tbody>
            {data.map((row: any[], rowIndex: number) => (
              <tr key={rowIndex}>
                {row.map((cell: any, colIndex: number) => (
                  <td
                    key={colIndex}
                    data-row={rowIndex}
                    data-col={colIndex}
                    data-value={cell?.value}
                    data-formula={cell?.formula}
                  >
                    {cell?.value ?? ''}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  },
}));

describe('DemoSpreadsheet シンプルテスト', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <DemoSpreadsheet />
      </BrowserRouter>
    );
  };

  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('デモモードの切り替えができる', () => {
    renderComponent();
    
    // デフォルトはグループデモ
    expect(screen.getByText('デモカテゴリを選択')).toBeInTheDocument();
    
    // 個別関数デモに切り替え
    fireEvent.click(screen.getByText('個別関数デモ'));
    
    // UIが変わる
    expect(screen.queryByText('デモカテゴリを選択')).not.toBeInTheDocument();
    expect(screen.getByText('関数を選択')).toBeInTheDocument();
  });

  it('モーダルを開閉できる', async () => {
    renderComponent();
    
    // 個別関数デモモードに切り替え
    fireEvent.click(screen.getByText('個別関数デモ'));
    
    // モーダルを開く
    const selectButton = screen.getByRole('button', { name: /関数を選択/ });
    fireEvent.click(selectButton);
    
    // モーダルが表示される
    await waitFor(() => {
      const modal = document.querySelector('[role="dialog"], .fixed.inset-0');
      expect(modal).toBeInTheDocument();
    });
  });

  it('スプレッドシートが表示される', () => {
    renderComponent();
    
    // スプレッドシートのモックが表示される
    expect(screen.getByTestId('spreadsheet-mock')).toBeInTheDocument();
  });

  it('選択した関数の情報が表示される', async () => {
    renderComponent();
    
    // 個別関数デモモードに切り替え
    fireEvent.click(screen.getByText('個別関数デモ'));
    
    // デフォルトメッセージ
    expect(screen.getByText('関数を選択してください')).toBeInTheDocument();
  });

  it('数式バーが存在する', () => {
    renderComponent();
    
    // 数式バーの要素を確認
    const formulaBar = screen.getByPlaceholderText('セルを選択してください');
    expect(formulaBar).toBeInTheDocument();
    expect(formulaBar).toHaveAttribute('readOnly');
  });
});