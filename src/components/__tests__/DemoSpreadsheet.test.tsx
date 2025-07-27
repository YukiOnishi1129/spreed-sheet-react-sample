import { describe, it, expect, beforeEach, vi } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import { BrowserRouter } from 'react-router-dom';
import DemoSpreadsheet from '../DemoSpreadsheet';
import '@testing-library/jest-dom';

// react-spreadsheetのモック
vi.mock('react-spreadsheet', () => ({
  default: ({ data }: any) => {
    return (
      <div data-testid="spreadsheet">
        {data.map((row: any[], rowIndex: number) => (
          <div key={rowIndex} data-testid={`row-${rowIndex}`}>
            {row.map((cell: any, colIndex: number) => (
              <div
                key={colIndex}
                data-testid={`cell-${rowIndex}-${colIndex}`}
                data-value={cell?.value}
                data-formula={cell?.formula}
              >
                {cell?.value ?? ''}
              </div>
            ))}
          </div>
        ))}
      </div>
    );
  },
}));

describe('DemoSpreadsheet - 個別関数テスト', () => {
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

  it('個別関数デモモードに切り替えられる', async () => {
    renderComponent();
    
    const individualButton = screen.getByText('個別関数デモ');
    fireEvent.click(individualButton);
    
    expect(screen.getByText('関数を選択')).toBeInTheDocument();
  });

  it('モーダルを開いて関数を選択できる', async () => {
    renderComponent();
    
    // 個別関数デモモードに切り替え
    fireEvent.click(screen.getByText('個別関数デモ'));
    
    // モーダルを開く
    fireEvent.click(screen.getByText('関数を選択'));
    
    // モーダルが表示される
    await waitFor(() => {
      expect(screen.getByText('関数を選択', { selector: 'h2' })).toBeInTheDocument();
    });
    
    // SUM関数を選択（完全一致）
    const sumButtons = screen.getAllByRole('button');
    const sumButton = sumButtons.find(button => {
      const text = button.textContent || '';
      return text.includes('SUM') && text.includes('数値の合計を計算');
    });
    
    if (sumButton) {
      fireEvent.click(sumButton);
    }
    
    // 関数が選択される
    await waitFor(() => {
      expect(screen.getByText('選択中: SUM')).toBeInTheDocument();
    });
  });

  // 主要な関数のテスト
  describe('関数の計算結果確認', () => {
    const testCases = [
      {
        functionName: 'SUM',
        expectedValues: { 'E2': 100 },
        cellToCheck: { row: 1, col: 4 }
      },
      {
        functionName: 'AVERAGE',
        expectedValues: { 'E2': 25 },
        cellToCheck: { row: 1, col: 4 }
      },
      {
        functionName: 'MAX',
        expectedValues: { 'E2': 40 },
        cellToCheck: { row: 1, col: 4 }
      },
      {
        functionName: 'MIN',
        expectedValues: { 'E2': 10 },
        cellToCheck: { row: 1, col: 4 }
      },
    ];

    testCases.forEach(({ functionName, expectedValues, cellToCheck }) => {
      it(`${functionName}関数の計算結果が正しい`, async () => {
        renderComponent();
        
        // 個別関数デモモードに切り替え
        fireEvent.click(screen.getByText('個別関数デモ'));
        
        // モーダルを開く
        fireEvent.click(screen.getByText('関数を選択'));
        
        // 関数を選択
        await waitFor(() => {
          const functionButton = screen.getByRole('button', { name: new RegExp(functionName) });
          fireEvent.click(functionButton);
        });
        
        // スプレッドシートのデータが更新されるのを待つ
        await waitFor(() => {
          const cell = screen.getByTestId(`cell-${cellToCheck.row}-${cellToCheck.col}`);
          const cellValue = cell.getAttribute('data-value');
          const expectedValue = Object.values(expectedValues)[0];
          
          expect(Number(cellValue)).toBe(expectedValue);
        }, { timeout: 5000 });
      });
    });
  });

  it('カテゴリーで関数をフィルタリングできる', async () => {
    renderComponent();
    
    // 個別関数デモモードに切り替え
    fireEvent.click(screen.getByText('個別関数デモ'));
    
    // モーダルを開く
    fireEvent.click(screen.getByText('関数を選択'));
    
    // 統計カテゴリーを選択
    await waitFor(() => {
      const statisticsButton = screen.getByRole('button', { name: /02\. 統計/ });
      fireEvent.click(statisticsButton);
    });
    
    // 統計関数のみが表示される
    expect(screen.getByRole('button', { name: /AVERAGE/ })).toBeInTheDocument();
    expect(screen.queryByRole('button', { name: /CONCATENATE/ })).not.toBeInTheDocument();
  });
});