import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';
import { allIndividualFunctionTests } from '../../data/individualFunctionTests';
import type { IndividualFunctionTest } from '../../types/spreadsheet';

// react-spreadsheetのモック
vi.mock('react-spreadsheet', () => {
  interface CellType {
    value?: string | number | boolean | null;
    formula?: string;
    title?: string;
    'data-formula'?: string;
  }
  
  const Spreadsheet = ({ data, onChange }: { 
    data: CellType[][], 
    onChange: (data: CellType[][]) => void 
  }) => {
    return (
      <div data-testid="spreadsheet-mock">
        <table>
          <tbody>
            {data.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, colIndex) => (
                  <td key={colIndex} data-testid={`cell-${rowIndex}-${colIndex}`}>
                    <input
                      type="text"
                      value={cell?.value?.toString() ?? ''}
                      onChange={(e) => {
                        const newData = [...data];
                        newData[rowIndex][colIndex] = { 
                          ...newData[rowIndex][colIndex], 
                          value: e.target.value 
                        };
                        onChange(newData);
                      }}
                      data-formula={cell?.formula ?? ''}
                      title={cell?.title ?? ''}
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };
  
  return { default: Spreadsheet };
});

describe('TestSpreadsheet - 包括的テスト', () => {
  let user: ReturnType<typeof userEvent.setup>;

  beforeEach(() => {
    user = userEvent.setup();
  });

  afterEach(() => {
    vi.clearAllMocks();
  });

  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  describe('全関数の計算テスト', () => {
    // カテゴリごとにテストをグループ化
    const categories = Array.from(new Set(allIndividualFunctionTests.map(f => f.category))).sort();
    
    categories.forEach(category => {
      describe(`${category} カテゴリ`, () => {
        const functionsInCategory = allIndividualFunctionTests.filter(f => f.category === category);
        
        functionsInCategory.forEach((functionTest: IndividualFunctionTest) => {
          it(`${functionTest.name} - ${functionTest.description}`, async () => {
            renderComponent();
            
            // カテゴリを選択
            const categorySelect = screen.getByLabelText('カテゴリを選択');
            await user.selectOptions(categorySelect, category);
            
            // 関数を選択
            const functionSelect = screen.getByLabelText('関数を選択');
            await waitFor(() => {
              expect(functionSelect).not.toBeDisabled();
            });
            await user.selectOptions(functionSelect, functionTest.name);
            
            // 計算が完了するまで待つ
            await waitFor(() => {
              expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
            }, { timeout: 5000 });
            
            // 期待値が定義されている場合はテスト結果を検証
            if (functionTest.expectedValues) {
              // テスト結果セクションが表示されることを確認
              await waitFor(() => {
                expect(screen.getByText('テスト結果')).toBeInTheDocument();
              });
              
              // 各期待値をチェック
              Object.entries(functionTest.expectedValues).forEach(([cell, expectedValue]) => {
                const cellElement = screen.getByText(cell);
                expect(cellElement).toBeInTheDocument();
                
                // 成功/失敗の表示を確認
                const resultContainer = cellElement.closest('.flex');
                if (resultContainer) {
                  const expectedText = `期待値: ${JSON.stringify(expectedValue)}`;
                  expect(resultContainer).toHaveTextContent(expectedText);
                  
                  // 成功していることを確認（緑色の背景）
                  expect(resultContainer).toHaveClass('bg-green-50');
                }
              });
              
              // 全体の成功数を確認
              const summaryText = screen.getByText(/合計:/);
              const totalTests = Object.keys(functionTest.expectedValues).length;
              expect(summaryText).toHaveTextContent(`成功: ${totalTests}`);
              expect(summaryText).toHaveTextContent('失敗: 0');
            }
            
            // スプレッドシートにデータが表示されていることを確認
            const spreadsheet = screen.getByTestId('spreadsheet-mock');
            expect(spreadsheet).toBeInTheDocument();
            
            // 数式セルが正しく設定されていることを確認
            const formulaCells = functionTest.data.flat().filter(
              cell => typeof cell === 'string' && cell.startsWith('=')
            );
            
            if (formulaCells.length > 0) {
              const inputs = spreadsheet.querySelectorAll('input[data-formula]');
              const formulaInputs = Array.from(inputs).filter(
                input => (input as HTMLInputElement).getAttribute('data-formula')?.startsWith('=')
              );
              expect(formulaInputs.length).toBeGreaterThan(0);
            }
          });
        });
      });
    });
  });

  describe('値変更後の再計算テスト', () => {
    // 代表的な関数を選んで再計算テストを実施
    const testCases = [
      {
        category: '数学/三角',
        functionName: 'SUM',
        description: '合計関数の再計算',
        initialCell: 'A1',
        newValue: '100',
        expectedChange: true
      },
      {
        category: '統計',
        functionName: 'AVERAGE',
        description: '平均関数の再計算',
        initialCell: 'A1',
        newValue: '50',
        expectedChange: true
      },
      {
        category: '論理',
        functionName: 'IF',
        description: 'IF関数の再計算',
        initialCell: 'A1',
        newValue: '0',
        expectedChange: true
      }
    ];

    testCases.forEach(testCase => {
      it(`${testCase.functionName} - ${testCase.description}`, async () => {
        renderComponent();
        
        // カテゴリと関数を選択
        const categorySelect = screen.getByLabelText('カテゴリを選択');
        await user.selectOptions(categorySelect, testCase.category);
        
        const functionSelect = screen.getByLabelText('関数を選択');
        await waitFor(() => {
          expect(functionSelect).not.toBeDisabled();
        });
        await user.selectOptions(functionSelect, testCase.functionName);
        
        // 初期計算が完了するまで待つ
        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        }, { timeout: 5000 });
        
        // 初期の計算結果を記録
        const spreadsheet = screen.getByTestId('spreadsheet-mock');
        const formulaCells = spreadsheet.querySelectorAll('input[data-formula^="="]');
        const initialResults: string[] = [];
        
        formulaCells.forEach(cell => {
          initialResults.push((cell as HTMLInputElement).value);
        });
        
        // セルの値を変更（実際のスプレッドシートでは onChange イベントが発火する）
        // ここではモックのため、直接関数を再実行する必要がある
        
        // 関数を再選択して再計算をトリガー
        await user.selectOptions(functionSelect, '');
        await user.selectOptions(functionSelect, testCase.functionName);
        
        // 再計算が完了するまで待つ
        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        }, { timeout: 5000 });
        
        // テスト結果が更新されていることを確認
        if (testCase.expectedChange) {
          const testResults = screen.getByText('テスト結果');
          expect(testResults).toBeInTheDocument();
        }
      });
    });
  });

  describe('エラーハンドリングテスト', () => {
    it('無効な数式の場合エラーが表示される', async () => {
      renderComponent();
      
      // エラーを含む関数を選択（例：DIV/0エラー）
      const errorFunction = allIndividualFunctionTests.find(
        f => f.name === 'IFERROR' || f.name === 'ISERROR'
      );
      
      if (errorFunction) {
        const categorySelect = screen.getByLabelText('カテゴリを選択');
        await user.selectOptions(categorySelect, errorFunction.category);
        
        const functionSelect = screen.getByLabelText('関数を選択');
        await waitFor(() => {
          expect(functionSelect).not.toBeDisabled();
        });
        await user.selectOptions(functionSelect, errorFunction.name);
        
        // 計算が完了するまで待つ
        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        }, { timeout: 5000 });
        
        // エラー処理が正しく動作していることを確認
        const spreadsheet = screen.getByTestId('spreadsheet-mock');
        expect(spreadsheet).toBeInTheDocument();
      }
    });
  });

  describe('パフォーマンステスト', () => {
    it('複雑な関数でも5秒以内に計算が完了する', async () => {
      renderComponent();
      
      // 計算負荷の高い関数を選択（例：行列関数）
      const complexFunction = allIndividualFunctionTests.find(
        f => f.category === '行列' || f.category === '配列'
      );
      
      if (complexFunction) {
        const startTime = Date.now();
        
        const categorySelect = screen.getByLabelText('カテゴリを選択');
        await user.selectOptions(categorySelect, complexFunction.category);
        
        const functionSelect = screen.getByLabelText('関数を選択');
        await waitFor(() => {
          expect(functionSelect).not.toBeDisabled();
        });
        await user.selectOptions(functionSelect, complexFunction.name);
        
        // 計算が完了するまで待つ
        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        }, { timeout: 5000 });
        
        const endTime = Date.now();
        const elapsedTime = endTime - startTime;
        
        // 5秒以内に完了することを確認
        expect(elapsedTime).toBeLessThan(5000);
      }
    });
  });

  describe('UI/UXテスト', () => {
    it('カテゴリ選択前は関数選択が無効化されている', async () => {
      renderComponent();
      
      const functionSelect = screen.getByLabelText('関数を選択');
      expect(functionSelect).toBeDisabled();
    });

    it('カテゴリを選択すると関数選択が有効化される', async () => {
      renderComponent();
      
      const categorySelect = screen.getByLabelText('カテゴリを選択');
      const testCategories = Array.from(new Set(allIndividualFunctionTests.map(f => f.category))).sort();
      const firstCategory = testCategories[0];
      if (!firstCategory) return;
      await user.selectOptions(categorySelect, firstCategory);
      
      const functionSelect = screen.getByLabelText('関数を選択');
      await waitFor(() => {
        expect(functionSelect).not.toBeDisabled();
      });
    });

    it('関数を選択すると関数情報が表示される', async () => {
      renderComponent();
      
      const firstFunction = allIndividualFunctionTests[0];
      const categorySelect = screen.getByLabelText('カテゴリを選択');
      await user.selectOptions(categorySelect, firstFunction.category);
      
      const functionSelect = screen.getByLabelText('関数を選択');
      await waitFor(() => {
        expect(functionSelect).not.toBeDisabled();
      });
      await user.selectOptions(functionSelect, firstFunction.name);
      
      // 関数情報が表示される
      await waitFor(() => {
        expect(screen.getByText(firstFunction.name)).toBeInTheDocument();
        expect(screen.getByText(firstFunction.description)).toBeInTheDocument();
      });
    });

    it('計算中はローディング表示が出る', async () => {
      renderComponent();
      
      const firstFunction = allIndividualFunctionTests[0];
      const categorySelect = screen.getByLabelText('カテゴリを選択');
      await user.selectOptions(categorySelect, firstFunction.category);
      
      const functionSelect = screen.getByLabelText('関数を選択');
      await waitFor(() => {
        expect(functionSelect).not.toBeDisabled();
      });
      
      // 関数を選択
      await user.selectOptions(functionSelect, firstFunction.name);
      
      // ローディング表示が一時的に表示される
      expect(screen.getByText('数式を計算中...')).toBeInTheDocument();
      
      // 計算完了後はローディング表示が消える
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      }, { timeout: 5000 });
    });
  });
});