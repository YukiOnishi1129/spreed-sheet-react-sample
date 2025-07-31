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
  
  const Spreadsheet = ({ data }: { data: CellType[][] }) => {
    return (
      <div data-testid="spreadsheet-mock">
        <table>
          <tbody>
            {data.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, colIndex) => {
                  const cellAddress = `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`;
                  return (
                    <td key={colIndex}>
                      <div
                        data-testid={`cell-${cellAddress}`}
                        data-value={cell?.value?.toString() ?? ''}
                        data-formula={cell?.formula ?? ''}
                      >
                        {cell?.value?.toString() ?? ''}
                      </div>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };
  
  return { default: Spreadsheet };
});

describe('TestSpreadsheet - Expected Values 検証テスト', () => {
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

  // expectedValuesが定義されている関数のみをフィルタリング
  const functionsWithExpectedValues = allIndividualFunctionTests.filter(
    func => func.expectedValues && Object.keys(func.expectedValues).length > 0
  );

  describe('全関数のExpected Values検証', () => {
    // カテゴリごとにグループ化
    const categoriesWithExpected = Array.from(
      new Set(functionsWithExpectedValues.map(f => f.category))
    ).sort();

    categoriesWithExpected.forEach(category => {
      describe(`${category} カテゴリ`, () => {
        const functionsInCategory = functionsWithExpectedValues.filter(
          f => f.category === category
        );

        functionsInCategory.forEach((functionTest: IndividualFunctionTest) => {
          it(`${functionTest.name}: ${functionTest.description} - 期待値検証`, async () => {
            renderComponent();

            // カテゴリを選択
            const categorySelect = screen.getByLabelText('カテゴリを選択');
            await user.selectOptions(categorySelect, category);

            // 関数を選択
            await waitFor(() => {
              const functionSelect = screen.getByLabelText('関数を選択');
              expect(functionSelect).not.toBeDisabled();
            });
            
            const functionSelect = screen.getByLabelText('関数を選択');
            await user.selectOptions(functionSelect, functionTest.name);

            // 計算完了を待つ
            await waitFor(() => {
              expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
            }, { timeout: 10000 });

            // テスト結果セクションの存在を確認
            await waitFor(() => {
              expect(screen.getByText('テスト結果')).toBeInTheDocument();
            });

            // 各期待値を検証
            if (functionTest.expectedValues) {
              for (const [cellAddress, expectedValue] of Object.entries(functionTest.expectedValues)) {
                // セルアドレスが表示されていることを確認
                const cellElement = screen.getByText(cellAddress);
                expect(cellElement).toBeInTheDocument();

                // 期待値と実際値の表示を確認
                const resultContainer = cellElement.closest('.flex');
                expect(resultContainer).toBeTruthy();

                // 期待値の表示
                expect(resultContainer).toHaveTextContent(`期待値: ${JSON.stringify(expectedValue)}`);

                // 実際の計算値を取得
                const cellDataElement = screen.getByTestId(`cell-${cellAddress}`);
                const actualValue = cellDataElement.getAttribute('data-value');

                // 数値の場合は近似値で比較、文字列の場合は完全一致
                if (typeof expectedValue === 'number' && actualValue) {
                  const actualNum = parseFloat(actualValue);
                  expect(actualNum).toBeCloseTo(expectedValue, 2);
                  
                  // 成功表示（緑色）を確認
                  expect(resultContainer).toHaveClass('bg-green-50');
                  expect(resultContainer).toHaveTextContent('成功');
                } else {
                  // 文字列やブール値の場合
                  expect(actualValue).toBe(String(expectedValue));
                  
                  // 成功表示を確認
                  expect(resultContainer).toHaveClass('bg-green-50');
                  expect(resultContainer).toHaveTextContent('成功');
                }
              }

              // 全体の成功数を確認
              const summaryText = screen.getByText(/合計:/);
              const totalExpected = Object.keys(functionTest.expectedValues).length;
              expect(summaryText).toHaveTextContent(`合計: ${totalExpected}`);
              expect(summaryText).toHaveTextContent(`成功: ${totalExpected}`);
              expect(summaryText).toHaveTextContent('失敗: 0');
            }
          });
        });
      });
    });
  });

  describe('特定の関数の詳細検証', () => {
    it('SUM関数: 複数セルの合計が正しく計算される', async () => {
      renderComponent();

      const sumTest = functionsWithExpectedValues.find(f => f.name === 'SUM');
      if (!sumTest) {
        console.warn('SUM関数のテストデータが見つかりません');
        return;
      }

      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), sumTest.category);
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'SUM');

      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });

      // E2セルの値が100であることを確認（expectedValues: { 'E2': 100 }）
      const cellE2 = screen.getByTestId('cell-E2');
      expect(cellE2.getAttribute('data-value')).toBe('100');
    });

    it('AVERAGE関数: 平均値が正しく計算される', async () => {
      renderComponent();

      const avgTest = functionsWithExpectedValues.find(f => f.name === 'AVERAGE');
      if (!avgTest) {
        console.warn('AVERAGE関数のテストデータが見つかりません');
        return;
      }

      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), avgTest.category);
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'AVERAGE');

      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });

      // 期待値を確認
      if (avgTest.expectedValues) {
        for (const [cell, expected] of Object.entries(avgTest.expectedValues)) {
          const cellElement = screen.getByTestId(`cell-${cell}`);
          const actualValue = parseFloat(cellElement.getAttribute('data-value') ?? '0');
          expect(actualValue).toBeCloseTo(expected as number, 2);
        }
      }
    });

    it('IF関数: 条件分岐が正しく動作する', async () => {
      renderComponent();

      const ifTest = functionsWithExpectedValues.find(f => f.name === 'IF');
      if (!ifTest) {
        console.warn('IF関数のテストデータが見つかりません');
        return;
      }

      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), ifTest.category);
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'IF');

      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });

      // 期待値を確認
      if (ifTest.expectedValues) {
        for (const [cell, expected] of Object.entries(ifTest.expectedValues)) {
          const cellElement = screen.getByTestId(`cell-${cell}`);
          const actualValue = cellElement.getAttribute('data-value');
          expect(actualValue).toBe(String(expected));
        }
      }
    });

    it('VLOOKUP関数: 検索が正しく動作する', async () => {
      renderComponent();

      const vlookupTest = functionsWithExpectedValues.find(f => f.name === 'VLOOKUP');
      if (!vlookupTest) {
        console.warn('VLOOKUP関数のテストデータが見つかりません');
        return;
      }

      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), vlookupTest.category);
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'VLOOKUP');

      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      });

      // 期待値を確認
      if (vlookupTest.expectedValues) {
        for (const [cell, expected] of Object.entries(vlookupTest.expectedValues)) {
          const cellElement = screen.getByTestId(`cell-${cell}`);
          const actualValue = cellElement.getAttribute('data-value');
          expect(actualValue).toBe(String(expected));
        }
      }
    });
  });

  describe('エラーケースの検証', () => {
    it('DIV/0エラーが正しく表示される', async () => {
      renderComponent();

      // エラー処理関数（IFERROR等）を探す
      const errorTest = functionsWithExpectedValues.find(
        f => f.name === 'IFERROR' || f.expectedValues && 
             Object.values(f.expectedValues).some(v => String(v).includes('#'))
      );

      if (errorTest) {
        await user.selectOptions(screen.getByLabelText('カテゴリを選択'), errorTest.category);
        await waitFor(() => {
          expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
        });
        await user.selectOptions(screen.getByLabelText('関数を選択'), errorTest.name);

        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        });

        // エラー値の確認
        if (errorTest.expectedValues) {
          for (const [cell, expected] of Object.entries(errorTest.expectedValues)) {
            if (String(expected).includes('#')) {
              const cellElement = screen.getByTestId(`cell-${cell}`);
              const actualValue = cellElement.getAttribute('data-value');
              expect(actualValue).toContain('#');
            }
          }
        }
      }
    });
  });

  describe('複雑な関数の検証', () => {
    it('配列関数が正しく動作する', async () => {
      renderComponent();

      const arrayTests = functionsWithExpectedValues.filter(
        f => f.category.includes('配列') || f.category.includes('行列')
      );

      for (const arrayTest of arrayTests.slice(0, 3)) { // 最初の3つだけテスト
        await user.selectOptions(screen.getByLabelText('カテゴリを選択'), arrayTest.category);
        await waitFor(() => {
          expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
        });
        await user.selectOptions(screen.getByLabelText('関数を選択'), arrayTest.name);

        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        }, { timeout: 10000 });

        if (arrayTest.expectedValues) {
          for (const [cell, expected] of Object.entries(arrayTest.expectedValues)) {
            const cellElement = screen.getByTestId(`cell-${cell}`);
            const actualValue = cellElement.getAttribute('data-value');
            
            if (typeof expected === 'number') {
              expect(parseFloat(actualValue ?? '0')).toBeCloseTo(expected, 2);
            } else {
              expect(actualValue).toBe(String(expected));
            }
          }
        }
      }
    });

    it('統計関数が正しく動作する', async () => {
      renderComponent();

      const statsTests = functionsWithExpectedValues.filter(
        f => f.category.includes('統計')
      );

      for (const statsTest of statsTests.slice(0, 3)) { // 最初の3つだけテスト
        await user.selectOptions(screen.getByLabelText('カテゴリを選択'), statsTest.category);
        await waitFor(() => {
          expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
        });
        await user.selectOptions(screen.getByLabelText('関数を選択'), statsTest.name);

        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        });

        if (statsTest.expectedValues) {
          const testResultsSection = screen.getByText('テスト結果');
          expect(testResultsSection).toBeInTheDocument();

          // 成功数の確認
          const summaryText = screen.getByText(/合計:/);
          const totalExpected = Object.keys(statsTest.expectedValues).length;
          expect(summaryText).toHaveTextContent(`成功: ${totalExpected}`);
        }
      }
    });
  });

  describe('サマリー検証', () => {
    it('expectedValuesを持つ全関数の概要', () => {
      console.log('\n=== Expected Values を持つ関数の概要 ===');
      
      const categorySummary: { [key: string]: number } = {};
      functionsWithExpectedValues.forEach(func => {
        categorySummary[func.category] = (categorySummary[func.category] || 0) + 1;
      });

      console.log(`総関数数: ${functionsWithExpectedValues.length}`);
      console.log('\nカテゴリ別:');
      Object.entries(categorySummary).forEach(([category, count]) => {
        console.log(`  ${category}: ${count}関数`);
      });

      console.log('\n期待値の総数:');
      let totalExpectedValues = 0;
      functionsWithExpectedValues.forEach(func => {
        if (func.expectedValues) {
          totalExpectedValues += Object.keys(func.expectedValues).length;
        }
      });
      console.log(`  ${totalExpectedValues}個の期待値`);
      
      expect(functionsWithExpectedValues.length).toBeGreaterThan(0);
      expect(totalExpectedValues).toBeGreaterThan(0);
    });
  });
});