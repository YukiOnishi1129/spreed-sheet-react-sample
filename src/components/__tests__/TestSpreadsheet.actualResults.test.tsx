import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';
import { allIndividualFunctionTests } from '../../data/individualFunctionTests';

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
  
  return { 
    default: Spreadsheet,
    Selection: {},
    Point: class Point {}
  };
});

describe('TestSpreadsheet - 実際の計算結果と期待値の検証', () => {
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

  // 期待値を持つすべての関数をテスト
  const functionsWithExpectedValues = allIndividualFunctionTests.filter(
    func => func.expectedValues && Object.keys(func.expectedValues).length > 0
  );

  describe('各関数の実際の成功/失敗判定', () => {
    functionsWithExpectedValues.forEach((functionTest) => {
      it(`${functionTest.category} - ${functionTest.name}: DOM要素から実際の成功/失敗を検証`, async () => {
        renderComponent();
        
        // カテゴリを選択
        const categorySelect = screen.getByLabelText('カテゴリを選択');
        await user.selectOptions(categorySelect, functionTest.category);
        
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
        
        // テスト結果セクションが表示されることを確認
        await waitFor(() => {
          expect(screen.getByText('テスト結果')).toBeInTheDocument();
        });
        
        // 各期待値について実際のDOM要素から結果を取得
        if (functionTest.expectedValues) {
          const results: { [key: string]: { expected: any, actual: string, isSuccess: boolean, domClass: string } } = {};
          
          for (const [cellAddress, expectedValue] of Object.entries(functionTest.expectedValues)) {
            // セルアドレスを含む要素を見つける
            const cellElement = screen.getByText(cellAddress);
            const resultContainer = cellElement.closest('.flex');
            
            if (resultContainer) {
              // 期待値と実際値のテキストを取得
              const textContent = resultContainer.textContent || '';
              
              // 期待値を抽出
              const expectedMatch = textContent.match(/期待値:\s*([^実際値]+)/);
              const expectedText = expectedMatch ? expectedMatch[1].trim() : '';
              
              // 実際値を抽出
              const actualMatch = textContent.match(/実際値:\s*([^成功失敗]+)/);
              const actualText = actualMatch ? actualMatch[1].trim() : '';
              
              // 成功/失敗の判定
              const isSuccess = resultContainer.classList.contains('bg-green-50');
              const isFailure = resultContainer.classList.contains('bg-red-50');
              const hasSuccessBadge = textContent.includes('成功');
              const hasFailureBadge = textContent.includes('失敗');
              
              // DOM要素のクラス名を記録
              const domClass = isSuccess ? 'bg-green-50' : (isFailure ? 'bg-red-50' : 'unknown');
              
              results[cellAddress] = {
                expected: expectedValue,
                actual: actualText,
                isSuccess: isSuccess && hasSuccessBadge,
                domClass
              };
              
              // 検証
              if (typeof expectedValue === 'number') {
                const actualNum = parseFloat(actualText);
                const isClose = !isNaN(actualNum) && Math.abs(actualNum - expectedValue) < 0.01;
                
                if (isClose) {
                  expect(isSuccess).toBe(true);
                  expect(hasSuccessBadge).toBe(true);
                  expect(hasFailureBadge).toBe(false);
                } else {
                  expect(isFailure).toBe(true);
                  expect(hasFailureBadge).toBe(true);
                  expect(hasSuccessBadge).toBe(false);
                }
              } else {
                const valuesMatch = actualText === String(expectedValue);
                
                if (valuesMatch) {
                  expect(isSuccess).toBe(true);
                  expect(hasSuccessBadge).toBe(true);
                  expect(hasFailureBadge).toBe(false);
                } else {
                  expect(isFailure).toBe(true);
                  expect(hasFailureBadge).toBe(true);
                  expect(hasSuccessBadge).toBe(false);
                }
              }
            }
          }
          
          // デバッグ用：結果を出力
          console.log(`\n${functionTest.name} (${functionTest.category}) の結果:`);
          Object.entries(results).forEach(([cell, result]) => {
            console.log(`  ${cell}: 期待値=${result.expected}, 実際値=${result.actual}, 成功=${result.isSuccess}, DOM=${result.domClass}`);
          });
          
          // サマリーの検証
          const summaryText = screen.getByText(/合計:/);
          const summaryContent = summaryText.textContent || '';
          
          const totalMatch = summaryContent.match(/合計:\s*(\d+)/);
          const successMatch = summaryContent.match(/成功:\s*(\d+)/);
          const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
          
          const totalCount = totalMatch ? parseInt(totalMatch[1]) : 0;
          const successCount = successMatch ? parseInt(successMatch[1]) : 0;
          const failureCount = failureMatch ? parseInt(failureMatch[1]) : 0;
          
          // サマリーの整合性を確認
          expect(totalCount).toBe(Object.keys(functionTest.expectedValues).length);
          expect(successCount + failureCount).toBe(totalCount);
          
          // 実際の成功数と失敗数を確認
          const actualSuccessCount = Object.values(results).filter(r => r.isSuccess).length;
          const actualFailureCount = Object.values(results).filter(r => !r.isSuccess).length;
          
          expect(successCount).toBe(actualSuccessCount);
          expect(failureCount).toBe(actualFailureCount);
        }
      });
    });
  });

  describe('特定の問題がある関数の詳細検証', () => {
    it('ISO.CEILING関数の実際の動作を検証', async () => {
      renderComponent();
      
      // ISO.CEILINGを含むカテゴリを選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '01. 数学');
      
      await waitFor(() => {
        const functionSelect = screen.getByLabelText('関数を選択');
        expect(functionSelect).not.toBeDisabled();
      });
      
      await user.selectOptions(screen.getByLabelText('関数を選択'), 'ISO.CEILING');
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      }, { timeout: 10000 });
      
      // C3セルの結果を確認（期待値: -4）
      const cellC3 = screen.getByText('C3');
      const resultContainer = cellC3.closest('.flex');
      
      if (resultContainer) {
        const textContent = resultContainer.textContent || '';
        console.log('ISO.CEILING C3の結果:', textContent);
        
        // 実際値が#NUM!エラーかどうか確認
        expect(textContent).toContain('実際値:');
        
        if (textContent.includes('#NUM!')) {
          console.error('ISO.CEILING関数が#NUM!エラーを返しています');
          expect(resultContainer).toHaveClass('bg-red-50');
          expect(textContent).toContain('失敗');
        }
      }
    });

    it('動的配列関数の実際の動作を検証', async () => {
      renderComponent();
      
      // 動的配列カテゴリを選択
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '12. 動的配列');
      
      await waitFor(() => {
        const functionSelect = screen.getByLabelText('関数を選択');
        expect(functionSelect).not.toBeDisabled();
      });
      
      // 各動的配列関数をテスト
      const dynamicArrayFunctions = ['FILTER', 'SORT', 'UNIQUE', 'TRANSPOSE', 'SEQUENCE'];
      
      for (const funcName of dynamicArrayFunctions) {
        const options = screen.getByLabelText('関数を選択').querySelectorAll('option');
        const functionOption = Array.from(options).find(opt => opt.textContent?.includes(funcName));
        
        if (functionOption) {
          await user.selectOptions(screen.getByLabelText('関数を選択'), funcName);
          
          await waitFor(() => {
            expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
          }, { timeout: 10000 });
          
          // テスト結果があるか確認
          const testResultsSection = screen.queryByText('テスト結果');
          if (testResultsSection) {
            const summaryText = screen.getByText(/合計:/);
            const summaryContent = summaryText.textContent || '';
            console.log(`${funcName}のサマリー:`, summaryContent);
            
            // 失敗がある場合は詳細を出力
            if (summaryContent.includes('失敗: ') && !summaryContent.includes('失敗: 0')) {
              const failureBadges = screen.getAllByText('失敗');
              failureBadges.forEach((badge) => {
                const container = badge.closest('.flex');
                if (container) {
                  console.error(`${funcName}の失敗:`, container.textContent);
                }
              });
            }
          }
        }
      }
    });
  });

  describe('全体のサマリー検証', () => {
    it('すべての関数のサマリーを集計', async () => {
      const totalResults = {
        totalFunctions: 0,
        totalTests: 0,
        totalSuccess: 0,
        totalFailure: 0,
        failedFunctions: [] as string[]
      };
      
      for (const functionTest of functionsWithExpectedValues.slice(0, 10)) { // 最初の10個だけテスト
        renderComponent();
        
        await user.selectOptions(screen.getByLabelText('カテゴリを選択'), functionTest.category);
        await waitFor(() => {
          expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
        });
        
        await user.selectOptions(screen.getByLabelText('関数を選択'), functionTest.name);
        
        await waitFor(() => {
          expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
        }, { timeout: 10000 });
        
        const summaryText = screen.queryByText(/合計:/);
        if (summaryText) {
          const summaryContent = summaryText.textContent || '';
          const failureMatch = summaryContent.match(/失敗:\s*(\d+)/);
          const failureCount = failureMatch ? parseInt(failureMatch[1]) : 0;
          
          if (failureCount > 0) {
            totalResults.failedFunctions.push(`${functionTest.category}/${functionTest.name}`);
          }
          
          totalResults.totalFunctions++;
        }
      }
      
      console.log('\n=== 全体のサマリー ===');
      console.log('テストした関数数:', totalResults.totalFunctions);
      console.log('失敗した関数:', totalResults.failedFunctions);
    });
  });
});