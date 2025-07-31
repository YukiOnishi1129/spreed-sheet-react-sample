import { describe, it, expect } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';
import { allIndividualFunctionTests } from '../../data/individualFunctionTests';

describe('実際の失敗を正確に検出', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('DATEDIFの実際の失敗を確認', async () => {
    const user = userEvent.setup();
    const { container } = renderComponent();
    
    // DATEDIFを探す
    const datedifTest = allIndividualFunctionTests.find(f => f.name === 'DATEDIF');
    
    if (!datedifTest) {
      throw new Error('DATEDIF test not found');
    }
    
    // カテゴリと関数を選択
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), datedifTest.category);
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'DATEDIF');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 10000 });
    
    // テスト結果セクションを探す
    const testResultsSection = container.querySelector('div:has(> h3:contains("テスト結果"))');
    
    if (testResultsSection) {
      // bg-red-50クラス（失敗）を持つ要素を探す
      const failedElements = testResultsSection.querySelectorAll('.bg-red-50');
      console.log(`\nDATEDIF: ${failedElements.length}個の失敗が見つかりました`);
      
      // 各失敗の詳細を出力
      failedElements.forEach((element) => {
        const text = element.textContent || '';
        console.log('失敗の詳細:', text);
      });
      
      // D2セルの結果を特に確認
      const d2Result = Array.from(testResultsSection.querySelectorAll('.font-mono')).find(
        el => el.textContent === 'D2'
      );
      
      if (d2Result) {
        const resultContainer = d2Result.closest('.flex');
        console.log('\nD2の結果:', resultContainer?.textContent);
      }
    }
    
    // 失敗バッジを数える（より正確な方法）
    const failureBadges = container.querySelectorAll('.bg-red-100.text-red-800');
    console.log(`\n失敗バッジ数: ${failureBadges.length}`);
    
    expect(failureBadges.length).toBeGreaterThan(0);
  });
  
  it('各カテゴリの実際の失敗を確認', async () => {
    const user = userEvent.setup();
    
    const testFunctions = [
      { category: '04. 日付', name: 'DATEDIF' },
      { category: '01. 数学', name: 'ISO.CEILING' },
      { category: '12. 動的配列', name: 'FILTER' },
      { category: '07. 財務', name: 'PMT' }
    ];
    
    for (const { category, name } of testFunctions) {
      const { container, unmount } = renderComponent();
      
      await user.selectOptions(screen.getByLabelText('カテゴリを選択'), category);
      await waitFor(() => {
        expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
      });
      
      await user.selectOptions(screen.getByLabelText('関数を選択'), name);
      
      await waitFor(() => {
        expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
      }, { timeout: 5000 });
      
      // 失敗要素を正確にカウント
      const failedCells = container.querySelectorAll('.bg-red-50');
      const failureBadges = container.querySelectorAll('.bg-red-100.text-red-800');
      
      console.log(`\n${category} - ${name}:`);
      console.log(`  失敗セル数: ${failedCells.length}`);
      console.log(`  失敗バッジ数: ${failureBadges.length}`);
      
      // 失敗の詳細を出力
      if (failedCells.length > 0) {
        const firstFailure = failedCells[0];
        console.log(`  最初の失敗: ${firstFailure.textContent}`);
      }
      
      unmount();
    }
  });
});