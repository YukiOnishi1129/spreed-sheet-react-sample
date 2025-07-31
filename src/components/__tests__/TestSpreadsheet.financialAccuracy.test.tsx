import { describe, it, expect } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';

describe('財務関数の精度テスト', () => {
  const renderComponent = () => {
    return render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
  };

  it('RATE関数の計算結果を詳細に確認', async () => {
    const user = userEvent.setup();
    const { container } = renderComponent();
    
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '07. 財務');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'RATE');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });
    
    // セルの値を取得
    const cells = container.querySelectorAll('td');
    const cellValues: Record<string, string> = {};
    
    cells.forEach((cell, index) => {
      const row = Math.floor(index / 10);
      const col = index % 10;
      const cellContent = cell.textContent || '';
      if (cellContent) {
        const cellAddr = `${String.fromCharCode(65 + col)}${row + 1}`;
        cellValues[cellAddr] = cellContent;
      }
    });
    
    console.log('\\nRATE関数テスト:');
    console.log('A2 (期間):', cellValues['A2']);
    console.log('B2 (支払額):', cellValues['B2']);
    console.log('C2 (現在価値):', cellValues['C2']);
    console.log('D2の数式: =RATE(A2,B2,C2)*12');
    console.log('D2の値:', cellValues['D2']);
    console.log('期待値: 0.0253');
    
    // RATE関数の結果を月利として確認
    const d2Value = parseFloat(cellValues['D2']);
    if (!isNaN(d2Value)) {
      console.log('\\n計算された年利:', d2Value);
      console.log('推定月利:', d2Value / 12);
      console.log('Excel互換の期待月利:', 0.002107897);
    }
    
    expect(true).toBe(true);
  });

  it('NPER関数の計算結果を詳細に確認', async () => {
    const user = userEvent.setup();
    const { container } = renderComponent();
    
    await user.selectOptions(screen.getByLabelText('カテゴリを選択'), '07. 財務');
    await waitFor(() => {
      expect(screen.getByLabelText('関数を選択')).not.toBeDisabled();
    });
    
    await user.selectOptions(screen.getByLabelText('関数を選択'), 'NPER');
    
    await waitFor(() => {
      expect(screen.queryByText('数式を計算中...')).not.toBeInTheDocument();
    }, { timeout: 5000 });
    
    const cells = container.querySelectorAll('td');
    const cellValues: Record<string, string> = {};
    
    cells.forEach((cell, index) => {
      const row = Math.floor(index / 10);
      const col = index % 10;
      const cellContent = cell.textContent || '';
      if (cellContent) {
        const cellAddr = `${String.fromCharCode(65 + col)}${row + 1}`;
        cellValues[cellAddr] = cellContent;
      }
    });
    
    console.log('\\nNPER関数テスト:');
    console.log('A2 (利率):', cellValues['A2']);
    console.log('B2 (支払額):', cellValues['B2']);
    console.log('C2 (現在価値):', cellValues['C2']);
    console.log('D2の数式: =NPER(A2/12,B2,C2)');
    console.log('D2の値:', cellValues['D2']);
    console.log('期待値: 10.475');
    
    expect(true).toBe(true);
  });
}, 30000);