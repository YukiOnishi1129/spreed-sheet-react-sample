import { describe, it, expect, vi } from 'vitest';
import { render, screen } from '@testing-library/react';
import { BrowserRouter } from 'react-router-dom';
import TestSpreadsheet from '../TestSpreadsheet';

// モックを簡略化
vi.mock('react-spreadsheet', () => ({
  default: vi.fn(() => <div data-testid="spreadsheet">Spreadsheet Mock</div>)
}));

describe('TestSpreadsheet Simple Test', () => {
  it('コンポーネントが正しくレンダリングされる', () => {
    render(
      <BrowserRouter>
        <TestSpreadsheet />
      </BrowserRouter>
    );
    
    expect(screen.getByText('Excel関数テストモード')).toBeInTheDocument();
    expect(screen.getByLabelText('カテゴリを選択')).toBeInTheDocument();
  });
});