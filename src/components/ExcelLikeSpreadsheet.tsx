import { useState, useCallback, useRef, useEffect } from 'react';
import Spreadsheet from 'react-spreadsheet';
import type { Matrix, CellBase } from 'react-spreadsheet';
import { HyperFormula } from 'hyperformula';

interface ExtendedCell extends CellBase<any> {
  formula?: string;
}

type ExtendedMatrix = Matrix<ExtendedCell>;

const ExcelLikeSpreadsheet: React.FC = () => {
  // HyperFormulaインスタンスを保持
  const hfRef = useRef<HyperFormula | null>(null);
  
  // 選択中のセル
  const [selectedCell, setSelectedCell] = useState<{ row: number; col: number } | null>(null);
  const [formulaBarValue, setFormulaBarValue] = useState<string>('');
  
  // 初期データ（数式を保持）
  const initialFormulas: ExtendedMatrix = [
    [{ value: '項目A' }, { value: '項目B' }, { value: '合計' }, { value: '倍率' }],
    [{ value: 10 }, { value: 20 }, { value: 30, formula: '=A2+B2' }, { value: 60, formula: '=C2*2' }],
    [{ value: 5 }, { value: 15 }, { value: 20, formula: '=A3+B3' }, { value: 10, formula: '=C3/2' }],
    [{ value: 15, formula: '=SUM(A2:A3)' }, { value: 35, formula: '=SUM(B2:B3)' }, { value: 50, formula: '=SUM(C2:C3)' }, { value: 16.67, formula: '=AVERAGE(A2:C3)' }],
    [{ value: '' }, { value: '' }, { value: '' }, { value: '' }],
    [{ value: '' }, { value: '' }, { value: '' }, { value: '' }],
  ];

  // 数式を保持する状態
  const [formulas, setFormulas] = useState<ExtendedMatrix>(initialFormulas);
  
  // HyperFormulaで評価する関数
  const evaluateFormulas = useCallback((formulaMatrix: ExtendedMatrix): ExtendedMatrix => {
    // データ配列を作成（数式がある場合は数式を、なければ値を使用）
    const rawData = formulaMatrix.map(row => 
      row.map(cell => {
        if (cell?.formula) {
          return cell.formula;
        }
        return cell?.value ?? '';
      })
    );

    // HyperFormulaインスタンスを作成または更新
    if (!hfRef.current) {
      hfRef.current = HyperFormula.buildFromArray(rawData, {
        licenseKey: 'gpl-v3',
      });
    } else {
      // 既存のインスタンスのデータを更新
      const sheetId = hfRef.current.getSheetId('Sheet1');
      if (sheetId !== undefined) {
        hfRef.current.setSheetContent(sheetId, rawData);
      }
    }

    // 評価された値を取得して返す
    return formulaMatrix.map((row, rowIndex) => 
      row.map((cell, colIndex) => {
        if (!hfRef.current) return cell || { value: '' };
        
        const evaluatedValue = hfRef.current.getCellValue({ sheet: 0, row: rowIndex, col: colIndex });
        
        // エラーオブジェクトの場合
        if (evaluatedValue && typeof evaluatedValue === 'object' && 'type' in evaluatedValue) {
          return { 
            ...cell,
            value: '#ERROR!',
            formula: cell?.formula
          };
        }
        
        // 数式セルの場合は評価結果を表示
        if (cell?.formula) {
          // 数値の場合は適切にフォーマット
          let displayValue = evaluatedValue ?? '';
          if (typeof displayValue === 'number') {
            displayValue = Math.round(displayValue * 100) / 100; // 小数点以下2桁
          }
          
          return {
            ...cell,
            value: displayValue,
            formula: cell.formula
          };
        }
        
        // 通常のセル
        return cell || { value: '' };
      })
    );
  }, []);

  // 表示用データ
  const [displayData, setDisplayData] = useState<ExtendedMatrix>(() => evaluateFormulas(initialFormulas));

  // セルがクリックされたとき
  const handleCellClick = useCallback((row: number, col: number) => {
    setSelectedCell({ row, col });
    const cell = formulas[row]?.[col];
    if (cell?.formula) {
      setFormulaBarValue(cell.formula);
    } else {
      setFormulaBarValue(String(cell?.value ?? ''));
    }
  }, [formulas]);

  // 数式バーの値が変更されたとき
  const handleFormulaBarChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setFormulaBarValue(e.target.value);
  };

  // 数式バーでEnterが押されたとき
  const handleFormulaBarKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' && selectedCell) {
      e.preventDefault();
      updateCell(selectedCell.row, selectedCell.col, formulaBarValue);
    }
  };

  // セルを更新する関数
  const updateCell = (row: number, col: number, value: string) => {
    const newFormulas = [...formulas];
    
    // 行が存在しない場合は作成
    while (newFormulas.length <= row) {
      newFormulas.push([]);
    }
    
    // 列が存在しない場合は作成
    while (newFormulas[row].length <= col) {
      newFormulas[row].push({ value: '' });
    }
    
    // 値が'='で始まる場合は数式として保存
    if (value.startsWith('=')) {
      newFormulas[row][col] = {
        value: '',
        formula: value
      };
    } else {
      // 数値に変換可能な場合は数値として保存
      const numValue = parseFloat(value);
      newFormulas[row][col] = {
        value: !isNaN(numValue) && value !== '' ? numValue : value,
        formula: undefined
      };
    }
    
    // 状態を更新
    setFormulas(newFormulas);
    const evaluated = evaluateFormulas(newFormulas);
    setDisplayData(evaluated);
  };

  // Spreadsheetコンポーネントの変更を処理
  const handleSpreadsheetChange = useCallback((newData: Matrix<ExtendedCell>) => {
    // 各セルの変更を検出して処理
    newData.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        const oldCell = displayData[rowIndex]?.[colIndex];
        const newValue = cell?.value;
        const oldValue = oldCell?.value;
        
        // 値が変更された場合のみ更新
        if (newValue !== oldValue) {
          updateCell(rowIndex, colIndex, String(newValue ?? ''));
        }
      });
    });
  }, [displayData]);

  // セル選択時にクリックイベントを補足
  useEffect(() => {
    const handleClick = (e: MouseEvent) => {
      const target = e.target as HTMLElement;
      const cell = target.closest('.Spreadsheet__cell');
      if (cell) {
        const row = parseInt(cell.getAttribute('data-row') || '0');
        const col = parseInt(cell.getAttribute('data-col') || '0');
        handleCellClick(row, col);
      }
    };

    // Spreadsheetコンポーネントにdata属性を追加するためのMutationObserver
    const observer = new MutationObserver(() => {
      document.querySelectorAll('.Spreadsheet__cell').forEach((cell, index) => {
        const row = Math.floor(index / 8); // 8列と仮定
        const col = index % 8;
        cell.setAttribute('data-row', String(row));
        cell.setAttribute('data-col', String(col));
      });
    });

    observer.observe(document.body, { childList: true, subtree: true });
    document.addEventListener('click', handleClick);

    return () => {
      observer.disconnect();
      document.removeEventListener('click', handleClick);
    };
  }, [handleCellClick]);

  return (
    <div className="excel-like-spreadsheet">
      <div className="formula-bar-container">
        <div className="formula-bar">
          <span className="formula-label">fx</span>
          <input
            type="text"
            className="formula-input"
            value={formulaBarValue}
            onChange={handleFormulaBarChange}
            onKeyDown={handleFormulaBarKeyDown}
            placeholder={selectedCell ? `セル ${String.fromCharCode(65 + selectedCell.col)}${selectedCell.row + 1}` : '数式を入力'}
          />
        </div>
      </div>
      
      <div className="spreadsheet-container">
        <Spreadsheet 
          data={displayData} 
          onChange={handleSpreadsheetChange}
          columnLabels={['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']}
          rowLabels={['1', '2', '3', '4', '5', '6', '7', '8']}
        />
      </div>
      
      <div className="info">
        <p>使い方:</p>
        <ul>
          <li>セルをクリックして選択</li>
          <li>数式バーに式を入力してEnterで確定</li>
          <li>=で始まる文字列は数式として評価されます</li>
        </ul>
        <p>使用例:</p>
        <ul>
          <li>=A1+B1 (セルの足し算)</li>
          <li>=SUM(A1:A10) (範囲の合計)</li>
          <li>=AVERAGE(B2:B5) (平均)</li>
          <li>=MAX(A1:C3) (最大値)</li>
          <li>=IF(A1&gt;10,"大","小") (条件分岐)</li>
        </ul>
      </div>
    </div>
  );
};

export default ExcelLikeSpreadsheet;