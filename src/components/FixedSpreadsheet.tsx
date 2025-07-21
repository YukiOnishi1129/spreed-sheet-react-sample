import { useState, useCallback, useRef } from 'react';
import Spreadsheet from 'react-spreadsheet';
import type { CellBase, Matrix, Selection } from 'react-spreadsheet';
import { HyperFormula } from 'hyperformula';

interface CellData extends CellBase<any> {
  formula?: string;
  readOnly?: boolean;
}

const FixedSpreadsheet: React.FC = () => {
  // HyperFormulaインスタンス
  const hfRef = useRef<HyperFormula | null>(null);
  
  // 選択中のセル
  const [selected, setSelected] = useState<{ row: number; column: number } | null>(null);
  const [formulaBarValue, setFormulaBarValue] = useState('');
  
  // セルのデータと数式を管理
  const [cellData, setCellData] = useState<Matrix<CellData>>(() => {
    const initial: Matrix<CellData> = [];
    // 8x8のグリッドを初期化
    for (let i = 0; i < 8; i++) {
      initial[i] = [];
      for (let j = 0; j < 8; j++) {
        initial[i][j] = { value: '' };
      }
    }
    
    // サンプルデータ
    initial[0][0] = { value: '項目A' };
    initial[0][1] = { value: '項目B' };
    initial[0][2] = { value: '合計' };
    initial[0][3] = { value: '倍率' };
    
    initial[1][0] = { value: 10 };
    initial[1][1] = { value: 20 };
    initial[1][2] = { value: 30, formula: '=A2+B2' };
    initial[1][3] = { value: 60, formula: '=C2*2' };
    
    initial[2][0] = { value: 5 };
    initial[2][1] = { value: 15 };
    initial[2][2] = { value: 20, formula: '=A3+B3' };
    initial[2][3] = { value: 10, formula: '=C3/2' };
    
    return initial;
  });
  
  // HyperFormulaで全体を再評価
  const evaluateAll = useCallback((data: Matrix<CellData>) => {
    // データ配列を作成
    const rawData: (string | number)[][] = [];
    for (let i = 0; i < 8; i++) {
      rawData[i] = [];
      for (let j = 0; j < 8; j++) {
        const cell = data[i]?.[j];
        if (cell?.formula) {
          rawData[i][j] = cell.formula;
        } else {
          rawData[i][j] = cell?.value ?? '';
        }
      }
    }
    
    // HyperFormulaインスタンスを更新
    if (!hfRef.current) {
      hfRef.current = HyperFormula.buildFromArray(rawData, {
        licenseKey: 'gpl-v3',
      });
    } else {
      hfRef.current.setSheetContent(0, rawData);
    }
    
    // 評価結果を反映
    const newData: Matrix<CellData> = [];
    for (let i = 0; i < 8; i++) {
      newData[i] = [];
      for (let j = 0; j < 8; j++) {
        const cell = data[i]?.[j] || { value: '' };
        
        if (cell.formula) {
          const evaluatedValue = hfRef.current.getCellValue({ sheet: 0, row: i, col: j });
          let displayValue: any = '';
          
          if (evaluatedValue && typeof evaluatedValue === 'object' && 'type' in evaluatedValue) {
            displayValue = '#ERROR!';
          } else if (typeof evaluatedValue === 'number') {
            displayValue = Math.round(evaluatedValue * 100) / 100;
          } else {
            displayValue = evaluatedValue ?? '';
          }
          
          newData[i][j] = {
            ...cell,
            value: displayValue,
            formula: cell.formula
          };
        } else {
          newData[i][j] = cell;
        }
      }
    }
    
    return newData;
  }, []);
  
  // 初期評価
  useState(() => {
    setCellData(evaluateAll(cellData));
  });
  
  // セル選択時
  const handleSelect = useCallback((selection: Selection) => {
    if (selection && typeof selection === 'object' && 'min' in selection) {
      const min = selection.min as { row: number; column: number };
      const { row, column } = min;
      setSelected({ row, column });
      
      const cell = cellData[row]?.[column];
      if (cell?.formula) {
        setFormulaBarValue(cell.formula);
      } else {
        setFormulaBarValue(String(cell?.value ?? ''));
      }
    }
  }, [cellData]);
  
  // 数式バーの値変更
  const handleFormulaBarChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setFormulaBarValue(e.target.value);
  };
  
  // 数式バーでEnter押下時
  const handleFormulaBarKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' && selected) {
      e.preventDefault();
      
      const newData = [...cellData];
      const { row, column } = selected;
      
      // 行と列が存在することを確認
      if (!newData[row]) newData[row] = [];
      
      // セルを更新
      if (formulaBarValue.startsWith('=')) {
        newData[row][column] = {
          value: '',
          formula: formulaBarValue
        };
      } else {
        const numValue = parseFloat(formulaBarValue);
        newData[row][column] = {
          value: !isNaN(numValue) && formulaBarValue !== '' ? numValue : formulaBarValue,
          formula: undefined
        };
      }
      
      // 全体を再評価
      const evaluated = evaluateAll(newData);
      setCellData(evaluated);
      
      // 更新後の値を数式バーに反映
      const updatedCell = evaluated[row][column];
      if (updatedCell?.formula) {
        setFormulaBarValue(updatedCell.formula);
      }
    }
  };
  
  // スプレッドシートからの直接編集を処理
  const handleCellsChanged = useCallback((changes: Matrix<CellData>) => {
    // 直接編集を無効化する場合はここで何もしない
    // 有効にする場合は以下のコメントを解除
    
    const newData = [...cellData];
    let hasChanges = false;
    
    changes.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        const oldCell = cellData[rowIndex]?.[colIndex];
        const newValue = cell?.value;
        
        if (newValue !== oldCell?.value) {
          hasChanges = true;
          if (!newData[rowIndex]) newData[rowIndex] = [];
          
          if (typeof newValue === 'string' && newValue.startsWith('=')) {
            newData[rowIndex][colIndex] = {
              value: '',
              formula: newValue
            };
          } else {
            const numValue = parseFloat(String(newValue));
            newData[rowIndex][colIndex] = {
              value: !isNaN(numValue) && newValue !== '' ? numValue : newValue,
              formula: undefined
            };
          }
        }
      });
    });
    
    if (hasChanges) {
      const evaluated = evaluateAll(newData);
      setCellData(evaluated);
    }
  }, [cellData, evaluateAll]);
  
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
            placeholder={selected ? `セル ${String.fromCharCode(65 + selected.column)}${selected.row + 1}` : 'セルを選択してください'}
          />
        </div>
      </div>
      
      <div className="spreadsheet-container">
        <Spreadsheet
          data={cellData}
          onChange={handleCellsChanged}
          onSelect={handleSelect}
          columnLabels={['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']}
          rowLabels={['1', '2', '3', '4', '5', '6', '7', '8']}
        />
      </div>
      
      <div className="info">
        <h3>使い方:</h3>
        <ul>
          <li>セルをクリックして選択</li>
          <li>数式バーに値や数式を入力</li>
          <li>Enterキーで確定</li>
          <li>セルに直接入力することも可能</li>
        </ul>
        <h3>数式の例:</h3>
        <ul>
          <li>=A1+B1 (足し算)</li>
          <li>=A1-B1 (引き算)</li>
          <li>=A1*B1 (掛け算)</li>
          <li>=A1/B1 (割り算)</li>
          <li>=SUM(A1:A5) (範囲の合計)</li>
          <li>=AVERAGE(A1:B3) (平均)</li>
          <li>=MAX(A1:C3) (最大値)</li>
          <li>=MIN(A1:C3) (最小値)</li>
        </ul>
      </div>
    </div>
  );
};

export default FixedSpreadsheet;