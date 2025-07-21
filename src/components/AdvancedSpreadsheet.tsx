import { useState, useCallback } from 'react';
import Spreadsheet from 'react-spreadsheet';
import type { Matrix, CellBase } from 'react-spreadsheet';
import { HyperFormula } from 'hyperformula';

interface ExtendedCell extends CellBase<any> {
  formula?: string;
}

type ExtendedMatrix = Matrix<ExtendedCell>;

const AdvancedSpreadsheet: React.FC = () => {
  // 初期データ（数式を保持）
  const initialFormulas: ExtendedMatrix = [
    [{ value: '項目A' }, { value: '項目B' }, { value: '合計' }, { value: '倍率' }],
    [{ value: 10 }, { value: 20 }, { value: '', formula: '=A2+B2' }, { value: '', formula: '=C2*2' }],
    [{ value: 5 }, { value: 15 }, { value: '', formula: '=A3+B3' }, { value: '', formula: '=C3/2' }],
    [{ value: '', formula: '=SUM(A2:A3)' }, { value: '', formula: '=SUM(B2:B3)' }, { value: '', formula: '=SUM(C2:C3)' }, { value: '', formula: '=AVERAGE(A2:C3)' }],
    [{ value: '' }, { value: '' }, { value: '' }, { value: '' }],
    [{ value: '' }, { value: '' }, { value: '' }, { value: '' }],
  ];

  // 数式を保持する状態
  const [formulas, setFormulas] = useState<ExtendedMatrix>(initialFormulas);
  
  // 表示用のデータを評価する関数
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

    // HyperFormulaインスタンスを配列から作成
    const hf = HyperFormula.buildFromArray(rawData, {
      licenseKey: 'gpl-v3',
    });

    // 評価された値を取得して返す
    return formulaMatrix.map((row, rowIndex) => 
      row.map((cell, colIndex) => {
        const evaluatedValue = hf.getCellValue({ sheet: 0, row: rowIndex, col: colIndex });
        
        // エラーオブジェクトの場合
        if (evaluatedValue && typeof evaluatedValue === 'object' && 'type' in evaluatedValue) {
          return { 
            ...cell,
            value: '#ERROR!' 
          };
        }
        
        // 数式セルの場合は評価結果を表示
        if (cell?.formula) {
          return {
            ...cell,
            value: evaluatedValue ?? ''
          };
        }
        
        // 通常のセル
        return cell || { value: '' };
      })
    );
  }, []);

  // 表示用データ
  const [displayData, setDisplayData] = useState<ExtendedMatrix>(() => evaluateFormulas(initialFormulas));

  // セルの変更を処理
  const handleChange = useCallback((newData: Matrix<ExtendedCell>) => {
    // 新しい数式マトリックスを作成
    const newFormulas = newData.map((row, rowIndex) => 
      row.map((cell, colIndex) => {
        const oldCell = formulas[rowIndex]?.[colIndex];
        
        // 値が文字列で'='で始まる場合は数式として保存
        if (cell && typeof cell.value === 'string' && cell.value.startsWith('=')) {
          return {
            ...cell,
            formula: cell.value,
            value: '' // 値は評価時に設定される
          };
        }
        
        // 既存の数式がある場合で、新しい値が数式でない場合は数式をクリア
        if (oldCell?.formula && cell && !String(cell.value).startsWith('=')) {
          return {
            ...cell,
            formula: undefined
          };
        }
        
        // その他の場合は既存の数式を保持
        return {
          value: cell?.value ?? '',
          ...cell,
          formula: oldCell?.formula
        };
      })
    );
    
    // 数式を保存
    setFormulas(newFormulas);
    
    // 表示データを再評価
    const evaluated = evaluateFormulas(newFormulas);
    setDisplayData(evaluated);
  }, [formulas, evaluateFormulas]);

  return (
    <div className="spreadsheet-demo">
      <h2>Excel風スプレッドシートデモ (HyperFormula搭載)</h2>
      <div className="info">
        <p>使い方:</p>
        <ul>
          <li>セルをクリックして値を入力</li>
          <li>=で始まる文字列は数式として評価されます</li>
          <li>例: =A2+B2, =SUM(A1:A10), =AVERAGE(B2:B5)</li>
        </ul>
        <p>サポートされている関数:</p>
        <ul>
          <li>基本演算: +, -, *, /</li>
          <li>SUM(範囲) - 合計</li>
          <li>AVERAGE(範囲) - 平均</li>
          <li>MIN/MAX(範囲) - 最小値/最大値</li>
          <li>COUNT(範囲) - カウント</li>
          <li>IF(条件, 真の値, 偽の値)</li>
          <li>その他多数のExcel関数</li>
        </ul>
      </div>
      <div className="spreadsheet-container">
        <Spreadsheet 
          data={displayData} 
          onChange={handleChange}
          columnLabels={['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']}
          rowLabels={['1', '2', '3', '4', '5', '6', '7', '8']}
        />
      </div>
    </div>
  );
};

export default AdvancedSpreadsheet;