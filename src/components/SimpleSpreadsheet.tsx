import { useState } from 'react';
import Spreadsheet from 'react-spreadsheet';
import type { Matrix } from 'react-spreadsheet';
import { HyperFormula } from 'hyperformula';

interface CellData {
  value: string | number | boolean;
  readOnly?: boolean;
}

const SimpleSpreadsheet: React.FC = () => {

  // 初期データ
  const initialData: Matrix<CellData> = [
    [{ value: '項目A' }, { value: '項目B' }, { value: '合計' }, { value: '倍率' }],
    [{ value: 10 }, { value: 20 }, { value: '=A2+B2' }, { value: '=C2*2' }],
    [{ value: 5 }, { value: 15 }, { value: '=A3+B3' }, { value: '=C3/2' }],
    [{ value: '=SUM(A2:A3)' }, { value: '=SUM(B2:B3)' }, { value: '=SUM(C2:C3)' }, { value: '=AVERAGE(A2:C3)' }],
    [{ value: '' }, { value: '' }, { value: '' }, { value: '' }],
    [{ value: '' }, { value: '' }, { value: '' }, { value: '' }],
  ];

  // データを評価する関数
  const evaluateData = (data: Matrix<CellData>): Matrix<CellData> => {
    // データ配列を作成
    const rawData = data.map(row => 
      row.map(cell => cell?.value ?? '')
    );

    // HyperFormulaインスタンスを配列から作成
    const hf = HyperFormula.buildFromArray(rawData, {
      licenseKey: 'gpl-v3',
    });

    // 評価された値を取得
    return data.map((row, rowIndex) => 
      row.map((_cell, colIndex) => {
        const evaluatedValue = hf.getCellValue({ sheet: 0, row: rowIndex, col: colIndex });
        // エラーオブジェクトの場合は文字列に変換
        if (evaluatedValue && typeof evaluatedValue === 'object' && 'type' in evaluatedValue) {
          return { value: '#ERROR!' };
        }
        return { value: evaluatedValue ?? '' };
      })
    );
  };

  const [data, setData] = useState<Matrix<CellData>>(() => evaluateData(initialData));

  const handleChange = (newData: Matrix<CellData>) => {
    const evaluatedData = evaluateData(newData);
    setData(evaluatedData);
  };

  return (
    <div className="spreadsheet-demo">
      <h2>Excel風スプレッドシートデモ (HyperFormula搭載)</h2>
      <div className="info">
        <p>サポートされている関数例:</p>
        <ul>
          <li>基本演算: +, -, *, /</li>
          <li>SUM(範囲) - 合計</li>
          <li>AVERAGE(範囲) - 平均</li>
          <li>MIN(範囲) - 最小値</li>
          <li>MAX(範囲) - 最大値</li>
          <li>COUNT(範囲) - カウント</li>
          <li>IF(条件, 真の値, 偽の値)</li>
          <li>その他多数のExcel関数</li>
        </ul>
      </div>
      <div className="spreadsheet-container">
        <Spreadsheet 
          data={data} 
          onChange={handleChange}
          columnLabels={['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']}
          rowLabels={['1', '2', '3', '4', '5', '6', '7', '8']}
        />
      </div>
    </div>
  );
};

export default SimpleSpreadsheet;