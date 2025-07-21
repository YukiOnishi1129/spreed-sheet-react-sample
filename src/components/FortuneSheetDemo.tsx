import { useState } from 'react';
import { Workbook } from '@fortune-sheet/react';

const FortuneSheetDemo: React.FC = () => {
  // Fortune Sheetのデータ構造
  const [data, setData] = useState([
    {
      name: "Sheet1",
      status: 1,
      order: 0,
      data: [
        // ヘッダー行
        [
          { v: "項目A" },
          { v: "項目B" },
          { v: "合計" },
          { v: "倍率" },
          null, null, null, null
        ],
        // データ行1
        [
          { v: 10 },
          { v: 20 },
          { v: 30, f: "=A2+B2" },
          { v: 60, f: "=C2*2" },
          null, null, null, null
        ],
        // データ行2
        [
          { v: 5 },
          { v: 15 },
          { v: 20, f: "=A3+B3" },
          { v: 10, f: "=C3/2" },
          null, null, null, null
        ],
        // 集計行
        [
          { v: 15, f: "=SUM(A2:A3)" },
          { v: 35, f: "=SUM(B2:B3)" },
          { v: 50, f: "=SUM(C2:C3)" },
          { v: 16.67, f: "=AVERAGE(A2:C3)" },
          null, null, null, null
        ],
        // 空行
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null]
      ]
    }
  ]);

  const handleChange = (newData: any) => {
    setData(newData);
    console.log('データが変更されました:', newData);
  };

  return (
    <div className="fortune-sheet-demo">
      <h2>Fortune Sheet デモ (Excel風スプレッドシート)</h2>
      
      <div className="info">
        <h3>特徴:</h3>
        <ul>
          <li>Excel風のUI（Handsontableライク）</li>
          <li>数式の自動計算</li>
          <li>セルの選択、コピー＆ペースト</li>
          <li>右クリックコンテキストメニュー</li>
          <li>MIT ライセンス（商用利用可能）</li>
        </ul>
        
        <h3>数式の例:</h3>
        <ul>
          <li>=A1+B1 (足し算)</li>
          <li>=SUM(A1:A10) (範囲の合計)</li>
          <li>=AVERAGE(A1:A10) (平均)</li>
          <li>=MAX(A1:C3) (最大値)</li>
          <li>=MIN(A1:C3) (最小値)</li>
          <li>=IF(A1&gt;10,"大","小") (条件分岐)</li>
        </ul>
      </div>

      <div className="spreadsheet-container" style={{ height: '500px', marginTop: '20px' }}>
        <Workbook
          data={data}
          onChange={handleChange}
        />
      </div>
      
      <div className="tips" style={{ marginTop: '20px', fontSize: '14px', color: '#666' }}>
        <p><strong>使い方:</strong></p>
        <ul>
          <li>セルをダブルクリックまたは F2 で編集モード</li>
          <li>数式は = で始める</li>
          <li>Ctrl+C でコピー、Ctrl+V でペースト</li>
          <li>右クリックでコンテキストメニュー</li>
          <li>Enterで次の行、Tabで次の列に移動</li>
        </ul>
      </div>
    </div>
  );
};

export default FortuneSheetDemo;