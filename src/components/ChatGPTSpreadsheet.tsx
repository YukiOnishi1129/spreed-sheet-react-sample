import { useState, useEffect } from 'react';
import Spreadsheet from 'react-spreadsheet';
import type { Matrix } from 'react-spreadsheet';
import { HyperFormula } from 'hyperformula';

// ChatGPT APIのレスポンス型定義
interface ExcelFunctionResponse {
  function_name: string;
  description: string;
  syntax: string;
  category: string;
  spreadsheet_data: any[][];
  examples?: string[];
}


const ChatGPTSpreadsheet: React.FC = () => {
  const [isLoading, setIsLoading] = useState(false);
  const [userInput, setUserInput] = useState('');
  const [currentFunction, setCurrentFunction] = useState<ExcelFunctionResponse | null>(null);
  const [selectedCellFormula, setSelectedCellFormula] = useState<string>('');
  const [selectedCellAddress, setSelectedCellAddress] = useState<string>('');
  
  
  // react-spreadsheet用のデータ構造
  const [sheetData, setSheetData] = useState<Matrix<any>>([
    [{ value: "関数を検索してください", className: "header-cell" }, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null]
  ]);

  // sheetDataの変更を監視
  useEffect(() => {
    console.log('sheetData変更検知:', sheetData);
  }, [sheetData]);

  // ChatGPT APIを呼び出す（模擬実装）
  const fetchExcelFunction = async (query: string): Promise<ExcelFunctionResponse> => {
    // 実際の実装では以下のようになります：
    /*
    const response = await fetch('/api/chatgpt', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        prompt: `以下のユーザーの要求に基づいて、Excel/スプレッドシート関数とサンプルデータを JSON 形式で返してください：
        
        ユーザー要求: "${query}"
        
        回答は以下の形式で返してください：
        {
          "function_name": "関数名",
          "description": "関数の説明",
          "syntax": "関数の構文",
          "category": "関数のカテゴリ",
          "spreadsheet_data": [
            // Fortune Sheet形式のデータ配列
            // 例：[{"v": "値", "ct": {"t": "s"}}, {"v": 100, "f": "=A1+10"}]
          ],
          "examples": ["使用例1", "使用例2"]
        }
        
        spreadsheet_dataは、関数の使用例が分かりやすい実用的なデータにしてください。`
      })
    });
    return response.json();
    */

    // デモ用のモック関数集
    const mockFunctions: { [key: string]: ExcelFunctionResponse } = {
      'sum': {
        function_name: 'SUM',
        description: '範囲内の数値の合計を計算します',
        syntax: 'SUM(range1, [range2], ...)',
        category: '数学関数',
        spreadsheet_data: [
          [
            { v: "月", ct: { t: "s" }, bg: "#E3F2FD" },
            { v: "売上", ct: { t: "s" }, bg: "#E3F2FD" },
            { v: "費用", ct: { t: "s" }, bg: "#E3F2FD" },
            { v: "利益", ct: { t: "s" }, bg: "#E3F2FD" },
            null, null, null, null
          ],
          [
            { v: "1月", ct: { t: "s" } },
            { v: 100000, ct: { t: "n" } },
            { v: 30000, ct: { t: "n" } },
            { v: 70000, f: "=B2-C2", bg: "#E8F5E8", fc: "#2E7D32" },
            null, null, null, null
          ],
          [
            { v: "2月", ct: { t: "s" } },
            { v: 120000, ct: { t: "n" } },
            { v: 35000, ct: { t: "n" } },
            { v: 85000, f: "=B3-C3", bg: "#E8F5E8", fc: "#2E7D32" },
            null, null, null, null
          ],
          [
            { v: "3月", ct: { t: "s" } },
            { v: 150000, ct: { t: "n" } },
            { v: 40000, ct: { t: "n" } },
            { v: 110000, f: "=B4-C4", bg: "#E8F5E8", fc: "#2E7D32" },
            null, null, null, null
          ],
          [
            { v: "合計", ct: { t: "s" }, bg: "#E8F5E8" },
            { v: 370000, f: "=SUM(B2:B4)", bg: "#FFE0B2", fc: "#D84315", bl: 1 },
            { v: 105000, f: "=SUM(C2:C4)", bg: "#FFE0B2", fc: "#D84315", bl: 1 },
            { v: 265000, f: "=SUM(D2:D4)", bg: "#FFE0B2", fc: "#D84315", bl: 1 },
            null, null, null, null
          ],
          [null, null, null, null, null, null, null, null],
          [null, null, null, null, null, null, null, null],
          [null, null, null, null, null, null, null, null]
        ],
        examples: ['=SUM(A1:A10)', '=SUM(A1,B1,C1)', '=SUM(A1:A5,C1:C5)']
      },
      'vlookup': {
        function_name: 'VLOOKUP',
        description: 'テーブルの左端の列で値を検索し、同じ行の指定した列から値を返します',
        syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
        category: '検索関数',
        spreadsheet_data: [
          [
            { v: "商品コード", ct: { t: "s" }, bg: "#E1F5FE" },
            { v: "商品名", ct: { t: "s" }, bg: "#E1F5FE" },
            { v: "価格", ct: { t: "s" }, bg: "#E1F5FE" },
            { v: "検索コード", ct: { t: "s" }, bg: "#FFF8E1" },
            { v: "検索結果", ct: { t: "s" }, bg: "#FFF8E1" },
            null, null, null
          ],
          [
            { v: "P001", ct: { t: "s" }, bg: "#F0F4C3" },
            { v: "ノートPC", ct: { t: "s" }, bg: "#F0F4C3" },
            { v: 80000, ct: { t: "n" }, bg: "#F0F4C3" },
            { v: "P002", ct: { t: "s" }, bg: "#FFECB3" },
            { v: null, f: "=VLOOKUP(D2,A2:C4,2,0)", bg: "#FFE0B2", fc: "#D84315", bl: 1 },
            null, null, null
          ],
          [
            { v: "P002", ct: { t: "s" }, bg: "#F0F4C3" },
            { v: "タブレット", ct: { t: "s" }, bg: "#F0F4C3" },
            { v: 45000, ct: { t: "n" }, bg: "#F0F4C3" },
            { v: "P003", ct: { t: "s" }, bg: "#FFECB3" },
            { v: null, f: "=VLOOKUP(D3,A2:C4,2,0)", bg: "#FFE0B2", fc: "#D84315", bl: 1 },
            null, null, null
          ],
          [
            { v: "P003", ct: { t: "s" }, bg: "#F0F4C3" },
            { v: "スマートフォン", ct: { t: "s" }, bg: "#F0F4C3" },
            { v: 65000, ct: { t: "n" }, bg: "#F0F4C3" },
            null, null, null, null, null
          ],
          [null, null, null, null, null, null, null, null],
          [null, null, null, null, null, null, null, null],
          [null, null, null, null, null, null, null, null],
          [null, null, null, null, null, null, null, null]
        ],
        examples: ['=VLOOKUP("P001",A2:C4,2,0)', '=VLOOKUP(D2,A2:C4,3,0)']
      },
      'if': {
        function_name: 'IF',
        description: '条件に基づいて異なる値を返します',
        syntax: 'IF(logical_test, value_if_true, value_if_false)',
        category: '論理関数',
        spreadsheet_data: [
          [
            { v: "学生名", ct: { t: "s" }, bg: "#E8EAF6" },
            { v: "点数", ct: { t: "s" }, bg: "#E8EAF6" },
            { v: "判定", ct: { t: "s" }, bg: "#E1F5FE" },
            { v: "ランク", ct: { t: "s" }, bg: "#F3E5F5" },
            null, null, null, null
          ],
          [
            { v: "田中", ct: { t: "s" } },
            { v: 85, ct: { t: "n" } },
            { v: "合格", f: '=IF(B2>=60,"合格","不合格")', bg: "#4CAF50", fc: "#FFFFFF", bl: 1 },
            { v: "A", f: '=IF(B2>=90,"S",IF(B2>=80,"A",IF(B2>=70,"B","C")))', bg: "#9C27B0", fc: "#FFFFFF", bl: 1 },
            null, null, null, null
          ],
          [
            { v: "佐藤", ct: { t: "s" } },
            { v: 45, ct: { t: "n" } },
            { v: "不合格", f: '=IF(B3>=60,"合格","不合格")', bg: "#4CAF50", fc: "#FFFFFF", bl: 1 },
            { v: "C", f: '=IF(B3>=90,"S",IF(B3>=80,"A",IF(B3>=70,"B","C")))', bg: "#9C27B0", fc: "#FFFFFF", bl: 1 },
            null, null, null, null
          ],
          [
            { v: "鈴木", ct: { t: "s" } },
            { v: 92, ct: { t: "n" } },
            { v: "合格", f: '=IF(B4>=60,"合格","不合格")', bg: "#4CAF50", fc: "#FFFFFF", bl: 1 },
            { v: "S", f: '=IF(B4>=90,"S",IF(B4>=80,"A",IF(B4>=70,"B","C")))', bg: "#9C27B0", fc: "#FFFFFF", bl: 1 },
            null, null, null, null
          ],
          [null, null, null, null, null, null, null, null],
          [null, null, null, null, null, null, null, null],
          [null, null, null, null, null, null, null, null],
          [null, null, null, null, null, null, null, null]
        ],
        examples: ['=IF(A1>10,"大","小")', '=IF(B1="",0,B1*2)']
      }
    };

    // デモ用：クエリに基づいて適切な関数を返す
    const lowerQuery = query.toLowerCase();
    if (lowerQuery.includes('合計') || lowerQuery.includes('sum')) {
      return mockFunctions.sum;
    } else if (lowerQuery.includes('検索') || lowerQuery.includes('vlookup') || lowerQuery.includes('lookup')) {
      return mockFunctions.vlookup;
    } else if (lowerQuery.includes('条件') || lowerQuery.includes('if') || lowerQuery.includes('判定')) {
      return mockFunctions.if;
    } else {
      // ランダムに関数を選択
      const functions = Object.values(mockFunctions);
      return functions[Math.floor(Math.random() * functions.length)];
    }
  };

  // 共通の検索実行関数
  const executeSearch = async (query: string) => {
    if (isLoading) return; // 既に実行中の場合は何もしない
    
    console.log('検索開始:', query);
    setIsLoading(true);
    
    try {
      const response = await fetchExcelFunction(query);
      console.log('APIレスポンス:', response);
      console.log('スプレッドシートデータの構造:', response.spreadsheet_data);
      console.log('最初の行:', response.spreadsheet_data[0]);
      console.log('最初のセル:', response.spreadsheet_data[0]?.[0]);
      
      setCurrentFunction(response);
      
      // HyperFormulaにデータを設定して計算
      const rawData = response.spreadsheet_data.map(row => 
        row.map(cell => {
          if (!cell) return '';
          return cell.f ? cell.f : cell.v || '';
        })
      );
      
      console.log('VLOOKUPデバッグ - rawData:', rawData);
      console.log('元データ構造:', response.spreadsheet_data);

      // HyperFormulaでデータを処理
      let calculationResults: any[][] = [];
      try {
        const tempHf = HyperFormula.buildFromArray(rawData, {
          licenseKey: 'gpl-v3',
          useColumnIndex: false,
          useArrayArithmetic: true,
          smartRounding: true,
          useStats: true,
          precisionEpsilon: 1e-13,
          precisionRounding: 14
        });
        
        // デバッグ: HyperFormulaインスタンスをチェック
        console.log('HyperFormulaインスタンス作成完了:', !!tempHf);
        
        // 利用可能な関数をチェック（メソッドが存在する場合のみ）
        if (typeof tempHf.getFunctionNames === 'function') {
          console.log('HyperFormula利用可能な関数:', tempHf.getFunctionNames());
        } else {
          console.log('getFunctionNames メソッドは利用できません（古いバージョンの可能性）');
        }
        
        // 計算結果を取得
        calculationResults = rawData.map((row, rowIndex) => 
          row.map((cell, colIndex) => {
            try {
              const result = tempHf.getCellValue({ sheet: 0, row: rowIndex, col: colIndex });
              
              // VLOOKUPセルの場合は詳細デバッグ
              if (typeof cell === 'string' && cell.includes('VLOOKUP')) {
                console.log(`=== VLOOKUP詳細デバッグ ===`);
                console.log(`セル位置: [${rowIndex}][${colIndex}]`);
                console.log(`数式: ${cell}`);
                console.log(`計算結果: ${result}`);
                console.log(`結果の型: ${typeof result}`);
                console.log(`検索値 D${rowIndex+1}の値:`, rawData[rowIndex]?.[3]); // D列は4番目（インデックス3）
                console.log(`検索テーブル A2:C4:`);
                for (let i = 1; i <= 3; i++) { // A2:C4 = 行1-3
                  console.log(`  行${i+1}:`, rawData[i]?.slice(0, 3));
                }
                
                // セルの座標も確認
                console.log(`HyperFormulaセル座標確認:`);
                console.log(`  D${rowIndex+1} = `, tempHf.getCellValue({ sheet: 0, row: rowIndex, col: 3 }));
                console.log(`  A2 = `, tempHf.getCellValue({ sheet: 0, row: 1, col: 0 }));
                console.log(`  B2 = `, tempHf.getCellValue({ sheet: 0, row: 1, col: 1 }));
                console.log(`  A3 = `, tempHf.getCellValue({ sheet: 0, row: 2, col: 0 }));
                console.log(`  B3 = `, tempHf.getCellValue({ sheet: 0, row: 2, col: 1 }));
                console.log(`  A4 = `, tempHf.getCellValue({ sheet: 0, row: 3, col: 0 }));
                console.log(`  B4 = `, tempHf.getCellValue({ sheet: 0, row: 3, col: 1 }));
                console.log(`========================`);
              }
              
              return result;
            } catch (cellError) {
              console.error(`セル[${rowIndex}][${colIndex}]でエラー:`, cellError);
              
              // VLOOKUPの場合はマニュアル計算を試す
              if (typeof cell === 'string' && cell.includes('VLOOKUP')) {
                console.log('VLOOKUPの手動計算を試行中...');
                return '#MANUAL_CALC_NEEDED';
              }
              
              return cell;
            }
          })
        );
      } catch (error) {
        console.error('HyperFormula計算エラー:', error);
        calculationResults = rawData; // フォールバック
      }

      // react-spreadsheet用のデータに変換
      const convertedData: Matrix<any> = response.spreadsheet_data.map((row, rowIndex) => 
        row.map((cell, colIndex) => {
          if (!cell) return null;
          
          let className = '';
          let calculatedValue = cell.v;
          
          // 数式セル（関数の結果）の場合
          if (cell.f) {
            // SUM関数やVLOOKUP関数の場合はオレンジ枠にする
            if (cell.f.toUpperCase().includes('SUM(') || cell.f.toUpperCase().includes('VLOOKUP(')) {
              className = 'formula-result-cell';
            } else {
              // その他の関数は緑系の色にする
              className = 'other-formula-cell';
            }
            // HyperFormulaから計算結果を取得
            const result = calculationResults[rowIndex]?.[colIndex];
            calculatedValue = result !== null && result !== undefined ? result : '#ERROR!';
          } else if (cell.bg) {
            // 背景色がある通常のセル
            if (cell.bg.includes('E3F2FD') || cell.bg.includes('E1F5FE')) {
              className = 'header-cell';
            } else if (cell.bg.includes('F0F4C3') || cell.bg.includes('FFECB3')) {
              className = 'data-cell';
            } else if (cell.bg.includes('FFF8E1') || cell.bg.includes('FFF3E0')) {
              className = 'input-cell';
            }
          }
          
          const cellData: any = {
            value: calculatedValue,
            className: className
          };
          
          // 数式セルの場合、ツールチップと数式プロパティを追加
          if (cell.f) {
            cellData.title = `数式: ${cell.f}`;
            cellData.formula = cell.f;
            cellData['data-formula'] = cell.f; // HTML属性として追加
            // react-spreadsheetでツールチップを表示するための追加設定
            cellData.DataEditor = undefined; // デフォルトエディターを無効化してツールチップを優先
          }
          
          return cellData;
        })
      );

      console.log('元のスプレッドシートデータ:', response.spreadsheet_data);
      console.log('HyperFormulaで処理したデータ:', rawData);
      console.log('変換後のデータ:', convertedData);
      console.log('現在のシートデータ:', sheetData);
      
      setSheetData(convertedData);
      console.log('シートデータ設定完了');
      
    } catch (error) {
      console.error('関数検索エラー:', error);
      const errorMessage = error instanceof Error ? error.message : '不明なエラー';
      alert('関数の検索中にエラーが発生しました: ' + errorMessage);
    }
    
    setIsLoading(false);
  };

  const handleSearch = async () => {
    if (!userInput.trim()) return;
    await executeSearch(userInput);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleSearch();
    }
  };

  return (
    <div className="chatgpt-spreadsheet">
      <h2>ChatGPT連携 Excel関数デモ</h2>
      
      <div className="search-section" style={{ marginBottom: '20px' }}>
        <div className="search-input" style={{ display: 'flex', gap: '10px', marginBottom: '10px' }}>
          <input
            type="text"
            value={userInput}
            onChange={(e) => setUserInput(e.target.value)}
            onKeyDown={handleKeyDown}
            placeholder="例：「合計を計算する関数を知りたい」「条件分岐の関数」「データを検索する関数」"
            style={{ 
              flex: 1, 
              padding: '10px', 
              border: '1px solid #ddd', 
              borderRadius: '4px',
              fontSize: '14px'
            }}
            disabled={isLoading}
          />
          <button
            onClick={handleSearch}
            disabled={isLoading || !userInput.trim()}
            style={{
              padding: '10px 20px',
              backgroundColor: isLoading ? '#ccc' : '#007bff',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: isLoading ? 'not-allowed' : 'pointer'
            }}
          >
            {isLoading ? '検索中...' : '関数を検索'}
          </button>
        </div>
        
        <div style={{ fontSize: '12px', color: '#666', marginBottom: '5px' }}>
          クイック入力:
        </div>
        <div className="quick-buttons" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
          {['合計を計算する関数', 'データを検索する関数', '条件分岐の関数', 'ランダムな関数'].map(query => (
            <button
              key={query}
              onClick={() => { 
                setUserInput(query); // 入力欄にテキストを設定するだけ
              }}
              style={{
                padding: '5px 10px',
                backgroundColor: '#f8f9fa',
                border: '1px solid #ddd',
                borderRadius: '4px',
                cursor: 'pointer',
                fontSize: '12px'
              }}
            >
              {query}
            </button>
          ))}
        </div>
      </div>

      {currentFunction && (
        <div className="function-info" style={{ 
          backgroundColor: '#ffffff', 
          padding: '20px', 
          borderRadius: '8px', 
          marginBottom: '20px',
          border: '2px solid #e0e0e0',
          boxShadow: '0 2px 8px rgba(0,0,0,0.1)',
          color: '#333333'
        }}>
          <h3 style={{ margin: '0 0 10px 0', color: '#007bff' }}>
            {currentFunction.function_name} 関数
          </h3>
          <p style={{ margin: '8px 0', color: '#555555', fontSize: '14px' }}>
            <strong style={{ color: '#333333' }}>説明:</strong> {currentFunction.description}
          </p>
          <p style={{ margin: '8px 0', color: '#555555', fontSize: '14px' }}>
            <strong style={{ color: '#333333' }}>構文:</strong> 
            <code style={{ 
              backgroundColor: '#f5f5f5', 
              padding: '2px 6px', 
              borderRadius: '3px', 
              color: '#d63384',
              fontFamily: 'monospace',
              fontSize: '13px'
            }}>
              {currentFunction.syntax}
            </code>
          </p>
          <p style={{ margin: '8px 0', color: '#555555', fontSize: '14px' }}>
            <strong style={{ color: '#333333' }}>カテゴリ:</strong> 
            <span style={{ 
              backgroundColor: '#e7f3ff', 
              color: '#0066cc', 
              padding: '2px 8px', 
              borderRadius: '12px', 
              fontSize: '12px',
              fontWeight: '500'
            }}>
              {currentFunction.category}
            </span>
          </p>
          <div style={{ 
            margin: '15px 0', 
            padding: '15px', 
            backgroundColor: '#f8f9fa', 
            borderRadius: '6px',
            border: '1px solid #e9ecef'
          }}>
            <strong style={{ color: '#495057', fontSize: '14px' }}>🎨 スプレッドシートの色分け:</strong>
            <ul style={{ 
              margin: '8px 0 0 0', 
              padding: '0', 
              listStyle: 'none',
              display: 'flex',
              flexWrap: 'wrap',
              gap: '8px'
            }}>
              <li style={{ display: 'flex', alignItems: 'center', margin: '4px 0' }}>
                <span style={{ 
                  backgroundColor: '#FFE0B2', 
                  color: '#D84315',
                  padding: '3px 8px', 
                  borderRadius: '4px',
                  fontSize: '11px',
                  fontWeight: 'bold',
                  border: '2px solid #FF9800',
                  marginRight: '6px'
                }}>📊 オレンジ枠</span>
                <span style={{ fontSize: '13px', color: '#666666' }}>✨ <strong>SUM・VLOOKUP関数の結果</strong></span>
              </li>
              <li style={{ display: 'flex', alignItems: 'center', margin: '4px 0' }}>
                <span style={{ 
                  backgroundColor: '#E8F5E8', 
                  color: '#2E7D32',
                  padding: '3px 8px', 
                  borderRadius: '4px',
                  fontSize: '11px',
                  fontWeight: 'bold',
                  border: '1px solid #4CAF50',
                  marginRight: '6px'
                }}>🔢 緑枠</span>
                <span style={{ fontSize: '13px', color: '#666666' }}>✨ <strong>その他の関数</strong></span>
              </li>
              <li style={{ display: 'flex', alignItems: 'center', margin: '4px 0' }}>
                <span style={{ 
                  backgroundColor: '#E3F2FD', 
                  color: '#1976D2',
                  padding: '3px 8px', 
                  borderRadius: '4px',
                  fontSize: '11px',
                  fontWeight: '600',
                  marginRight: '6px'
                }}>薄青</span>
                <span style={{ fontSize: '13px', color: '#666666' }}>ヘッダー行</span>
              </li>
              <li style={{ display: 'flex', alignItems: 'center', margin: '4px 0' }}>
                <span style={{ 
                  backgroundColor: '#F0F4C3', 
                  color: '#689F38',
                  padding: '3px 8px', 
                  borderRadius: '4px',
                  fontSize: '11px',
                  fontWeight: '600',
                  marginRight: '6px'
                }}>薄緑</span>
                <span style={{ fontSize: '13px', color: '#666666' }}>データ行</span>
              </li>
            </ul>
            <div style={{ 
              marginTop: '8px', 
              padding: '8px', 
              backgroundColor: '#FFF3E0', 
              borderRadius: '4px',
              fontSize: '12px'
            }}>
              💡 <strong>ヒント:</strong> オレンジ枠（SUM関数）や緑枠（その他関数）のセルにマウスを置くと、使用している数式が表示されます
            </div>
          </div>
          {currentFunction.examples && (
            <div style={{ marginTop: '15px' }}>
              <strong style={{ color: '#495057', fontSize: '14px' }}>💡 使用例:</strong>
              <div style={{ 
                margin: '8px 0', 
                display: 'flex', 
                flexWrap: 'wrap', 
                gap: '8px' 
              }}>
                {currentFunction.examples.map((example, index) => (
                  <code key={index} style={{
                    backgroundColor: '#f1f3f4',
                    color: '#d73a49',
                    padding: '6px 10px',
                    borderRadius: '4px',
                    fontSize: '12px',
                    fontFamily: 'Monaco, Consolas, monospace',
                    border: '1px solid #e1e4e8',
                    display: 'inline-block'
                  }}>
                    {example}
                  </code>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {/* 数式バー */}
      <div className="formula-bar-container" style={{ marginBottom: '15px' }}>
        <div className="formula-bar" style={{
          display: 'flex',
          alignItems: 'center',
          backgroundColor: '#f8f9fa',
          border: '1px solid #dee2e6',
          borderRadius: '6px',
          padding: '8px 12px',
          boxShadow: '0 1px 3px rgba(0,0,0,0.1)'
        }}>
          <div className="cell-address" style={{
            minWidth: '60px',
            padding: '4px 8px',
            backgroundColor: '#e9ecef',
            border: '1px solid #ced4da',
            borderRadius: '4px',
            fontWeight: 'bold',
            fontSize: '13px',
            color: '#495057',
            marginRight: '8px'
          }}>
            {selectedCellAddress || 'A1'}
          </div>
          <div className="formula-display" style={{
            flex: 1,
            padding: '4px 8px',
            backgroundColor: 'white',
            border: '1px solid #ced4da',
            borderRadius: '4px',
            fontFamily: 'Monaco, Consolas, "Lucida Console", monospace',
            fontSize: '13px',
            color: selectedCellFormula.startsWith('=') ? '#d73a49' : '#333',
            minHeight: '22px',
            display: 'flex',
            alignItems: 'center'
          }}>
            {selectedCellFormula || ''}
          </div>
        </div>
      </div>
      
      <div className="spreadsheet-container" style={{ height: '500px' }}>
        <Spreadsheet
          data={sheetData}
          onChange={(data) => {
            console.log('Spreadsheet onChange:', data);
            setSheetData(data);
          }}
          onSelect={(selected) => {
            // セル選択時に数式を表示
            console.log('Selected cells:', selected);
            if (selected && typeof selected === 'object' && 'min' in selected) {
              const { row, column } = selected.min as { row: number; column: number };
              const cell = sheetData[row]?.[column];
              const cellAddress = `${String.fromCharCode(65 + column)}${row + 1}`;
              setSelectedCellAddress(cellAddress);
              
              if (cell && cell.formula) {
                setSelectedCellFormula(cell.formula);
                console.log(`セル ${cellAddress} の数式:`, cell.formula);
              } else if (cell && cell.value !== undefined) {
                setSelectedCellFormula(cell.value.toString());
              } else {
                setSelectedCellFormula('');
              }
            } else {
              setSelectedCellAddress('');
              setSelectedCellFormula('');
            }
          }}
          onActivate={(point) => {
            // セルアクティブ時に数式を表示
            if (point) {
              const { row, column } = point;
              const cell = sheetData[row]?.[column];
              const cellAddress = `${String.fromCharCode(65 + column)}${row + 1}`;
              setSelectedCellAddress(cellAddress);
              
              if (cell && cell.formula) {
                setSelectedCellFormula(cell.formula);
                console.log(`アクティブセル ${cellAddress} の数式:`, cell.formula);
              } else if (cell && cell.value !== undefined) {
                setSelectedCellFormula(cell.value.toString());
              } else {
                setSelectedCellFormula('');
              }
            }
          }}
          columnLabels={['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']}
          rowLabels={['1', '2', '3', '4', '5', '6', '7', '8']}
        />
      </div>
      
      <div className="tips" style={{ marginTop: '20px', fontSize: '14px', color: '#666' }}>
        <h4>使い方:</h4>
        <ol>
          <li><strong>クイック入力ボタン</strong>をクリックして検索欄にテキストを入力</li>
          <li>または検索欄に直接、知りたい関数について自然言語で入力</li>
          <li><strong>「関数を検索」ボタン</strong>をクリックしてデモを実行</li>
          <li>スプレッドシートに関数とサンプルデータが表示されます</li>
          <li>セルをダブルクリックして数式を編集・確認可能</li>
        </ol>
        
        <h4>対応予定の機能:</h4>
        <ul>
          <li>すべてのExcel/Google Sheets関数</li>
          <li>複雑な数式の組み合わせ</li>
          <li>実用的なビジネスシナリオ</li>
          <li>関数の学習履歴</li>
        </ul>
      </div>
    </div>
  );
};

export default ChatGPTSpreadsheet;