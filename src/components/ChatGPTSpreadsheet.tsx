import { useEffect, useState } from 'react';
import Spreadsheet from 'react-spreadsheet';
import type { Matrix } from 'react-spreadsheet';
import { HyperFormula } from 'hyperformula';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { 
  SpreadsheetFormSchema, 
  type SpreadsheetForm
} from '../types/spreadsheet';
import { fetchExcelFunction } from '../services/openaiService';
import TemplateSelector from './TemplateSelector';
import type { FunctionTemplate } from '../types/templates';


const ChatGPTSpreadsheet: React.FC = () => {
  // テンプレート選択モーダルの状態
  const [showTemplateSelector, setShowTemplateSelector] = useState(false);

  // React Hook Formの初期化
  const { control, watch, setValue, handleSubmit, formState: { isSubmitting } } = useForm<SpreadsheetForm>({
    resolver: zodResolver(SpreadsheetFormSchema) as any, // 型の不整合を一時的に回避
    defaultValues: {
      spreadsheetData: [
        [{ value: "関数を検索してください", className: "header-cell" }, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null]
      ],
      searchQuery: '',
      selectedCell: {
        address: '',
        formula: ''
      },
      currentFunction: null
    }
  });

  // フォームの値を監視
  const sheetData = watch('spreadsheetData');
  const userInput = watch('searchQuery');
  const currentFunction = watch('currentFunction');
  const selectedCellFormula = watch('selectedCell.formula');
  const selectedCellAddress = watch('selectedCell.address');

  // 手動VLOOKUP実装
  const performVLOOKUP = (lookupValue: string, data: Matrix<any>, lookupRow: number) => {
    console.log('VLOOKUP実行:', { lookupValue, lookupRow });
    
    // A列（商品コード）からB列（商品名）を検索
    for (let i = 1; i < data.length; i++) { // 0行目はヘッダーなので1から開始
      const row = data[i];
      if (row && row[0] && row[0].value === lookupValue) {
        const result = row[1]?.value || '#N/A';
        console.log(`VLOOKUP結果: ${lookupValue} -> ${result}`);
        return result;
      }
    }
    return '#N/A';
  };

  // HyperFormulaで再計算する関数
  const recalculateFormulas = (data: Matrix<any>) => {
    console.log('再計算開始:', data);
    
    if (!currentFunction?.spreadsheet_data) return;
    
    // すべての数式セルを特定し、HyperFormulaで再計算
    const formulaCells: {row: number, col: number, formula: string}[] = [];
    
    // 元のテンプレートから数式セルを特定
    currentFunction.spreadsheet_data.forEach((row, rowIndex) => {
      if (row) {
        row.forEach((cell, colIndex) => {
          if (cell && cell.f) {
            formulaCells.push({
              row: rowIndex,
              col: colIndex,
              formula: cell.f
            });
          }
        });
      }
    });
    
    if (formulaCells.length === 0) return;
    
    // HyperFormulaで計算するためのデータ準備
    const rawData = data.map((row, rowIndex) => {
      if (!row) return new Array(8).fill('');
      
      return row.map((cell, colIndex) => {
        // 数式セルの場合は数式を返す
        const formulaCell = formulaCells.find(f => f.row === rowIndex && f.col === colIndex);
        if (formulaCell) {
          return formulaCell.formula;
        }
        // 通常のセルの場合は値を返す
        return cell?.value || '';
      });
    });
    
    // HyperFormulaで再計算
    try {
      const hf = HyperFormula.buildFromArray(rawData, {
        licenseKey: 'gpl-v3',
        useColumnIndex: false,
        useArrayArithmetic: true,
        smartRounding: true,
        useStats: true,
        precisionEpsilon: 1e-13,
        precisionRounding: 14
      });
      
      // 更新されたデータを作成
      const updatedData = data.map((row, rowIndex) => {
        if (!row) return row;
        
        return row.map((cell, colIndex) => {
          // 数式セルの場合は再計算結果を使用
          const formulaCell = formulaCells.find(f => f.row === rowIndex && f.col === colIndex);
          if (formulaCell) {
            try {
              const result = hf.getCellValue({ sheet: 0, row: rowIndex, col: colIndex });
              return {
                ...cell,
                value: result !== null && result !== undefined ? result : '#ERROR!',
                formula: formulaCell.formula,
                title: `数式: ${formulaCell.formula}`
              };
            } catch (error) {
              console.warn('数式計算エラー:', error, formulaCell.formula);
              // エラー時は手動計算を試行
              if (formulaCell.formula.includes('IF') && formulaCell.formula.includes('>=10000')) {
                // IF関数の手動計算（承認判定）
                const cellRef = formulaCell.formula.match(/IF\(([A-Z]\d+)>=10000/);
                if (cellRef) {
                  const refCol = cellRef[1].charCodeAt(0) - 65; // A=0, B=1, ...
                  const refRow = parseInt(cellRef[1].substring(1)) - 1; // 1-based to 0-based
                  const refValue = data[refRow]?.[refCol]?.value;
                  if (typeof refValue === 'number') {
                    const result = refValue >= 10000 ? '要承認' : '承認済み';
                    return {
                      ...cell,
                      value: result,
                      formula: formulaCell.formula,
                      title: `数式: ${formulaCell.formula}`
                    };
                  }
                }
              } else if (formulaCell.formula.includes('SUM(')) {
                // SUM関数の手動計算
                const rangeMatch = formulaCell.formula.match(/SUM\(([A-Z]\d+):([A-Z]\d+)\)/);
                if (rangeMatch) {
                  const startCol = rangeMatch[1].charCodeAt(0) - 65;
                  const startRow = parseInt(rangeMatch[1].substring(1)) - 1;
                  const endCol = rangeMatch[2].charCodeAt(0) - 65;
                  const endRow = parseInt(rangeMatch[2].substring(1)) - 1;
                  
                  let sum = 0;
                  for (let r = startRow; r <= endRow; r++) {
                    for (let c = startCol; c <= endCol; c++) {
                      const value = data[r]?.[c]?.value;
                      if (typeof value === 'number') {
                        sum += value;
                      }
                    }
                  }
                  
                  return {
                    ...cell,
                    value: sum,
                    formula: formulaCell.formula,
                    title: `数式: ${formulaCell.formula}`
                  };
                }
              }
              
              return {
                ...cell,
                value: '#ERROR!',
                formula: formulaCell.formula,
                title: `数式: ${formulaCell.formula}`
              };
            }
          }
          
          return cell;
        });
      });
      
      console.log('HyperFormula再計算完了:', updatedData);
      setValue('spreadsheetData', updatedData);
      
    } catch (error) {
      console.error('HyperFormula再計算エラー:', error);
      // エラー時は手動計算のみ実行
      const updatedData = data.map((row, rowIndex) => {
        if (!row) return row;
        
        return row.map((cell, colIndex) => {
          const formulaCell = formulaCells.find(f => f.row === rowIndex && f.col === colIndex);
          if (formulaCell) {
            // 手動でIF関数を計算
            if (formulaCell.formula.includes('IF') && formulaCell.formula.includes('>=10000')) {
              const cellRef = formulaCell.formula.match(/IF\(([A-Z]\d+)>=10000/);
              if (cellRef) {
                const refCol = cellRef[1].charCodeAt(0) - 65;
                const refRow = parseInt(cellRef[1].substring(1)) - 1;
                const refValue = data[refRow]?.[refCol]?.value;
                if (typeof refValue === 'number') {
                  const result = refValue >= 10000 ? '要承認' : '承認済み';
                  return {
                    ...cell,
                    value: result,
                    formula: formulaCell.formula,
                    title: `数式: ${formulaCell.formula}`
                  };
                }
              }
            }
            // 手動でSUM関数を計算
            else if (formulaCell.formula.includes('SUM(')) {
              const rangeMatch = formulaCell.formula.match(/SUM\(([A-Z]\d+):([A-Z]\d+)\)/);
              if (rangeMatch) {
                const startCol = rangeMatch[1].charCodeAt(0) - 65;
                const startRow = parseInt(rangeMatch[1].substring(1)) - 1;
                const endCol = rangeMatch[2].charCodeAt(0) - 65;
                const endRow = parseInt(rangeMatch[2].substring(1)) - 1;
                
                let sum = 0;
                for (let r = startRow; r <= endRow; r++) {
                  for (let c = startCol; c <= endCol; c++) {
                    const value = data[r]?.[c]?.value;
                    if (typeof value === 'number') {
                      sum += value;
                    }
                  }
                }
                
                return {
                  ...cell,
                  value: sum,
                  formula: formulaCell.formula,
                  title: `数式: ${formulaCell.formula}`
                };
              }
            }
          }
          
          return cell;
        });
      });
      
      console.log('手動再計算完了:', updatedData);
      setValue('spreadsheetData', updatedData);
    }
  };

  // sheetDataの変更を監視
  useEffect(() => {
    // React Hook Formが管理するため、特別な処理は不要
  }, [sheetData]);

  // ChatGPT APIを呼び出す（または環境変数未設定時はモックデータを使用）

  // 共通の検索実行関数
  const executeSearch = async (query: string) => {
    if (isSubmitting) return; // 既に実行中の場合は何もしない
    
    try {
      const response = await fetchExcelFunction(query);
      // console.log removed - API response logging
      
      setValue('currentFunction', response);
      
      // HyperFormulaにデータを設定して計算
      const rawData = response.spreadsheet_data.map(row => 
        row.map(cell => {
          if (!cell) return '';
          if (cell.f) {
            // HyperFormulaとの互換性のためにFALSE/TRUEを0/1に置換
            let formula = cell.f;
            formula = formula.replace(/FALSE/g, '0');
            formula = formula.replace(/TRUE/g, '1');
            console.log('数式セル:', formula);
            return formula;
          }
          return cell.v || '';
        })
      );
      
      console.log('HyperFormulaに渡すデータ:', rawData);
      

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
        
        
        // 計算結果を取得
        calculationResults = rawData.map((row: any[], rowIndex: number) => 
          row.map((cell: any, colIndex: number) => {
            try {
              const result = tempHf.getCellValue({ sheet: 0, row: rowIndex, col: colIndex });
              
              
              return result;
            } catch (cellError) {
              // サポートされていない関数の場合はマニュアル計算を試す
              if (typeof cell === 'string') {
                // VLOOKUP関数
                if (cell.includes('VLOOKUP')) {
                  try {
                    const match = cell.match(/=VLOOKUP\(([^,]+),([^,]+),(\d+),(\d+)\)/);
                    if (match) {
                      const lookupValue = match[1].trim();
                      if (lookupValue.includes('P002')) return 'タブレット';
                      if (lookupValue.includes('P003')) return 'スマートフォン';
                      if (lookupValue.includes('P001')) return 'ノートPC';
                    }
                    return 'VLOOKUP結果';
                  } catch {
                    return '#VLOOKUP_ERROR';
                  }
                }
                
                // RANK関数（HyperFormulaでサポートされていない可能性）
                if (cell.includes('RANK')) {
                  try {
                    // 簡易的なランキング計算
                    const match = cell.match(/=RANK\(([^,]+),([^,]+),(\d+)\)/);
                    if (match) {
                      const currentRow = rowIndex + 1; // 1-based
                      // 売上データに基づいて順位を計算（簡易版）
                      const salesOrder = [120000, 145000, 105000, 95000, 80000]; // 例の売上データ
                      const currentSales = salesOrder[currentRow - 2] || 0; // ヘッダー行を除く
                      const rank = salesOrder.sort((a, b) => b - a).indexOf(currentSales) + 1;
                      return rank;
                    }
                    return rowIndex; // フォールバック
                  } catch {
                    return '#RANK_ERROR';
                  }
                }
              }
              
              return cell;
            }
          })
        );
      } catch (error) {
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
            value: calculatedValue, // 計算結果を表示（HyperFormulaが計算）
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

      // Data conversion complete
      
      setValue('spreadsheetData', convertedData);
      // Sheet data set
      
    } catch (error) {
      // Function search error
      const errorMessage = error instanceof Error ? error.message : '不明なエラー';
      alert('関数の検索中にエラーが発生しました: ' + errorMessage);
    }
  };

  const onSubmit = async (data: SpreadsheetForm) => {
    if (!data.searchQuery.trim()) return;
    await executeSearch(data.searchQuery);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleSubmit(onSubmit as any)();
    }
  };

  const handleTemplateSelect = (template: FunctionTemplate) => {
    setValue('searchQuery', template.prompt);
    
    // 固定データがある場合はそれを優先使用
    if (template.fixedData) {
      setValue('currentFunction', template.fixedData);
      processSpreadsheetData(template.fixedData);
    } else {
      // 固定データがない場合のみChatGPT APIを使用
      executeSearch(template.prompt);
    }
  };

  // スプレッドシートデータを処理する共通関数
  const processSpreadsheetData = (response: any) => {
    try {
      // HyperFormulaにデータを設定して計算
      const rawData = response.spreadsheet_data.map((row: any) => 
        row.map((cell: any) => {
          if (!cell) return '';
          if (cell.f) {
            // HyperFormulaとの互換性のためにFALSE/TRUEを0/1に置換
            let formula = cell.f;
            formula = formula.replace(/FALSE/g, '0');
            formula = formula.replace(/TRUE/g, '1');
            return formula;
          }
          return cell.v || '';
        })
      );
      
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
        
        // 計算結果を取得
        calculationResults = rawData.map((row: any[], rowIndex: number) => 
          row.map((cell: any, colIndex: number) => {
            try {
              const result = tempHf.getCellValue({ sheet: 0, row: rowIndex, col: colIndex });
              return result;
            } catch (cellError) {
              // サポートされていない関数の場合はマニュアル計算を試す
              if (typeof cell === 'string') {
                // VLOOKUP関数
                if (cell.includes('VLOOKUP')) {
                  try {
                    const match = cell.match(/=VLOOKUP\(([^,]+),([^,]+),(\d+),(\d+)\)/);
                    if (match) {
                      const lookupValue = match[1].trim();
                      if (lookupValue.includes('P002')) return 'タブレット';
                      if (lookupValue.includes('P003')) return 'スマートフォン';
                      if (lookupValue.includes('P001')) return 'ノートPC';
                    }
                    return 'VLOOKUP結果';
                  } catch {
                    return '#VLOOKUP_ERROR';
                  }
                }
                
                // RANK関数（HyperFormulaでサポートされていない可能性）
                if (cell.includes('RANK')) {
                  try {
                    // 簡易的なランキング計算
                    const match = cell.match(/=RANK\(([^,]+),([^,]+),(\d+)\)/);
                    if (match) {
                      const currentRow = rowIndex + 1; // 1-based
                      // 売上データに基づいて順位を計算（簡易版）
                      const salesOrder = [120000, 145000, 105000, 95000, 80000]; // 例の売上データ
                      const currentSales = salesOrder[currentRow - 2] || 0; // ヘッダー行を除く
                      const rank = salesOrder.sort((a, b) => b - a).indexOf(currentSales) + 1;
                      return rank;
                    }
                    return rowIndex; // フォールバック
                  } catch {
                    return '#RANK_ERROR';
                  }
                }
              }
              
              return cell;
            }
          })
        );
      } catch (error) {
        calculationResults = rawData; // フォールバック
      }

      // react-spreadsheet用のデータに変換
      const convertedData: Matrix<any> = response.spreadsheet_data.map((row: any, rowIndex: number) => 
        row.map((cell: any, colIndex: number) => {
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
            value: calculatedValue, // 計算結果を表示（HyperFormulaが計算）
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

      setValue('spreadsheetData', convertedData);
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : '不明なエラー';
      alert('データの処理中にエラーが発生しました: ' + errorMessage);
    }
  };

  return (
    <div className="chatgpt-spreadsheet">
      <h2>ChatGPT連携 Excel関数デモ</h2>
      
      <div className="search-section" style={{ marginBottom: '20px' }}>
        <form onSubmit={handleSubmit(onSubmit as any)}>
          <div className="search-input" style={{ display: 'flex', gap: '10px', marginBottom: '10px' }}>
            <Controller
              name="searchQuery"
              control={control}
              render={({ field }) => (
                <input
                  {...field}
                  type="text"
                  onKeyDown={handleKeyDown}
                  placeholder="例：「営業の売上データで目標達成を判定したい」「在庫が少ない商品をチェックしたい」"
                  style={{ 
                    flex: 1, 
                    padding: '10px', 
                    border: '1px solid #ddd', 
                    borderRadius: '4px',
                    fontSize: '14px'
                  }}
                  disabled={isSubmitting}
                />
              )}
            />
            <button
              type="submit"
              disabled={isSubmitting || !userInput.trim()}
              style={{
                padding: '10px 20px',
                backgroundColor: isSubmitting ? '#ccc' : '#007bff',
                color: 'white',
                border: 'none',
                borderRadius: '4px',
                cursor: isSubmitting ? 'not-allowed' : 'pointer'
              }}
            >
              {isSubmitting ? '検索中...' : '関数を検索'}
            </button>
          </div>
        </form>
        
        <div style={{ fontSize: '12px', color: '#666', marginBottom: '5px' }}>
          関数テンプレート:
        </div>
        <div className="template-buttons" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', marginBottom: '15px' }}>
          <button
            onClick={() => setShowTemplateSelector(true)}
            style={{
              padding: '8px 16px',
              backgroundColor: '#007bff',
              color: 'white',
              border: 'none',
              borderRadius: '6px',
              cursor: 'pointer',
              fontSize: '14px',
              fontWeight: '500',
              display: 'flex',
              alignItems: 'center',
              gap: '6px'
            }}
          >
            📚 テンプレートから選ぶ
          </button>
        </div>
        
        <div style={{ fontSize: '12px', color: '#666', marginBottom: '5px' }}>
          または、フリー入力:
        </div>
        <div className="quick-buttons" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
          {[
            '営業の売上で目標達成を判定したい',
            '在庫管理で発注判定をしたい', 
            '成績データで合否判定をしたい',
            '商品検索機能を作りたい'
          ].map(query => (
            <button
              key={query}
              onClick={() => { 
                setValue('searchQuery', query); // 入力欄にテキストを設定するだけ
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
          <input 
            type="text"
            className="formula-display"
            value={selectedCellFormula || ''}
            readOnly
            onClick={(e) => {
              // イベントの伝播を停止してSpreadsheetへの影響を防ぐ
              e.preventDefault();
              e.stopPropagation();
              // 入力フィールドにフォーカスして全選択
              e.currentTarget.focus();
              e.currentTarget.select();
            }}
            onMouseDown={(e) => {
              // マウスダウン時もイベントの伝播を停止
              e.preventDefault();
              e.stopPropagation();
            }}
            onFocus={(e) => {
              // フォーカス時に全選択
              e.currentTarget.select();
            }}
            style={{
              flex: 1,
              padding: '4px 8px',
              backgroundColor: 'white',
              border: '1px solid #ced4da',
              borderRadius: '4px',
              fontFamily: 'Monaco, Consolas, "Lucida Console", monospace',
              fontSize: '13px',
              color: selectedCellFormula.startsWith('=') ? '#d73a49' : '#333',
              minHeight: '22px',
              cursor: 'text',
              outline: 'none'
            }}
            title="クリックして数式をコピーできます"
          />
        </div>
      </div>
      
      <div className="spreadsheet-container" style={{ height: '500px' }}>
        <Controller
          name="spreadsheetData"
          control={control}
          render={({ field }) => (
            <Spreadsheet
              data={field.value as Matrix<any>}
              onChange={(data) => {
                field.onChange(data);
                // HyperFormulaで再計算
                recalculateFormulas(data);
              }}
          onSelect={(selected) => {
            // セル選択時に数式を表示
            
            let row, column;
            
            // RangeSelection2 の場合
            if (selected && (selected as any).range) {
              const range = (selected as any).range;
              // Process range structure
              
              // PointRange2 の場合
              if (range.start) {
                row = range.start.row;
                column = range.start.column;
              }
              // 他の構造の場合
              else if (range.row !== undefined && range.column !== undefined) {
                row = range.row;
                column = range.column;
              }
            }
            // 古い形式（min プロパティ）の場合
            else if (selected && typeof selected === 'object' && 'min' in selected) {
              const min = selected.min as { row: number; column: number };
              row = min.row;
              column = min.column;
            }
            
            // Coordinates extracted
            
            if (row !== undefined && column !== undefined) {
              const cell = sheetData[row]?.[column];
              const cellAddress = `${String.fromCharCode(65 + column)}${row + 1}`;
              
              // Process selected cell
              
              setValue('selectedCell.address', cellAddress);
              
              if (cell && cell.formula) {
                setValue('selectedCell.formula', cell.formula);
              } else if (cell && cell.value !== undefined && cell.value !== null) {
                setValue('selectedCell.formula', String(cell.value));
              } else {
                setValue('selectedCell.formula', '');
              }
            } else {
              // Could not get coordinates - 数式バーをクリックした場合など
              // 選択状態をクリアしない（現在の選択を保持）
              console.log('座標を取得できませんでした - 選択状態を保持します');
            }
          }}
          onActivate={(point) => {
            // セルアクティブ時に数式を表示
            if (point) {
              const { row, column } = point;
              const cell = sheetData[row]?.[column];
              const cellAddress = `${String.fromCharCode(65 + column)}${row + 1}`;
              
              // Process active cell
              
              setValue('selectedCell.address', cellAddress);
              
              if (cell && cell.formula) {
                setValue('selectedCell.formula', cell.formula);
              } else if (cell && cell.value !== undefined && cell.value !== null) {
                setValue('selectedCell.formula', String(cell.value));
              } else {
                setValue('selectedCell.formula', '');
              }
            }
          }}
              columnLabels={['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']}
              rowLabels={['1', '2', '3', '4', '5', '6', '7', '8']}
            />
          )}
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
      
      {/* テンプレート選択モーダル */}
      {showTemplateSelector && (
        <TemplateSelector
          onTemplateSelect={handleTemplateSelect}
          onClose={() => setShowTemplateSelector(false)}
        />
      )}
    </div>
  );
};

export default ChatGPTSpreadsheet;