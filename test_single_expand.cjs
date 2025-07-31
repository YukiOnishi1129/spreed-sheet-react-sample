// EXPAND関数の実装を直接テストする
const { createFormulas } = require('./src/utils/spreadsheet/formulaParser.ts');

// テストデータ
const testData = {
  A1: { value: 1 },
  B1: { value: 2 },
  A2: { value: 3 },
  B2: { value: 4 }
};

// EXPAND関数をテスト
try {
  const formulas = createFormulas();
  const formula = '=EXPAND(A1:B2,4,3,"-")';
  
  console.log('Testing EXPAND function');
  console.log('Formula:', formula);
  console.log('Input data:', testData);
  
  // 簡易的なテスト実行
  // 注: 実際の実装では、Spreadsheetコンポーネントで処理される
  
} catch (error) {
  console.error('Error:', error);
}