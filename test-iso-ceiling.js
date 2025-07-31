// ISO.CEILING関数のテスト
const { ISO_CEILING } = require('./src/utils/spreadsheet/logic/01-math-trigonometry/mathLogic');

// モックのコンテキスト
const mockContext = {
  data: [
    [{ value: '数値' }, { value: '基準値' }, { value: '切り上げ' }],
    [{ value: 4.3 }, { value: 1 }, { value: '' }],
    [{ value: -4.3 }, { value: 1 }, { value: '' }]
  ],
  row: 2,
  col: 2
};

// getCellValueのモック
global.getCellValue = (ref, context) => {
  // A3の場合
  if (ref === 'A3') return -4.3;
  // B3の場合
  if (ref === 'B3') return 1;
  return null;
};

// テスト実行
const match = 'ISO.CEILING(A3,B3)'.match(/ISO\.CEILING\(([^,)]+)(?:,\s*([^)]+))?\)/i);
if (match) {
  console.log('Match result:', match);
  console.log('numberRef:', match[1]);
  console.log('significanceRef:', match[2]);
  
  try {
    const result = ISO_CEILING.calculate(match, mockContext);
    console.log('計算結果:', result);
  } catch (error) {
    console.error('エラー:', error);
  }
}

// 手動計算
console.log('\n手動計算:');
console.log('Math.ceil(-4.3 / 1) =', Math.ceil(-4.3 / 1));
console.log('Math.ceil(-4.3 / 1) * 1 =', Math.ceil(-4.3 / 1) * 1);