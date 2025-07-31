// ISO.CEILING関数のデバッグ - getCellValue のシミュレーション

// テストデータ
const testContext = {
  data: [
    [{ value: '数値' }, { value: '基準値' }, { value: '切り上げ' }],
    [{ value: 4.3 }, { value: 1 }, { value: '' }],
    [{ value: -4.3 }, { value: 1 }, { value: '' }]
  ],
  row: 2,
  col: 2
};

// getCellValueのシミュレーション
function getCellValue(cellRef, context) {
  // A3の場合
  if (cellRef === 'A3') {
    const row = 2; // 3行目は配列の2番目
    const col = 0; // A列は配列の0番目
    const cellData = context.data[row][col];
    return cellData.value; // -4.3
  }
  // B3の場合
  if (cellRef === 'B3') {
    const row = 2;
    const col = 1;
    const cellData = context.data[row][col];
    return cellData.value; // 1
  }
  return null;
}

// toNumberのシミュレーション
function toNumber(value) {
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const num = parseFloat(value);
    if (!isNaN(num)) return num;
  }
  return null;
}

// ISO.CEILING(A3,B3)のテスト
const numberRef = 'A3';
const significanceRef = 'B3';

const numberValue = getCellValue(numberRef, testContext);
console.log('numberValue:', numberValue, typeof numberValue);

const number = toNumber(numberValue) ?? parseFloat(numberRef);
console.log('number:', number);

const significanceValue = getCellValue(significanceRef, testContext);
console.log('significanceValue:', significanceValue, typeof significanceValue);

const significance = toNumber(significanceValue) ?? parseFloat(significanceRef);
console.log('significance:', significance);

// 計算
if (number < 0) {
  const result = Math.ceil(number / significance) * significance;
  console.log('Result (negative):', result);
} else {
  const result = Math.ceil(number / Math.abs(significance)) * Math.abs(significance);
  console.log('Result (positive):', result);
}