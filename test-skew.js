// SKEW関数の計算確認
const values = [3, 4, 5, 2, 1];

// Excelの SKEW 関数の計算式
// SKEW = n/((n-1)(n-2)) * Σ((xi - x̄)/s)^3
// ここで、n = データ数、x̄ = 平均、s = 標準偏差（標本）

const n = values.length;
const mean = values.reduce((sum, val) => sum + val, 0) / n;
console.log('平均:', mean);

// 標本分散と標準偏差
const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / (n - 1);
const stdDev = Math.sqrt(variance);
console.log('標本分散:', variance);
console.log('標本標準偏差:', stdDev);

// 各値の標準化とその3乗
console.log('\n各値の計算:');
let sumCubed = 0;
values.forEach(val => {
  const standardized = (val - mean) / stdDev;
  const cubed = Math.pow(standardized, 3);
  console.log(`値${val}: 標準化=${standardized.toFixed(4)}, 3乗=${cubed.toFixed(4)}`);
  sumCubed += cubed;
});

console.log('\n3乗の合計:', sumCubed);

// SKEW計算
const skew = (n / ((n - 1) * (n - 2))) * sumCubed;
console.log('SKEW:', skew);
console.log('SKEW（小数点3桁）:', skew.toFixed(3));
console.log('期待値: 0.406');

// 検証：値の順序を変えてみる
const values2 = [1, 2, 3, 4, 5];
const mean2 = values2.reduce((sum, val) => sum + val, 0) / n;
const variance2 = values2.reduce((sum, val) => sum + Math.pow(val - mean2, 2), 0) / (n - 1);
const stdDev2 = Math.sqrt(variance2);
let sumCubed2 = 0;
values2.forEach(val => {
  const standardized = (val - mean2) / stdDev2;
  sumCubed2 += Math.pow(standardized, 3);
});
const skew2 = (n / ((n - 1) * (n - 2))) * sumCubed2;
console.log('\n順序を変えた場合のSKEW:', skew2.toFixed(3));