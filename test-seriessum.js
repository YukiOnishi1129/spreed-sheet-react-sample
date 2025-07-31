// SERIESSUM関数のテスト
// =SERIESSUM(A2,B2,C2,D2:F2) where A2=2, B2=1, C2=1, D2:F2=[1,2,3]

// 手動計算
const x = 2;
const n = 1;
const m = 1;
const coefficients = [1, 2, 3];

let sum = 0;
for (let i = 0; i < coefficients.length; i++) {
  const power = n + i * m;
  const term = coefficients[i] * Math.pow(x, power);
  console.log(`係数${i+1}: ${coefficients[i]} * ${x}^${power} = ${term}`);
  sum += term;
}

console.log(`合計: ${sum}`);
console.log(`期待値: 22`);

// 期待値22になる組み合わせを探す
console.log('\n別の可能性:');
// もし係数が[1, 5]なら
const coeff2 = [1, 5];
let sum2 = 0;
for (let i = 0; i < coeff2.length; i++) {
  const power = n + i * m;
  const term = coeff2[i] * Math.pow(x, power);
  console.log(`係数${i+1}: ${coeff2[i]} * ${x}^${power} = ${term}`);
  sum2 += term;
}
console.log(`合計: ${sum2}`);