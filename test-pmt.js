// PMT関数の直接テスト

// 正規表現パターンのテスト
const pmtPattern = /PMT\(([^,)]+),\s*([^,)]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i;
const testFormula = 'PMT(B2,C2,-A2)';

console.log('Formula:', testFormula);
console.log('Pattern:', pmtPattern.toString());

const matches = testFormula.match(pmtPattern);
console.log('Matches:', matches);

if (matches) {
  console.log('- Full match:', matches[0]);
  console.log('- Rate (arg1):', matches[1]);
  console.log('- Nper (arg2):', matches[2]);
  console.log('- Pv (arg3):', matches[3]);
  console.log('- Fv (arg4):', matches[4]);
  console.log('- Type (arg5):', matches[5]);
}

// 他のパターンもテスト
const testFormulas = [
  'PMT(0.005,12,-100000)',
  'PMT(B2,C2,-A2)',
  'PMT(B2, C2, -A2)',
  'PMT(B2,C2,-A2,0)',
  'PMT(B2,C2,-A2,0,0)',
  'PMT(A1/12,B1*12,-C1)'
];

console.log('\n--- Testing various PMT formulas ---');
testFormulas.forEach(formula => {
  const m = formula.match(pmtPattern);
  console.log(`${formula}: ${m ? 'MATCH' : 'NO MATCH'}`);
});