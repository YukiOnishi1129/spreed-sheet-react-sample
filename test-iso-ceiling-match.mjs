// ISO.CEILING関数のマッチングテスト
import { matchFormula, ISO_CEILING } from './src/utils/spreadsheet/logic/index.ts';

console.log('ISO_CEILING object:', ISO_CEILING);
console.log('Pattern:', ISO_CEILING.pattern);

// テストする数式
const formulas = [
  'ISO.CEILING(A3,B3)',
  'ISO.CEILING(-4.3,1)',
  'ISO.CEILING(-4.3, 1)',
  'ISO.CEILING(A3)',
];

formulas.forEach(formula => {
  console.log(`\nTesting: ${formula}`);
  const result = matchFormula(formula);
  if (result) {
    console.log('Match found!');
    console.log('Function name:', result.function.name);
    console.log('Matches:', result.matches);
  } else {
    console.log('No match found');
  }
});