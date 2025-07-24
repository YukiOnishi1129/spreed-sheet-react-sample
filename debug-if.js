// Debug IF function matching
const { matchFormula } = require('./src/utils/spreadsheet/logic');

const testFormulas = [
  '=IF(B2>=100000,"目標達成","要改善")',
  'IF(B2>=100000,"目標達成","要改善")'
];

testFormulas.forEach(formula => {
  console.log(`\nTesting: ${formula}`);
  const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;
  console.log(`Clean formula: ${cleanFormula}`);
  
  // Test pattern matching
  const pattern = /IF\s*\((.+)\)/i;
  const match = cleanFormula.match(pattern);
  
  if (match) {
    console.log('Pattern matched!');
    console.log('Full match:', match[0]);
    console.log('Arguments:', match[1]);
  } else {
    console.log('Pattern did NOT match');
  }
});