import React from 'react';
import { calculateSingleFormula } from './src/components/utils/customFormulaCalculations';

// テストコンポーネント
function TestISO_CEILING() {
  const data = [
    [{ value: '数値' }, { value: '基準値' }, { value: '切り上げ' }],
    [{ value: 4.3 }, { value: 1 }, { value: '' }],
    [{ value: -4.3 }, { value: 1 }, { value: '' }]
  ];

  // Test positive number
  const result1 = calculateSingleFormula('=ISO.CEILING(A2,B2)', data, 1, 2);
  console.log('ISO.CEILING(4.3, 1) =', result1);

  // Test negative number  
  const result2 = calculateSingleFormula('=ISO.CEILING(A3,B3)', data, 2, 2);
  console.log('ISO.CEILING(-4.3, 1) =', result2);

  return (
    <div>
      <h1>ISO.CEILING Test</h1>
      <p>ISO.CEILING(4.3, 1) = {result1}</p>
      <p>ISO.CEILING(-4.3, 1) = {result2}</p>
      <p>Expected: 5 and -4</p>
    </div>
  );
}

// Run the test
const test = new TestISO_CEILING();
console.log('Test complete');