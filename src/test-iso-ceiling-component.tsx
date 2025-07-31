import React, { useEffect, useState } from 'react';
import { calculateSingleFormula } from './components/utils/customFormulaCalculations';
import type { CellData } from './utils/spreadsheet/logic';

export function TestISO_CEILING() {
  const [results, setResults] = useState<any[]>([]);

  useEffect(() => {
    // テストデータ
    const data: CellData[][] = [
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
    
    // Debug: Check getCellValue
    const testContext = { data, row: 2, col: 2 };
    
    setResults([
      { formula: 'ISO.CEILING(4.3, 1)', result: result1, expected: 5 },
      { formula: 'ISO.CEILING(-4.3, 1)', result: result2, expected: -4 }
    ]);
  }, []);

  return (
    <div style={{ padding: '20px' }}>
      <h1>ISO.CEILING Test</h1>
      {results.map((test, index) => (
        <div key={index} style={{ marginBottom: '10px' }}>
          <p>{test.formula} = {test.result} (expected: {test.expected})</p>
          <p style={{ color: test.result === test.expected ? 'green' : 'red' }}>
            {test.result === test.expected ? '✓ PASS' : '✗ FAIL'}
          </p>
        </div>
      ))}
    </div>
  );
}