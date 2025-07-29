import { describe, it, expect } from 'vitest';
import {
  IF, AND, OR, NOT, IFS, XOR, TRUE, FALSE, IFERROR, IFNA
} from '../logicLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Logic Functions', () => {
  const mockContext = createContext([
    [true, false, 1, 0, 'TRUE', 'FALSE', '', 'text', 10, 5],
    [20, 15, 'Hello', 'World', 25, '#N/A', '#VALUE!', '#DIV/0!', null, undefined],
    [100, 50, 30, 40, 60, 70, 80, 90, 'Yes', 'No'],
    ['Apple', 'Banana', 'Cherry', 1, 2, 3, 0, -5, true, false],
    [0.5, 1.5, 2.5, 'TRUE', 'FALSE', '1', '0', '', 'data', 'test']
  ]);

  // セル参照をエミュレート
  mockContext.cells = {
    A1: { value: true }, A2: { value: 20 }, A3: { value: 100 }, A4: { value: 'Apple' }, A5: { value: 0.5 },
    B1: { value: false }, B2: { value: 15 }, B3: { value: 50 }, B4: { value: 'Banana' }, B5: { value: 1.5 },
    C1: { value: 1 }, C2: { value: 'Hello' }, C3: { value: 30 }, C4: { value: 'Cherry' }, C5: { value: 2.5 },
    D1: { value: 0 }, D2: { value: 'World' }, D3: { value: 40 }, D4: { value: 1 }, D5: { value: 'TRUE' },
    E1: { value: 'TRUE' }, E2: { value: 25 }, E3: { value: 60 }, E4: { value: 2 }, E5: { value: 'FALSE' },
    F1: { value: 'FALSE' }, F2: { value: '#N/A' }, F3: { value: 70 }, F4: { value: 3 }, F5: { value: '1' },
    G1: { value: '' }, G2: { value: '#VALUE!' }, G3: { value: 80 }, G4: { value: 0 }, G5: { value: '0' },
    H1: { value: 'text' }, H2: { value: '#DIV/0!' }, H3: { value: 90 }, H4: { value: -5 }, H5: { value: '' },
    I1: { value: 10 }, I2: { value: null }, I3: { value: 'Yes' }, I4: { value: true }, I5: { value: 'data' },
    J1: { value: 5 }, J2: { value: undefined }, J3: { value: 'No' }, J4: { value: false }, J5: { value: 'test' }
  };

  describe('IF Function', () => {
    it('should return true value when condition is true', () => {
      const matches = ['IF(A1, "Yes", "No")', 'A1, "Yes", "No"'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe('Yes');
    });

    it('should return false value when condition is false', () => {
      const matches = ['IF(B1, "Yes", "No")', 'B1, "Yes", "No"'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe('No');
    });

    it('should handle numeric conditions', () => {
      const matches = ['IF(C1, "Positive", "Zero")', 'C1, "Positive", "Zero"'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe('Positive');

      const matches2 = ['IF(D1, "Positive", "Zero")', 'D1, "Positive", "Zero"'] as RegExpMatchArray;
      const result2 = IF.calculate(matches2, mockContext);
      expect(result2).toBe('Zero');
    });

    it('should handle comparison operators', () => {
      const matches = ['IF(A2>B2, "A is greater", "B is greater")', 'A2>B2, "A is greater", "B is greater"'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe('A is greater');

      const matches2 = ['IF(A2<B2, "Less", "Greater or Equal")', 'A2<B2, "Less", "Greater or Equal"'] as RegExpMatchArray;
      const result2 = IF.calculate(matches2, mockContext);
      expect(result2).toBe('Greater or Equal');
    });

    it('should handle string comparisons', () => {
      const matches = ['IF(C2="Hello", "Match", "No Match")', 'C2="Hello", "Match", "No Match"'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe('Match');
    });

    it('should handle cell references in results', () => {
      const matches = ['IF(A1, A2, B2)', 'A1, A2, B2'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe(20);
    });

    it('should return VALUE error for invalid arguments', () => {
      const matches = ['IF(A1, "Yes")', 'A1, "Yes"'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('AND Function', () => {
    it('should return true when all arguments are true', () => {
      const matches = ['AND(A1, C1)', 'A1, C1'] as RegExpMatchArray;
      const result = AND.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false when any argument is false', () => {
      const matches = ['AND(A1, B1)', 'A1, B1'] as RegExpMatchArray;
      const result = AND.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should handle numeric values', () => {
      const matches = ['AND(C1, I1)', 'C1, I1'] as RegExpMatchArray;
      const result = AND.calculate(matches, mockContext);
      expect(result).toBe(true);

      const matches2 = ['AND(C1, D1)', 'C1, D1'] as RegExpMatchArray;
      const result2 = AND.calculate(matches2, mockContext);
      expect(result2).toBe(false);
    });

    it('should handle string values', () => {
      const matches = ['AND(E1, F5)', 'E1, F5'] as RegExpMatchArray;
      const result = AND.calculate(matches, mockContext);
      expect(result).toBe(true);

      const matches2 = ['AND(F1, G5)', 'F1, G5'] as RegExpMatchArray;
      const result2 = AND.calculate(matches2, mockContext);
      expect(result2).toBe(false);
    });

    it('should return VALUE error for no arguments', () => {
      const matches = ['AND()', ''] as RegExpMatchArray;
      const result = AND.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('OR Function', () => {
    it('should return true when any argument is true', () => {
      const matches = ['OR(A1, B1)', 'A1, B1'] as RegExpMatchArray;
      const result = OR.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false when all arguments are false', () => {
      const matches = ['OR(B1, D1)', 'B1, D1'] as RegExpMatchArray;
      const result = OR.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should handle multiple arguments', () => {
      const matches = ['OR(B1, D1, C1)', 'B1, D1, C1'] as RegExpMatchArray;
      const result = OR.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should handle string values', () => {
      const matches = ['OR(F1, E1)', 'F1, E1'] as RegExpMatchArray;
      const result = OR.calculate(matches, mockContext);
      expect(result).toBe(true);
    });
  });

  describe('NOT Function', () => {
    it('should return opposite of boolean value', () => {
      const matches = ['NOT(A1)', 'A1'] as RegExpMatchArray;
      const result = NOT.calculate(matches, mockContext);
      expect(result).toBe(false);

      const matches2 = ['NOT(B1)', 'B1'] as RegExpMatchArray;
      const result2 = NOT.calculate(matches2, mockContext);
      expect(result2).toBe(true);
    });

    it('should handle numeric values', () => {
      const matches = ['NOT(C1)', 'C1'] as RegExpMatchArray;
      const result = NOT.calculate(matches, mockContext);
      expect(result).toBe(false);

      const matches2 = ['NOT(D1)', 'D1'] as RegExpMatchArray;
      const result2 = NOT.calculate(matches2, mockContext);
      expect(result2).toBe(true);
    });

    it('should handle string values', () => {
      const matches = ['NOT(E1)', 'E1'] as RegExpMatchArray;
      const result = NOT.calculate(matches, mockContext);
      expect(result).toBe(false);

      const matches2 = ['NOT(G1)', 'G1'] as RegExpMatchArray;
      const result2 = NOT.calculate(matches2, mockContext);
      expect(result2).toBe(true);
    });
  });

  describe('IFS Function', () => {
    it('should return first matching condition result', () => {
      const matches = ['IFS(A2>B2, "A is greater", B2>A2, "B is greater")', 'A2>B2, "A is greater", B2>A2, "B is greater"'] as RegExpMatchArray;
      const result = IFS.calculate(matches, mockContext);
      expect(result).toBe('A is greater');
    });

    it('should continue to next condition if first fails', () => {
      const matches = ['IFS(B2>A2, "B is greater", A2>B2, "A is greater")', 'B2>A2, "B is greater", A2>B2, "A is greater"'] as RegExpMatchArray;
      const result = IFS.calculate(matches, mockContext);
      expect(result).toBe('A is greater');
    });

    it('should handle boolean conditions', () => {
      const matches = ['IFS(A1, "True", B1, "False")', 'A1, "True", B1, "False"'] as RegExpMatchArray;
      const result = IFS.calculate(matches, mockContext);
      expect(result).toBe('True');
    });

    it('should return NA error when no conditions match', () => {
      const matches = ['IFS(FALSE, "Never", 0, "Zero")', 'FALSE, "Never", 0, "Zero"'] as RegExpMatchArray;
      const result = IFS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });

    it('should handle cell references in results', () => {
      const matches = ['IFS(A1, A2, B1, B2)', 'A1, A2, B1, B2'] as RegExpMatchArray;
      const result = IFS.calculate(matches, mockContext);
      expect(result).toBe(20);
    });
  });

  describe('XOR Function', () => {
    it('should return true for odd number of true values', () => {
      const matches = ['XOR(TRUE, FALSE, FALSE)', 'TRUE, FALSE, FALSE'] as RegExpMatchArray;
      const result = XOR.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for even number of true values', () => {
      const matches = ['XOR(TRUE, TRUE)', 'TRUE, TRUE'] as RegExpMatchArray;
      const result = XOR.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should handle cell references', () => {
      const matches = ['XOR(A1, B1, C1)', 'A1, B1, C1'] as RegExpMatchArray;
      const result = XOR.calculate(matches, mockContext);
      expect(result).toBe(false); // Two true values (A1=true, C1=1)
    });

    it('should handle numeric values', () => {
      const matches = ['XOR(1, 0, 1)', '1, 0, 1'] as RegExpMatchArray;
      const result = XOR.calculate(matches, mockContext);
      expect(result).toBe(false); // Two true values
    });
  });

  describe('TRUE and FALSE Functions', () => {
    it('should return true for TRUE()', () => {
      const matches = ['TRUE()'] as RegExpMatchArray;
      const result = TRUE.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for FALSE()', () => {
      const matches = ['FALSE()'] as RegExpMatchArray;
      const result = FALSE.calculate(matches, mockContext);
      expect(result).toBe(false);
    });
  });

  describe('IFERROR Function', () => {
    it('should return value when no error', () => {
      const matches = ['IFERROR(A2, "Error")', 'A2, "Error"'] as RegExpMatchArray;
      const result = IFERROR.calculate(matches, mockContext);
      expect(result).toBe(20);
    });

    it('should return error value when error occurs', () => {
      const matches = ['IFERROR(G2, "Handle Error")', 'G2, "Handle Error"'] as RegExpMatchArray;
      const result = IFERROR.calculate(matches, mockContext);
      expect(result).toBe('Handle Error');
    });

    it('should handle cell references in error value', () => {
      const matches = ['IFERROR(H2, A2)', 'H2, A2'] as RegExpMatchArray;
      const result = IFERROR.calculate(matches, mockContext);
      expect(result).toBe(20);
    });

    it('should return VALUE error for invalid arguments', () => {
      const matches = ['IFERROR(A2)', 'A2'] as RegExpMatchArray;
      const result = IFERROR.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('IFNA Function', () => {
    it('should return value when no #N/A error', () => {
      const matches = ['IFNA(A2, "Not Available")', 'A2, "Not Available"'] as RegExpMatchArray;
      const result = IFNA.calculate(matches, mockContext);
      expect(result).toBe(20);
    });

    it('should return NA value when #N/A error occurs', () => {
      const matches = ['IFNA(F2, "Not Available")', 'F2, "Not Available"'] as RegExpMatchArray;
      const result = IFNA.calculate(matches, mockContext);
      expect(result).toBe('Not Available');
    });

    it('should not handle other errors', () => {
      const matches = ['IFNA(G2, "Not Available")', 'G2, "Not Available"'] as RegExpMatchArray;
      const result = IFNA.calculate(matches, mockContext);
      expect(result).toBe('#VALUE!'); // Should return the error, not the replacement
    });

    it('should handle cell references in NA value', () => {
      const matches = ['IFNA(F2, A2)', 'F2, A2'] as RegExpMatchArray;
      const result = IFNA.calculate(matches, mockContext);
      expect(result).toBe(20);
    });
  });

  describe('Edge Cases', () => {
    it('should handle null and undefined values', () => {
      const matches = ['IF(I2, "HasValue", "NoValue")', 'I2, "HasValue", "NoValue"'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe('NoValue');

      const matches2 = ['AND(I2, J2)', 'I2, J2'] as RegExpMatchArray;
      const result2 = AND.calculate(matches2, mockContext);
      expect(result2).toBe(false);
    });

    it('should handle empty string values', () => {
      const matches = ['IF(G1, "NotEmpty", "Empty")', 'G1, "NotEmpty", "Empty"'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe('Empty');
    });

    it('should handle complex nested conditions', () => {
      const matches = ['IF(AND(A1, C1), "Both True", "Not Both True")', 'AND(A1, C1), "Both True", "Not Both True"'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe('Both True');
    });

    it('should handle inequality operators', () => {
      const matches = ['IF(A2<>B2, "Different", "Same")', 'A2<>B2, "Different", "Same"'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe('Different');

      const matches2 = ['IF(A2!=B2, "Different", "Same")', 'A2!=B2, "Different", "Same"'] as RegExpMatchArray;
      const result2 = IF.calculate(matches2, mockContext);
      expect(result2).toBe('Different');
    });

    it('should handle greater than or equal to operators', () => {
      const matches = ['IF(A2>=B2, "Greater or Equal", "Less")', 'A2>=B2, "Greater or Equal", "Less"'] as RegExpMatchArray;
      const result = IF.calculate(matches, mockContext);
      expect(result).toBe('Greater or Equal');

      const matches2 = ['IF(B2<=A2, "Less or Equal", "Greater")', 'B2<=A2, "Less or Equal", "Greater"'] as RegExpMatchArray;
      const result2 = IF.calculate(matches2, mockContext);
      expect(result2).toBe('Less or Equal');
    });
  });
});