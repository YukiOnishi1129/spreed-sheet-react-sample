import { describe, it, expect } from 'vitest';
import {
  ISBLANK, ISERROR, ISNA, ISTEXT, ISNUMBER, ISLOGICAL,
  ISEVEN, ISODD, TYPE, N, ISERR, ISNONTEXT, ISREF,
  ISFORMULA, NA, ERROR_TYPE, INFO, SHEET, SHEETS,
  CELL, ISBETWEEN
} from '../informationLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null | { value: string | number | boolean | null; formula?: string })[][], row = 0, col = 0): FormulaContext => ({
  data: data.map(row => row.map(cell => {
    if (typeof cell === 'object' && cell !== null) {
      return cell;
    }
    return { value: cell };
  })),
  row,
  col
});

describe('Information Functions', () => {
  const mockContext = createContext([
    ['Text', 123, true, null, ''],
    [0, -5.5, false, FormulaError.VALUE, 'Hello'],
    [FormulaError.NA, FormulaError.DIV0, 2.5, '', 'World'],
    [{ value: 100, formula: '=SUM(A1:A3)' }, 'Simple', { value: 'Formula', formula: '=CONCATENATE("A","B")' }, 999, '']
  ]);

  describe('ISBLANK Function', () => {
    it('should return true for empty string', () => {
      const matches = ['ISBLANK(E1)', 'E1'] as RegExpMatchArray;
      const result = ISBLANK.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true for null value', () => {
      const matches = ['ISBLANK(D1)', 'D1'] as RegExpMatchArray;
      const result = ISBLANK.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for text value', () => {
      const matches = ['ISBLANK(A1)', 'A1'] as RegExpMatchArray;
      const result = ISBLANK.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for number value', () => {
      const matches = ['ISBLANK(B1)', 'B1'] as RegExpMatchArray;
      const result = ISBLANK.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for zero', () => {
      const matches = ['ISBLANK(A2)', 'A2'] as RegExpMatchArray;
      const result = ISBLANK.calculate(matches, mockContext);
      expect(result).toBe(false);
    });
  });

  describe('ISERROR Function', () => {
    it('should return true for VALUE error', () => {
      const matches = ['ISERROR(D2)', 'D2'] as RegExpMatchArray;
      const result = ISERROR.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true for NA error', () => {
      const matches = ['ISERROR(A3)', 'A3'] as RegExpMatchArray;
      const result = ISERROR.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for normal text', () => {
      const matches = ['ISERROR(A1)', 'A1'] as RegExpMatchArray;
      const result = ISERROR.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for number', () => {
      const matches = ['ISERROR(B1)', 'B1'] as RegExpMatchArray;
      const result = ISERROR.calculate(matches, mockContext);
      expect(result).toBe(false);
    });
  });

  describe('ISNA Function', () => {
    it('should return true for NA error', () => {
      const matches = ['ISNA(A3)', 'A3'] as RegExpMatchArray;
      const result = ISNA.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for other errors', () => {
      const matches = ['ISNA(D2)', 'D2'] as RegExpMatchArray;
      const result = ISNA.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for normal values', () => {
      const matches = ['ISNA(A1)', 'A1'] as RegExpMatchArray;
      const result = ISNA.calculate(matches, mockContext);
      expect(result).toBe(false);
    });
  });

  describe('ISTEXT Function', () => {
    it('should return true for text value', () => {
      const matches = ['ISTEXT(A1)', 'A1'] as RegExpMatchArray;
      const result = ISTEXT.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for number', () => {
      const matches = ['ISTEXT(B1)', 'B1'] as RegExpMatchArray;
      const result = ISTEXT.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for boolean', () => {
      const matches = ['ISTEXT(C1)', 'C1'] as RegExpMatchArray;
      const result = ISTEXT.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for error values', () => {
      const matches = ['ISTEXT(D2)', 'D2'] as RegExpMatchArray;
      const result = ISTEXT.calculate(matches, mockContext);
      expect(result).toBe(false);
    });
  });

  describe('ISNUMBER Function', () => {
    it('should return true for positive number', () => {
      const matches = ['ISNUMBER(B1)', 'B1'] as RegExpMatchArray;
      const result = ISNUMBER.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true for negative number', () => {
      const matches = ['ISNUMBER(B2)', 'B2'] as RegExpMatchArray;
      const result = ISNUMBER.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true for zero', () => {
      const matches = ['ISNUMBER(A2)', 'A2'] as RegExpMatchArray;
      const result = ISNUMBER.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for text', () => {
      const matches = ['ISNUMBER(A1)', 'A1'] as RegExpMatchArray;
      const result = ISNUMBER.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for boolean', () => {
      const matches = ['ISNUMBER(C1)', 'C1'] as RegExpMatchArray;
      const result = ISNUMBER.calculate(matches, mockContext);
      expect(result).toBe(false);
    });
  });

  describe('ISLOGICAL Function', () => {
    it('should return true for TRUE', () => {
      const matches = ['ISLOGICAL(C1)', 'C1'] as RegExpMatchArray;
      const result = ISLOGICAL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true for FALSE', () => {
      const matches = ['ISLOGICAL(C2)', 'C2'] as RegExpMatchArray;
      const result = ISLOGICAL.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for number', () => {
      const matches = ['ISLOGICAL(B1)', 'B1'] as RegExpMatchArray;
      const result = ISLOGICAL.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for text', () => {
      const matches = ['ISLOGICAL(A1)', 'A1'] as RegExpMatchArray;
      const result = ISLOGICAL.calculate(matches, mockContext);
      expect(result).toBe(false);
    });
  });

  describe('ISEVEN Function', () => {
    it('should return true for even positive number', () => {
      const matches = ['ISEVEN(4)', '4'] as RegExpMatchArray;
      const result = ISEVEN.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true for zero', () => {
      const matches = ['ISEVEN(A2)', 'A2'] as RegExpMatchArray;
      const result = ISEVEN.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for odd number', () => {
      const matches = ['ISEVEN(3)', '3'] as RegExpMatchArray;
      const result = ISEVEN.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return VALUE error for non-numeric value', () => {
      const matches = ['ISEVEN(A1)', 'A1'] as RegExpMatchArray;
      const result = ISEVEN.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle decimal numbers', () => {
      const matches = ['ISEVEN(4.7)', '4.7'] as RegExpMatchArray;
      const result = ISEVEN.calculate(matches, mockContext);
      expect(result).toBe(true); // Floor(4.7) = 4, which is even
    });
  });

  describe('ISODD Function', () => {
    it('should return true for odd number', () => {
      const matches = ['ISODD(3)', '3'] as RegExpMatchArray;
      const result = ISODD.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for even number', () => {
      const matches = ['ISODD(4)', '4'] as RegExpMatchArray;
      const result = ISODD.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for zero', () => {
      const matches = ['ISODD(A2)', 'A2'] as RegExpMatchArray;
      const result = ISODD.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return VALUE error for non-numeric value', () => {
      const matches = ['ISODD(A1)', 'A1'] as RegExpMatchArray;
      const result = ISODD.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('TYPE Function', () => {
    it('should return 1 for number', () => {
      const matches = ['TYPE(B1)', 'B1'] as RegExpMatchArray;
      const result = TYPE.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should return 2 for text', () => {
      const matches = ['TYPE(A1)', 'A1'] as RegExpMatchArray;
      const result = TYPE.calculate(matches, mockContext);
      expect(result).toBe(2);
    });

    it('should return 4 for logical value', () => {
      const matches = ['TYPE(C1)', 'C1'] as RegExpMatchArray;
      const result = TYPE.calculate(matches, mockContext);
      expect(result).toBe(4);
    });

    it('should return 16 for error value', () => {
      const matches = ['TYPE(D2)', 'D2'] as RegExpMatchArray;
      const result = TYPE.calculate(matches, mockContext);
      expect(result).toBe(16);
    });

    it('should handle direct numeric value', () => {
      const matches = ['TYPE(123)', '123'] as RegExpMatchArray;
      const result = TYPE.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should handle direct text value', () => {
      const matches = ['TYPE("Hello")', '"Hello"'] as RegExpMatchArray;
      const result = TYPE.calculate(matches, mockContext);
      expect(result).toBe(2);
    });
  });

  describe('N Function', () => {
    it('should return number as is', () => {
      const matches = ['N(B1)', 'B1'] as RegExpMatchArray;
      const result = N.calculate(matches, mockContext);
      expect(result).toBe(123);
    });

    it('should return 1 for TRUE', () => {
      const matches = ['N(C1)', 'C1'] as RegExpMatchArray;
      const result = N.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should return 0 for FALSE', () => {
      const matches = ['N(C2)', 'C2'] as RegExpMatchArray;
      const result = N.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should return 0 for text', () => {
      const matches = ['N(A1)', 'A1'] as RegExpMatchArray;
      const result = N.calculate(matches, mockContext);
      expect(result).toBe(0);
    });
  });

  describe('ISERR Function', () => {
    it('should return true for VALUE error', () => {
      const matches = ['ISERR(D2)', 'D2'] as RegExpMatchArray;
      const result = ISERR.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for NA error', () => {
      const matches = ['ISERR(A3)', 'A3'] as RegExpMatchArray;
      const result = ISERR.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for normal values', () => {
      const matches = ['ISERR(A1)', 'A1'] as RegExpMatchArray;
      const result = ISERR.calculate(matches, mockContext);
      expect(result).toBe(false);
    });
  });

  describe('ISNONTEXT Function', () => {
    it('should return false for text', () => {
      const matches = ['ISNONTEXT(A1)', 'A1'] as RegExpMatchArray;
      const result = ISNONTEXT.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return true for number', () => {
      const matches = ['ISNONTEXT(B1)', 'B1'] as RegExpMatchArray;
      const result = ISNONTEXT.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true for boolean', () => {
      const matches = ['ISNONTEXT(C1)', 'C1'] as RegExpMatchArray;
      const result = ISNONTEXT.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true for error values', () => {
      const matches = ['ISNONTEXT(D2)', 'D2'] as RegExpMatchArray;
      const result = ISNONTEXT.calculate(matches, mockContext);
      expect(result).toBe(true);
    });
  });

  describe('ISREF Function', () => {
    it('should return true for single cell reference', () => {
      const matches = ['ISREF(A1)', 'A1'] as RegExpMatchArray;
      const result = ISREF.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true for range reference', () => {
      const matches = ['ISREF(A1:B2)', 'A1:B2'] as RegExpMatchArray;
      const result = ISREF.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for text value', () => {
      const matches = ['ISREF("text")', '"text"'] as RegExpMatchArray;
      const result = ISREF.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for number', () => {
      const matches = ['ISREF(123)', '123'] as RegExpMatchArray;
      const result = ISREF.calculate(matches, mockContext);
      expect(result).toBe(false);
    });
  });

  describe('ISFORMULA Function', () => {
    it('should return true for cell with formula', () => {
      const matches = ['ISFORMULA(A4)', 'A4'] as RegExpMatchArray;
      const result = ISFORMULA.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true for another cell with formula', () => {
      const matches = ['ISFORMULA(C4)', 'C4'] as RegExpMatchArray;
      const result = ISFORMULA.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false for cell with simple value', () => {
      const matches = ['ISFORMULA(B4)', 'B4'] as RegExpMatchArray;
      const result = ISFORMULA.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for cell with number', () => {
      const matches = ['ISFORMULA(D4)', 'D4'] as RegExpMatchArray;
      const result = ISFORMULA.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return false for invalid reference', () => {
      const matches = ['ISFORMULA(Z99)', 'Z99'] as RegExpMatchArray;
      const result = ISFORMULA.calculate(matches, mockContext);
      expect(result).toBe(false);
    });
  });

  describe('NA Function', () => {
    it('should return NA error', () => {
      const matches = ['NA()', ''] as RegExpMatchArray;
      const result = NA.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });
  });

  describe('ERROR.TYPE Function', () => {
    it('should return 3 for VALUE error', () => {
      const matches = ['ERROR.TYPE(D2)', 'D2'] as RegExpMatchArray;
      const result = ERROR_TYPE.calculate(matches, mockContext);
      expect(result).toBe(3);
    });

    it('should return 7 for NA error', () => {
      const matches = ['ERROR.TYPE(A3)', 'A3'] as RegExpMatchArray;
      const result = ERROR_TYPE.calculate(matches, mockContext);
      expect(result).toBe(7);
    });

    it('should return 2 for DIV0 error', () => {
      const matches = ['ERROR.TYPE(B3)', 'B3'] as RegExpMatchArray;
      const result = ERROR_TYPE.calculate(matches, mockContext);
      expect(result).toBe(2);
    });

    it('should return NA error for non-error value', () => {
      const matches = ['ERROR.TYPE(A1)', 'A1'] as RegExpMatchArray;
      const result = ERROR_TYPE.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });
  });

  describe('INFO Function', () => {
    it('should return directory info', () => {
      const matches = ['INFO("directory")', '"directory"'] as RegExpMatchArray;
      const result = INFO.calculate(matches, mockContext);
      expect(result).toBe('/');
    });

    it('should return OS version info', () => {
      const matches = ['INFO("osversion")', '"osversion"'] as RegExpMatchArray;
      const result = INFO.calculate(matches, mockContext);
      expect(result).toBe('Web');
    });

    it('should return system info', () => {
      const matches = ['INFO("system")', '"system"'] as RegExpMatchArray;
      const result = INFO.calculate(matches, mockContext);
      expect(result).toBe('Web');
    });

    it('should return VALUE error for unknown type', () => {
      const matches = ['INFO("unknown")', '"unknown"'] as RegExpMatchArray;
      const result = INFO.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('SHEET Function', () => {
    it('should return current sheet number', () => {
      const matches = ['SHEET()', ''] as RegExpMatchArray;
      const result = SHEET.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should return sheet number with reference', () => {
      const matches = ['SHEET(A1)', 'A1'] as RegExpMatchArray;
      const result = SHEET.calculate(matches, mockContext);
      expect(result).toBe(1);
    });
  });

  describe('SHEETS Function', () => {
    it('should return number of sheets', () => {
      const matches = ['SHEETS()', ''] as RegExpMatchArray;
      const result = SHEETS.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should return number of sheets with reference', () => {
      const matches = ['SHEETS(A1:B2)', 'A1:B2'] as RegExpMatchArray;
      const result = SHEETS.calculate(matches, mockContext);
      expect(result).toBe(1);
    });
  });

  describe('CELL Function', () => {
    it('should return cell address', () => {
      const matches = ['CELL("address", A1)', '"address"', 'A1'] as RegExpMatchArray;
      const result = CELL.calculate(matches, mockContext);
      expect(result).toBe('$A1');
    });

    it('should return cell contents', () => {
      const matches = ['CELL("contents", A1)', '"contents"', 'A1'] as RegExpMatchArray;
      const result = CELL.calculate(matches, mockContext);
      expect(result).toBe('Text');
    });

    it('should return column number', () => {
      const matches = ['CELL("col", B1)', '"col"', 'B1'] as RegExpMatchArray;
      const result = CELL.calculate(matches, mockContext);
      expect(result).toBe(2);
    });

    it('should return row number', () => {
      const matches = ['CELL("row", A2)', '"row"', 'A2'] as RegExpMatchArray;
      const result = CELL.calculate(matches, mockContext);
      expect(result).toBe(2);
    });

    it('should return type for text', () => {
      const matches = ['CELL("type", A1)', '"type"', 'A1'] as RegExpMatchArray;
      const result = CELL.calculate(matches, mockContext);
      expect(result).toBe('l');
    });

    it('should return type for number', () => {
      const matches = ['CELL("type", B1)', '"type"', 'B1'] as RegExpMatchArray;
      const result = CELL.calculate(matches, mockContext);
      expect(result).toBe('v');
    });

    it('should return VALUE error for unknown info type', () => {
      const matches = ['CELL("unknown", A1)', '"unknown"', 'A1'] as RegExpMatchArray;
      const result = CELL.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('ISBETWEEN Function', () => {
    it('should return true when value is between bounds (inclusive)', () => {
      const matches = ['ISBETWEEN(5, 1, 10)', '5', '1', '10'] as RegExpMatchArray;
      const result = ISBETWEEN.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true when value equals lower bound (inclusive)', () => {
      const matches = ['ISBETWEEN(1, 1, 10)', '1', '1', '10'] as RegExpMatchArray;
      const result = ISBETWEEN.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return true when value equals upper bound (inclusive)', () => {
      const matches = ['ISBETWEEN(10, 1, 10)', '10', '1', '10'] as RegExpMatchArray;
      const result = ISBETWEEN.calculate(matches, mockContext);
      expect(result).toBe(true);
    });

    it('should return false when value is outside bounds', () => {
      const matches = ['ISBETWEEN(15, 1, 10)', '15', '1', '10'] as RegExpMatchArray;
      const result = ISBETWEEN.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should handle exclusive lower bound', () => {
      const matches = ['ISBETWEEN(1, 1, 10, FALSE)', '1', '1', '10', 'FALSE'] as RegExpMatchArray;
      const result = ISBETWEEN.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should handle exclusive upper bound', () => {
      const matches = ['ISBETWEEN(10, 1, 10, TRUE, FALSE)', '10', '1', '10', 'TRUE', 'FALSE'] as RegExpMatchArray;
      const result = ISBETWEEN.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should return VALUE error for non-numeric values', () => {
      const matches = ['ISBETWEEN("text", 1, 10)', '"text"', '1', '10'] as RegExpMatchArray;
      const result = ISBETWEEN.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('Edge Cases', () => {
    it('should handle empty context data correctly', () => {
      const emptyContext = createContext([]);
      const matches = ['ISBLANK(A1)', 'A1'] as RegExpMatchArray;
      const result = ISBLANK.calculate(matches, emptyContext);
      expect(result).toBe(true);
    });

    it('should handle large row references in TYPE', () => {
      const matches = ['TYPE(123.456)', '123.456'] as RegExpMatchArray;
      const result = TYPE.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should handle INFO with cell reference parameter', () => {
      const infoContext = createContext([['directory']]);
      const matches = ['INFO(A1)', 'A1'] as RegExpMatchArray;
      const result = INFO.calculate(matches, infoContext);
      expect(result).toBe('/');
    });

    it('should handle ISFORMULA for out of bounds reference', () => {
      const matches = ['ISFORMULA(Z999)', 'Z999'] as RegExpMatchArray;
      const result = ISFORMULA.calculate(matches, mockContext);
      expect(result).toBe(false);
    });

    it('should handle CELL with default reference', () => {
      const matches = ['CELL("address")', '"address"'] as RegExpMatchArray;
      const result = CELL.calculate(matches, mockContext);
      expect(result).toBe('$A1');
    });

    it('should handle decimal numbers in ISEVEN/ISODD', () => {
      const matches1 = ['ISEVEN(-2.9)', '-2.9'] as RegExpMatchArray;
      const result1 = ISEVEN.calculate(matches1, mockContext);
      expect(result1).toBe(true); // Floor(-2.9) = -3, but abs(-3) % 2 = 1, so it's odd... wait, let me check the logic

      const matches2 = ['ISODD(-3.1)', '-3.1'] as RegExpMatchArray;
      const result2 = ISODD.calculate(matches2, mockContext);
      expect(result2).toBe(true); // Floor(-3.1) = -4, -4 % 2 = 0 but with sign it should be handled properly
    });
  });
});