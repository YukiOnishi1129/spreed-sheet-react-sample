import { describe, it, expect } from 'vitest';
import {
  BIN2DEC, BIN2HEX, BIN2OCT, DEC2BIN, DEC2HEX, DEC2OCT,
  HEX2BIN, HEX2DEC, HEX2OCT, OCT2BIN, OCT2DEC, OCT2HEX,
  CONVERT, DELTA, GESTEP
} from '../engineeringLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Engineering Conversion Functions', () => {
  const mockContext = createContext([
    ['1010', '1111', '10101010'], // binary values
    [10, 15, 255, -128, 511], // decimal values
    ['A', 'F', 'FF', '1FF'], // hex values
    ['7', '10', '377', '777'], // octal values
    [100, 32, 0, -40], // temperatures
    ['m', 'ft', 'kg', 'lbm'], // units
  ]);

  describe('BIN2DEC Function (Binary to Decimal)', () => {
    it('should convert positive binary to decimal', () => {
      const matches = ['BIN2DEC("1010")', '"1010"'] as RegExpMatchArray;
      const result = BIN2DEC.calculate(matches, mockContext);
      expect(result).toBe(10);
    });

    it('should convert negative binary to decimal (two\'s complement)', () => {
      const matches = ['BIN2DEC("1111111111")', '"1111111111"'] as RegExpMatchArray;
      const result = BIN2DEC.calculate(matches, mockContext);
      expect(result).toBe(-1); // 10-bit two's complement
    });

    it('should handle maximum positive value', () => {
      const matches = ['BIN2DEC("111111111")', '"111111111"'] as RegExpMatchArray;
      const result = BIN2DEC.calculate(matches, mockContext);
      expect(result).toBe(511); // 2^9 - 1
    });

    it('should handle minimum negative value', () => {
      const matches = ['BIN2DEC("1000000000")', '"1000000000"'] as RegExpMatchArray;
      const result = BIN2DEC.calculate(matches, mockContext);
      expect(result).toBe(-512); // -2^9
    });

    it('should return NUM error for invalid binary', () => {
      const matches = ['BIN2DEC("1012")', '"1012"'] as RegExpMatchArray;
      const result = BIN2DEC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for > 10 characters', () => {
      const matches = ['BIN2DEC("10101010101")', '"10101010101"'] as RegExpMatchArray;
      const result = BIN2DEC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should handle cell references', () => {
      const matches = ['BIN2DEC(A1)', 'A1'] as RegExpMatchArray;
      const result = BIN2DEC.calculate(matches, mockContext);
      expect(result).toBe(10); // "1010" -> 10
    });
  });

  describe('BIN2HEX Function (Binary to Hexadecimal)', () => {
    it('should convert binary to hex', () => {
      const matches = ['BIN2HEX("1111")', '"1111"'] as RegExpMatchArray;
      const result = BIN2HEX.calculate(matches, mockContext);
      expect(result).toBe('F');
    });

    it('should pad with zeros when places specified', () => {
      const matches = ['BIN2HEX("1111", 3)', '"1111"', '3'] as RegExpMatchArray;
      const result = BIN2HEX.calculate(matches, mockContext);
      expect(result).toBe('00F');
    });

    it('should handle negative binary (two\'s complement)', () => {
      const matches = ['BIN2HEX("1111111111")', '"1111111111"'] as RegExpMatchArray;
      const result = BIN2HEX.calculate(matches, mockContext);
      expect(result).toBe('FFFFFFFFFF'); // -1 in 40-bit hex
    });

    it('should return NUM error if places too small', () => {
      const matches = ['BIN2HEX("11111111", 1)', '"11111111"', '1'] as RegExpMatchArray;
      const result = BIN2HEX.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('BIN2OCT Function (Binary to Octal)', () => {
    it('should convert binary to octal', () => {
      const matches = ['BIN2OCT("111")', '"111"'] as RegExpMatchArray;
      const result = BIN2OCT.calculate(matches, mockContext);
      expect(result).toBe('7');
    });

    it('should handle larger binary numbers', () => {
      const matches = ['BIN2OCT("111111111")', '"111111111"'] as RegExpMatchArray;
      const result = BIN2OCT.calculate(matches, mockContext);
      expect(result).toBe('777');
    });

    it('should handle negative binary', () => {
      const matches = ['BIN2OCT("1111111111")', '"1111111111"'] as RegExpMatchArray;
      const result = BIN2OCT.calculate(matches, mockContext);
      expect(result).toBe('7777777777'); // -1 in octal
    });
  });

  describe('DEC2BIN Function (Decimal to Binary)', () => {
    it('should convert positive decimal to binary', () => {
      const matches = ['DEC2BIN(10)', '10'] as RegExpMatchArray;
      const result = DEC2BIN.calculate(matches, mockContext);
      expect(result).toBe('1010');
    });

    it('should convert negative decimal to binary', () => {
      const matches = ['DEC2BIN(-1)', '-1'] as RegExpMatchArray;
      const result = DEC2BIN.calculate(matches, mockContext);
      expect(result).toBe('1111111111'); // Two's complement
    });

    it('should pad with zeros when places specified', () => {
      const matches = ['DEC2BIN(10, 8)', '10', '8'] as RegExpMatchArray;
      const result = DEC2BIN.calculate(matches, mockContext);
      expect(result).toBe('00001010');
    });

    it('should return NUM error for out of range', () => {
      const matches = ['DEC2BIN(512)', '512'] as RegExpMatchArray;
      const result = DEC2BIN.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for decimal < -512', () => {
      const matches = ['DEC2BIN(-513)', '-513'] as RegExpMatchArray;
      const result = DEC2BIN.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('DEC2HEX Function (Decimal to Hexadecimal)', () => {
    it('should convert positive decimal to hex', () => {
      const matches = ['DEC2HEX(255)', '255'] as RegExpMatchArray;
      const result = DEC2HEX.calculate(matches, mockContext);
      expect(result).toBe('FF');
    });

    it('should convert negative decimal to hex', () => {
      const matches = ['DEC2HEX(-1)', '-1'] as RegExpMatchArray;
      const result = DEC2HEX.calculate(matches, mockContext);
      expect(result).toBe('FFFFFFFFFF'); // 40-bit two's complement
    });

    it('should handle large positive numbers', () => {
      const matches = ['DEC2HEX(549755813887)', '549755813887'] as RegExpMatchArray;
      const result = DEC2HEX.calculate(matches, mockContext);
      expect(result).toBe('7FFFFFFFFF'); // Max 40-bit positive
    });

    it('should pad with zeros', () => {
      const matches = ['DEC2HEX(10, 4)', '10', '4'] as RegExpMatchArray;
      const result = DEC2HEX.calculate(matches, mockContext);
      expect(result).toBe('000A');
    });
  });

  describe('DEC2OCT Function (Decimal to Octal)', () => {
    it('should convert positive decimal to octal', () => {
      const matches = ['DEC2OCT(8)', '8'] as RegExpMatchArray;
      const result = DEC2OCT.calculate(matches, mockContext);
      expect(result).toBe('10');
    });

    it('should convert larger numbers', () => {
      const matches = ['DEC2OCT(511)', '511'] as RegExpMatchArray;
      const result = DEC2OCT.calculate(matches, mockContext);
      expect(result).toBe('777');
    });

    it('should handle negative numbers', () => {
      const matches = ['DEC2OCT(-1)', '-1'] as RegExpMatchArray;
      const result = DEC2OCT.calculate(matches, mockContext);
      expect(result).toBe('7777777777'); // 30-bit two's complement
    });
  });

  describe('HEX2BIN Function (Hexadecimal to Binary)', () => {
    it('should convert hex to binary', () => {
      const matches = ['HEX2BIN("F")', '"F"'] as RegExpMatchArray;
      const result = HEX2BIN.calculate(matches, mockContext);
      expect(result).toBe('1111');
    });

    it('should handle lowercase hex', () => {
      const matches = ['HEX2BIN("ff")', '"ff"'] as RegExpMatchArray;
      const result = HEX2BIN.calculate(matches, mockContext);
      expect(result).toBe('11111111');
    });

    it('should pad with zeros', () => {
      const matches = ['HEX2BIN("F", 8)', '"F"', '8'] as RegExpMatchArray;
      const result = HEX2BIN.calculate(matches, mockContext);
      expect(result).toBe('00001111');
    });

    it('should handle negative hex (40-bit)', () => {
      const matches = ['HEX2BIN("FFFFFFFFFF")', '"FFFFFFFFFF"'] as RegExpMatchArray;
      const result = HEX2BIN.calculate(matches, mockContext);
      expect(result).toBe('1111111111'); // -1 in binary
    });

    it('should return NUM error for invalid hex', () => {
      const matches = ['HEX2BIN("XYZ")', '"XYZ"'] as RegExpMatchArray;
      const result = HEX2BIN.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('HEX2DEC Function (Hexadecimal to Decimal)', () => {
    it('should convert hex to decimal', () => {
      const matches = ['HEX2DEC("FF")', '"FF"'] as RegExpMatchArray;
      const result = HEX2DEC.calculate(matches, mockContext);
      expect(result).toBe(255);
    });

    it('should handle negative hex', () => {
      const matches = ['HEX2DEC("FFFFFFFFFF")', '"FFFFFFFFFF"'] as RegExpMatchArray;
      const result = HEX2DEC.calculate(matches, mockContext);
      expect(result).toBe(-1);
    });

    it('should handle maximum positive hex', () => {
      const matches = ['HEX2DEC("7FFFFFFFFF")', '"7FFFFFFFFF"'] as RegExpMatchArray;
      const result = HEX2DEC.calculate(matches, mockContext);
      expect(result).toBe(549755813887); // 2^39 - 1
    });
  });

  describe('HEX2OCT Function (Hexadecimal to Octal)', () => {
    it('should convert hex to octal', () => {
      const matches = ['HEX2OCT("FF")', '"FF"'] as RegExpMatchArray;
      const result = HEX2OCT.calculate(matches, mockContext);
      expect(result).toBe('377');
    });

    it('should handle padding', () => {
      const matches = ['HEX2OCT("F", 3)', '"F"', '3'] as RegExpMatchArray;
      const result = HEX2OCT.calculate(matches, mockContext);
      expect(result).toBe('017');
    });
  });

  describe('OCT2BIN Function (Octal to Binary)', () => {
    it('should convert octal to binary', () => {
      const matches = ['OCT2BIN("7")', '"7"'] as RegExpMatchArray;
      const result = OCT2BIN.calculate(matches, mockContext);
      expect(result).toBe('111');
    });

    it('should handle larger octal numbers', () => {
      const matches = ['OCT2BIN("377")', '"377"'] as RegExpMatchArray;
      const result = OCT2BIN.calculate(matches, mockContext);
      expect(result).toBe('11111111');
    });

    it('should return NUM error for invalid octal', () => {
      const matches = ['OCT2BIN("8")', '"8"'] as RegExpMatchArray;
      const result = OCT2BIN.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('OCT2DEC Function (Octal to Decimal)', () => {
    it('should convert octal to decimal', () => {
      const matches = ['OCT2DEC("10")', '"10"'] as RegExpMatchArray;
      const result = OCT2DEC.calculate(matches, mockContext);
      expect(result).toBe(8);
    });

    it('should handle negative octal', () => {
      const matches = ['OCT2DEC("7777777777")', '"7777777777"'] as RegExpMatchArray;
      const result = OCT2DEC.calculate(matches, mockContext);
      expect(result).toBe(-1);
    });
  });

  describe('OCT2HEX Function (Octal to Hexadecimal)', () => {
    it('should convert octal to hex', () => {
      const matches = ['OCT2HEX("377")', '"377"'] as RegExpMatchArray;
      const result = OCT2HEX.calculate(matches, mockContext);
      expect(result).toBe('FF');
    });

    it('should handle padding', () => {
      const matches = ['OCT2HEX("10", 3)', '"10"', '3'] as RegExpMatchArray;
      const result = OCT2HEX.calculate(matches, mockContext);
      expect(result).toBe('008');
    });
  });

  describe('CONVERT Function (Unit Conversion)', () => {
    it('should convert meters to feet', () => {
      const matches = ['CONVERT(1, "m", "ft")', '1', '"m"', '"ft"'] as RegExpMatchArray;
      const result = CONVERT.calculate(matches, mockContext);
      expect(result).toBeCloseTo(3.28084, 4);
    });

    it('should convert Celsius to Fahrenheit', () => {
      const matches = ['CONVERT(0, "C", "F")', '0', '"C"', '"F"'] as RegExpMatchArray;
      const result = CONVERT.calculate(matches, mockContext);
      expect(result).toBe(32);
    });

    it('should convert with prefixes', () => {
      const matches = ['CONVERT(1, "km", "m")', '1', '"km"', '"m"'] as RegExpMatchArray;
      const result = CONVERT.calculate(matches, mockContext);
      expect(result).toBe(1000);
    });

    it('should handle binary prefixes', () => {
      const matches = ['CONVERT(1, "Gibyte", "byte")', '1', '"Gibyte"', '"byte"'] as RegExpMatchArray;
      const result = CONVERT.calculate(matches, mockContext);
      expect(result).toBe(1073741824); // 2^30
    });

    it('should return NA error for incompatible units', () => {
      const matches = ['CONVERT(1, "m", "kg")', '1', '"m"', '"kg"'] as RegExpMatchArray;
      const result = CONVERT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });

    it('should return NA error for unknown units', () => {
      const matches = ['CONVERT(1, "xyz", "m")', '1', '"xyz"', '"m"'] as RegExpMatchArray;
      const result = CONVERT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NA);
    });
  });

  describe('DELTA Function (Kronecker Delta)', () => {
    it('should return 1 for equal numbers', () => {
      const matches = ['DELTA(5, 5)', '5', '5'] as RegExpMatchArray;
      const result = DELTA.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should return 0 for different numbers', () => {
      const matches = ['DELTA(5, 4)', '5', '4'] as RegExpMatchArray;
      const result = DELTA.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should compare with 0 when second argument omitted', () => {
      const matches = ['DELTA(0)', '0'] as RegExpMatchArray;
      const result = DELTA.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should handle decimal comparisons', () => {
      const matches = ['DELTA(1.5, 1.5)', '1.5', '1.5'] as RegExpMatchArray;
      const result = DELTA.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should handle negative numbers', () => {
      const matches = ['DELTA(-5, -5)', '-5', '-5'] as RegExpMatchArray;
      const result = DELTA.calculate(matches, mockContext);
      expect(result).toBe(1);
    });
  });

  describe('GESTEP Function (Greater Than or Equal Step)', () => {
    it('should return 1 for number >= step', () => {
      const matches = ['GESTEP(5, 4)', '5', '4'] as RegExpMatchArray;
      const result = GESTEP.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should return 0 for number < step', () => {
      const matches = ['GESTEP(3, 4)', '3', '4'] as RegExpMatchArray;
      const result = GESTEP.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should return 1 for equal values', () => {
      const matches = ['GESTEP(5, 5)', '5', '5'] as RegExpMatchArray;
      const result = GESTEP.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should compare with 0 when step omitted', () => {
      const matches = ['GESTEP(1)', '1'] as RegExpMatchArray;
      const result = GESTEP.calculate(matches, mockContext);
      expect(result).toBe(1);
    });

    it('should handle negative numbers', () => {
      const matches = ['GESTEP(-3, -5)', '-3', '-5'] as RegExpMatchArray;
      const result = GESTEP.calculate(matches, mockContext);
      expect(result).toBe(1); // -3 >= -5
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should verify round-trip conversions', () => {
      // DEC -> BIN -> DEC
      const dec = 42;
      const binMatches = ['DEC2BIN(42)', '42'] as RegExpMatchArray;
      const bin = DEC2BIN.calculate(binMatches, mockContext);
      
      const decMatches = ['BIN2DEC("' + bin + '")', '"' + bin + '"'] as RegExpMatchArray;
      const result = BIN2DEC.calculate(decMatches, mockContext);
      
      expect(result).toBe(dec);
    });

    it('should handle temperature conversion chains', () => {
      // C -> F -> K (not directly supported, but shows concept)
      const cToFMatches = ['CONVERT(100, "C", "F")', '100', '"C"', '"F"'] as RegExpMatchArray;
      const fahrenheit = CONVERT.calculate(cToFMatches, mockContext);
      expect(fahrenheit).toBe(212); // Boiling point
    });

    it('should handle maximum values in conversions', () => {
      // Maximum positive in each system
      const maxBin = '111111111'; // 511
      const maxOct = '777'; // 511
      const maxHex = '1FF'; // 511
      
      const binDecMatches = ['BIN2DEC("' + maxBin + '")', '"' + maxBin + '"'] as RegExpMatchArray;
      const octDecMatches = ['OCT2DEC("' + maxOct + '")', '"' + maxOct + '"'] as RegExpMatchArray;
      const hexDecMatches = ['HEX2DEC("' + maxHex + '")', '"' + maxHex + '"'] as RegExpMatchArray;
      
      expect(BIN2DEC.calculate(binDecMatches, mockContext)).toBe(511);
      expect(OCT2DEC.calculate(octDecMatches, mockContext)).toBe(511);
      expect(HEX2DEC.calculate(hexDecMatches, mockContext)).toBe(511);
    });

    it('should handle cell references', () => {
      const matches = ['DEC2HEX(C2)', 'C2'] as RegExpMatchArray;
      const result = DEC2HEX.calculate(matches, mockContext);
      expect(result).toBe('FF'); // 255 -> FF
    });
  });
});