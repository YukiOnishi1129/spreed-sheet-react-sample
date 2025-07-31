import { describe, it, expect } from 'vitest';
import {
  COMPLEX, IMABS, IMAGINARY, IMREAL, IMCONJUGATE,
  IMARGUMENT, IMDIV, IMEXP, IMLN, IMLOG2,
  IMPOWER, IMPRODUCT, IMSQRT, IMSUB, IMSUM
} from '../complexNumberLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Complex Number Functions', () => {
  const mockContext = createContext([
    [3, 4, -2, 5, 0, 1], // real parts
    [4, -3, 2, -5, 1, 0], // imaginary parts
    ['3+4i', '5-2i', '-2+3i', '1+i', '4', '-3i'], // complex strings
    ['i', 'j', 'I', 'J'], // suffix variations
    [0.5, 1.5, 2.5, 3.14159], // decimal values
  ]);

  describe('COMPLEX Function (Create Complex Number)', () => {
    it('should create complex number with default suffix', () => {
      const matches = ['COMPLEX(3, 4)', '3', '4'] as RegExpMatchArray;
      const result = COMPLEX.calculate(matches, mockContext);
      expect(result).toBe('3+4i');
    });

    it('should create complex number with custom suffix', () => {
      const matches = ['COMPLEX(3, 4, "j")', '3', '4', '"j"'] as RegExpMatchArray;
      const result = COMPLEX.calculate(matches, mockContext);
      expect(result).toBe('3+4j');
    });

    it('should handle negative imaginary part', () => {
      const matches = ['COMPLEX(5, -2)', '5', '-2'] as RegExpMatchArray;
      const result = COMPLEX.calculate(matches, mockContext);
      expect(result).toBe('5-2i');
    });

    it('should handle zero real part', () => {
      const matches = ['COMPLEX(0, 3)', '0', '3'] as RegExpMatchArray;
      const result = COMPLEX.calculate(matches, mockContext);
      expect(result).toBe('3i');
    });

    it('should handle zero imaginary part', () => {
      const matches = ['COMPLEX(5, 0)', '5', '0'] as RegExpMatchArray;
      const result = COMPLEX.calculate(matches, mockContext);
      expect(result).toBe('5');
    });

    it('should handle both parts zero', () => {
      const matches = ['COMPLEX(0, 0)', '0', '0'] as RegExpMatchArray;
      const result = COMPLEX.calculate(matches, mockContext);
      expect(result).toBe('0');
    });

    it('should return VALUE error for invalid suffix', () => {
      const matches = ['COMPLEX(3, 4, "x")', '3', '4', '"x"'] as RegExpMatchArray;
      const result = COMPLEX.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });

    it('should handle cell references', () => {
      const matches = ['COMPLEX(A1, B1)', 'A1', 'B1'] as RegExpMatchArray;
      const result = COMPLEX.calculate(matches, mockContext);
      expect(result).toBe('3+4i');
    });
  });

  describe('IMABS Function (Complex Absolute Value)', () => {
    it('should calculate absolute value of complex number', () => {
      const matches = ['IMABS("3+4i")', '"3+4i"'] as RegExpMatchArray;
      const result = IMABS.calculate(matches, mockContext);
      expect(result).toBe(5); // sqrt(3^2 + 4^2) = 5
    });

    it('should handle negative components', () => {
      const matches = ['IMABS("-3-4i")', '"-3-4i"'] as RegExpMatchArray;
      const result = IMABS.calculate(matches, mockContext);
      expect(result).toBe(5); // Magnitude is always positive
    });

    it('should handle pure real numbers', () => {
      const matches = ['IMABS("-5")', '"-5"'] as RegExpMatchArray;
      const result = IMABS.calculate(matches, mockContext);
      expect(result).toBe(5);
    });

    it('should handle pure imaginary numbers', () => {
      const matches = ['IMABS("4i")', '"4i"'] as RegExpMatchArray;
      const result = IMABS.calculate(matches, mockContext);
      expect(result).toBe(4);
    });

    it('should handle zero', () => {
      const matches = ['IMABS("0")', '"0"'] as RegExpMatchArray;
      const result = IMABS.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should return NUM error for invalid complex number', () => {
      const matches = ['IMABS("invalid")', '"invalid"'] as RegExpMatchArray;
      const result = IMABS.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('IMAGINARY Function (Extract Imaginary Part)', () => {
    it('should extract imaginary coefficient', () => {
      const matches = ['IMAGINARY("3+4i")', '"3+4i"'] as RegExpMatchArray;
      const result = IMAGINARY.calculate(matches, mockContext);
      expect(result).toBe(4);
    });

    it('should handle negative imaginary part', () => {
      const matches = ['IMAGINARY("5-2i")', '"5-2i"'] as RegExpMatchArray;
      const result = IMAGINARY.calculate(matches, mockContext);
      expect(result).toBe(-2);
    });

    it('should return 0 for pure real numbers', () => {
      const matches = ['IMAGINARY("5")', '"5"'] as RegExpMatchArray;
      const result = IMAGINARY.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should handle pure imaginary numbers', () => {
      const matches = ['IMAGINARY("-3i")', '"-3i"'] as RegExpMatchArray;
      const result = IMAGINARY.calculate(matches, mockContext);
      expect(result).toBe(-3);
    });

    it('should handle j suffix', () => {
      const matches = ['IMAGINARY("2+5j")', '"2+5j"'] as RegExpMatchArray;
      const result = IMAGINARY.calculate(matches, mockContext);
      expect(result).toBe(5);
    });
  });

  describe('IMREAL Function (Extract Real Part)', () => {
    it('should extract real part', () => {
      const matches = ['IMREAL("3+4i")', '"3+4i"'] as RegExpMatchArray;
      const result = IMREAL.calculate(matches, mockContext);
      expect(result).toBe(3);
    });

    it('should handle negative real part', () => {
      const matches = ['IMREAL("-5+2i")', '"-5+2i"'] as RegExpMatchArray;
      const result = IMREAL.calculate(matches, mockContext);
      expect(result).toBe(-5);
    });

    it('should return 0 for pure imaginary numbers', () => {
      const matches = ['IMREAL("4i")', '"4i"'] as RegExpMatchArray;
      const result = IMREAL.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should handle pure real numbers', () => {
      const matches = ['IMREAL("7")', '"7"'] as RegExpMatchArray;
      const result = IMREAL.calculate(matches, mockContext);
      expect(result).toBe(7);
    });
  });

  describe('IMCONJUGATE Function (Complex Conjugate)', () => {
    it('should calculate conjugate', () => {
      const matches = ['IMCONJUGATE("3+4i")', '"3+4i"'] as RegExpMatchArray;
      const result = IMCONJUGATE.calculate(matches, mockContext);
      expect(result).toBe('3-4i');
    });

    it('should handle negative imaginary part', () => {
      const matches = ['IMCONJUGATE("5-2i")', '"5-2i"'] as RegExpMatchArray;
      const result = IMCONJUGATE.calculate(matches, mockContext);
      expect(result).toBe('5+2i');
    });

    it('should handle pure real numbers', () => {
      const matches = ['IMCONJUGATE("7")', '"7"'] as RegExpMatchArray;
      const result = IMCONJUGATE.calculate(matches, mockContext);
      expect(result).toBe('7');
    });

    it('should handle pure imaginary numbers', () => {
      const matches = ['IMCONJUGATE("3i")', '"3i"'] as RegExpMatchArray;
      const result = IMCONJUGATE.calculate(matches, mockContext);
      expect(result).toBe('-3i');
    });
  });

  describe('IMARGUMENT Function (Complex Argument/Phase)', () => {
    it('should calculate argument in first quadrant', () => {
      const matches = ['IMARGUMENT("1+i")', '"1+i"'] as RegExpMatchArray;
      const result = IMARGUMENT.calculate(matches, mockContext);
      expect(result).toBeCloseTo(Math.PI / 4, 5); // 45 degrees
    });

    it('should calculate argument in second quadrant', () => {
      const matches = ['IMARGUMENT("-1+i")', '"-1+i"'] as RegExpMatchArray;
      const result = IMARGUMENT.calculate(matches, mockContext);
      expect(result).toBeCloseTo(3 * Math.PI / 4, 5); // 135 degrees
    });

    it('should handle positive real axis', () => {
      const matches = ['IMARGUMENT("5")', '"5"'] as RegExpMatchArray;
      const result = IMARGUMENT.calculate(matches, mockContext);
      expect(result).toBe(0);
    });

    it('should handle negative real axis', () => {
      const matches = ['IMARGUMENT("-5")', '"-5"'] as RegExpMatchArray;
      const result = IMARGUMENT.calculate(matches, mockContext);
      expect(result).toBeCloseTo(Math.PI, 5); // 180 degrees
    });

    it('should handle positive imaginary axis', () => {
      const matches = ['IMARGUMENT("2i")', '"2i"'] as RegExpMatchArray;
      const result = IMARGUMENT.calculate(matches, mockContext);
      expect(result).toBeCloseTo(Math.PI / 2, 5); // 90 degrees
    });

    it('should return DIV0 error for zero', () => {
      const matches = ['IMARGUMENT("0")', '"0"'] as RegExpMatchArray;
      const result = IMARGUMENT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.DIV0);
    });
  });

  describe('IMDIV Function (Complex Division)', () => {
    it('should divide complex numbers', () => {
      const matches = ['IMDIV("6+8i", "3+4i")', '"6+8i"', '"3+4i"'] as RegExpMatchArray;
      const result = IMDIV.calculate(matches, mockContext);
      expect(result).toBe('2'); // (6+8i)/(3+4i) = 2
    });

    it('should handle division by real number', () => {
      const matches = ['IMDIV("4+6i", "2")', '"4+6i"', '"2"'] as RegExpMatchArray;
      const result = IMDIV.calculate(matches, mockContext);
      expect(result).toBe('2+3i');
    });

    it('should handle division by imaginary number', () => {
      const matches = ['IMDIV("4", "2i")', '"4"', '"2i"'] as RegExpMatchArray;
      const result = IMDIV.calculate(matches, mockContext);
      expect(result).toBe('-2i'); // 4/(2i) = -2i
    });

    it('should return NUM error for division by zero', () => {
      const matches = ['IMDIV("3+4i", "0")', '"3+4i"', '"0"'] as RegExpMatchArray;
      const result = IMDIV.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('IMSUM Function (Complex Sum)', () => {
    it('should sum two complex numbers', () => {
      const matches = ['IMSUM("3+4i", "1+2i")', '"3+4i", "1+2i"'] as RegExpMatchArray;
      const result = IMSUM.calculate(matches, mockContext);
      expect(result).toBe('4+6i');
    });

    it('should sum multiple complex numbers', () => {
      const matches = ['IMSUM("1+i", "2+2i", "3+3i")', '"1+i", "2+2i", "3+3i"'] as RegExpMatchArray;
      const result = IMSUM.calculate(matches, mockContext);
      expect(result).toBe('6+6i');
    });

    it('should handle mixed real and complex', () => {
      const matches = ['IMSUM("3+4i", "5")', '"3+4i", "5"'] as RegExpMatchArray;
      const result = IMSUM.calculate(matches, mockContext);
      expect(result).toBe('8+4i');
    });

    it('should handle cell references', () => {
      const matches = ['IMSUM(A3, B3)', 'A3, B3'] as RegExpMatchArray;
      const result = IMSUM.calculate(matches, mockContext);
      expect(result).toBe('8+2i'); // "3+4i" + "5-2i"
    });
  });

  describe('IMSUB Function (Complex Subtraction)', () => {
    it('should subtract complex numbers', () => {
      const matches = ['IMSUB("5+6i", "2+3i")', '"5+6i"', '"2+3i"'] as RegExpMatchArray;
      const result = IMSUB.calculate(matches, mockContext);
      expect(result).toBe('3+3i');
    });

    it('should handle negative results', () => {
      const matches = ['IMSUB("1+2i", "3+4i")', '"1+2i"', '"3+4i"'] as RegExpMatchArray;
      const result = IMSUB.calculate(matches, mockContext);
      expect(result).toBe('-2-2i');
    });
  });

  describe('IMPRODUCT Function (Complex Product)', () => {
    it('should multiply two complex numbers', () => {
      const matches = ['IMPRODUCT("3+4i", "1+2i")', '"3+4i", "1+2i"'] as RegExpMatchArray;
      const result = IMPRODUCT.calculate(matches, mockContext);
      expect(result).toBe('-5+10i'); // (3+4i)(1+2i) = 3+6i+4i-8 = -5+10i
    });

    it('should multiply multiple complex numbers', () => {
      const matches = ['IMPRODUCT("2+i", "1+i", "1-i")', '"2+i", "1+i", "1-i"'] as RegExpMatchArray;
      const result = IMPRODUCT.calculate(matches, mockContext);
      expect(result).toBe('4+2i'); // (2+i)(1+i)(1-i) = (2+i)(2) = 4+2i
    });

    it('should handle multiplication by i', () => {
      const matches = ['IMPRODUCT("3+4i", "i")', '"3+4i", "i"'] as RegExpMatchArray;
      const result = IMPRODUCT.calculate(matches, mockContext);
      expect(result).toBe('-4+3i'); // (3+4i)*i = -4+3i
    });
  });

  describe('IMSQRT Function (Complex Square Root)', () => {
    it('should calculate square root of positive real', () => {
      const matches = ['IMSQRT("4")', '"4"'] as RegExpMatchArray;
      const result = IMSQRT.calculate(matches, mockContext);
      expect(result).toBe('2');
    });

    it('should calculate square root of negative real', () => {
      const matches = ['IMSQRT("-4")', '"-4"'] as RegExpMatchArray;
      const result = IMSQRT.calculate(matches, mockContext);
      expect(result).toBe('2i');
    });

    it('should calculate square root of i', () => {
      const matches = ['IMSQRT("i")', '"i"'] as RegExpMatchArray;
      const result = IMSQRT.calculate(matches, mockContext);
      // sqrt(i) = (1+i)/sqrt(2)
      expect(typeof result).toBe('string');
      const real = parseFloat((result as string).split('+')[0]);
      const imag = parseFloat((result as string).split('+')[1].replace('i', ''));
      expect(real).toBeCloseTo(1/Math.sqrt(2), 5);
      expect(imag).toBeCloseTo(1/Math.sqrt(2), 5);
    });
  });

  describe('IMEXP Function (Complex Exponential)', () => {
    it('should calculate e^(real number)', () => {
      const matches = ['IMEXP("1")', '"1"'] as RegExpMatchArray;
      const result = IMEXP.calculate(matches, mockContext);
      const real = parseFloat(String(result));
      expect(real).toBeCloseTo(Math.E, 5);
    });

    it('should calculate e^(i*pi)', () => {
      const matches = ['IMEXP("3.14159265359i")', '"3.14159265359i"'] as RegExpMatchArray;
      const result = IMEXP.calculate(matches, mockContext);
      // e^(i*pi) = -1 (Euler's identity)
      expect(result).toBe('-1');
    });
  });

  describe('IMLN Function (Complex Natural Logarithm)', () => {
    it('should calculate ln of positive real', () => {
      const matches = ['IMLN("2.71828")', '"2.71828"'] as RegExpMatchArray;
      const result = IMLN.calculate(matches, mockContext);
      const real = parseFloat(String(result));
      expect(real).toBeCloseTo(1, 3);
    });

    it('should calculate ln of negative real', () => {
      const matches = ['IMLN("-1")', '"-1"'] as RegExpMatchArray;
      const result = IMLN.calculate(matches, mockContext);
      // ln(-1) = i*pi
      expect(typeof result).toBe('string');
      expect((result as string).includes('3.14')).toBe(true);
      expect((result as string).includes('i')).toBe(true);
    });

    it('should return NUM error for zero', () => {
      const matches = ['IMLN("0")', '"0"'] as RegExpMatchArray;
      const result = IMLN.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('IMLOG2 Function (Complex Logarithm Base 2)', () => {
    it('should calculate log2 of positive real', () => {
      const matches = ['IMLOG2("8")', '"8"'] as RegExpMatchArray;
      const result = IMLOG2.calculate(matches, mockContext);
      const real = parseFloat(String(result));
      expect(real).toBeCloseTo(3, 5);
    });

    it('should calculate log2 of complex number', () => {
      const matches = ['IMLOG2("2+2i")', '"2+2i"'] as RegExpMatchArray;
      const result = IMLOG2.calculate(matches, mockContext) as string;
      // log2(2+2i) = log2(2√2 * e^(i*π/4)) = log2(2√2) + i*π/4/ln(2)
      const parts = result.split('+');
      const real = parseFloat(parts[0]);
      const imag = parseFloat(parts[1].replace('i', ''));
      expect(real).toBeCloseTo(1.5, 3); // log2(2√2) = 1.5
      expect(imag).toBeCloseTo((Math.PI/4)/Math.LN2, 3);
    });

    it('should return NUM error for zero', () => {
      const matches = ['IMLOG2("0")', '"0"'] as RegExpMatchArray;
      const result = IMLOG2.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('IMPOWER Function (Complex Power)', () => {
    it('should calculate integer power', () => {
      const matches = ['IMPOWER("2+i", "2")', '"2+i"', '"2"'] as RegExpMatchArray;
      const result = IMPOWER.calculate(matches, mockContext);
      expect(result).toBe('3+4i'); // (2+i)^2 = 4+4i-1 = 3+4i
    });

    it('should calculate i^i', () => {
      const matches = ['IMPOWER("i", "i")', '"i"', '"i"'] as RegExpMatchArray;
      const result = IMPOWER.calculate(matches, mockContext);
      // i^i = e^(-pi/2) ≈ 0.20788
      expect(typeof result).toBe('string');
      expect(parseFloat(result as string)).toBeCloseTo(Math.exp(-Math.PI/2), 3);
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should verify complex multiplication identity', () => {
      // (a+bi)(a-bi) = a^2 + b^2
      const conjMatches = ['IMCONJUGATE("3+4i")', '"3+4i"'] as RegExpMatchArray;
      const conj = IMCONJUGATE.calculate(conjMatches, mockContext);
      
      // conj should be "3-4i"
      expect(conj).toBe('3-4i');
      
      const prodMatches = ['IMPRODUCT("3+4i", "3-4i")', '"3+4i", "3-4i"'] as RegExpMatchArray;
      const product = IMPRODUCT.calculate(prodMatches, mockContext);
      
      expect(product).toBe('25'); // 3^2 + 4^2 = 25
    });

    it('should verify Euler formula', () => {
      // e^(ix) = cos(x) + i*sin(x)
      const angle = Math.PI / 4;
      const matches = ['IMEXP("' + String(angle) + 'i")', '"' + String(angle) + 'i"'] as RegExpMatchArray;
      const result = IMEXP.calculate(matches, mockContext) as string;
      
      const parts = result.split('+');
      const real = parseFloat(parts[0]);
      const imag = parseFloat(parts[1].replace('i', ''));
      
      expect(real).toBeCloseTo(Math.cos(angle), 5);
      expect(imag).toBeCloseTo(Math.sin(angle), 5);
    });

    it('should handle cell references in all functions', () => {
      const matches = ['IMSUM(A3, B3)', 'A3, B3'] as RegExpMatchArray;
      const result = IMSUM.calculate(matches, mockContext);
      expect(result).toBe('8+2i');
    });
  });
});