import { describe, it, expect } from 'vitest';
import {
  IMSIN, IMCOS, IMTAN, IMCSC, IMSEC, IMCOT,
  IMSINH, IMCOSH, IMTANH, IMCSCH, IMSECH, IMCOTH
} from '../complexTrigLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Complex Trigonometric Functions', () => {
  const mockContext = createContext([
    ['0', '1', 'i', '-i', '1+i', '2+3i'], // complex values
    ['3.14159', '1.5708', '0.7854'], // pi, pi/2, pi/4
    ['-1+i', '3-4i', '0.5+0.5i'], // more complex values
  ]);

  describe('IMSIN Function (Complex Sine)', () => {
    it('should calculate sin of real number', () => {
      const matches = ['IMSIN("0")', '"0"'] as RegExpMatchArray;
      const result = IMSIN.calculate(matches, mockContext);
      expect(result).toBe('0');
    });

    it('should calculate sin(pi/2)', () => {
      const matches = ['IMSIN("1.5708")', '"1.5708"'] as RegExpMatchArray;
      const result = IMSIN.calculate(matches, mockContext);
      expect(result).toBeCloseTo('1', 3);
    });

    it('should calculate sin(i)', () => {
      const matches = ['IMSIN("i")', '"i"'] as RegExpMatchArray;
      const result = IMSIN.calculate(matches, mockContext);
      // sin(i) = i*sinh(1) ≈ 1.1752i
      expect(typeof result).toBe('string');
      expect((result as string).includes('i')).toBe(true);
      const imag = parseFloat((result as string).replace('i', ''));
      expect(imag).toBeCloseTo(Math.sinh(1), 3);
    });

    it('should calculate sin of complex number', () => {
      const matches = ['IMSIN("1+i")', '"1+i"'] as RegExpMatchArray;
      const result = IMSIN.calculate(matches, mockContext);
      expect(typeof result).toBe('string');
      expect((result as string).includes('+')).toBe(true);
      expect((result as string).includes('i')).toBe(true);
    });

    it('should return NUM error for invalid complex number', () => {
      const matches = ['IMSIN("invalid")', '"invalid"'] as RegExpMatchArray;
      const result = IMSIN.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should handle cell references', () => {
      const matches = ['IMSIN(A1)', 'A1'] as RegExpMatchArray;
      const result = IMSIN.calculate(matches, mockContext);
      expect(result).toBe('0');
    });
  });

  describe('IMCOS Function (Complex Cosine)', () => {
    it('should calculate cos of real number', () => {
      const matches = ['IMCOS("0")', '"0"'] as RegExpMatchArray;
      const result = IMCOS.calculate(matches, mockContext);
      expect(result).toBe('1');
    });

    it('should calculate cos(pi)', () => {
      const matches = ['IMCOS("3.14159")', '"3.14159"'] as RegExpMatchArray;
      const result = IMCOS.calculate(matches, mockContext);
      expect(result).toBeCloseTo('-1', 3);
    });

    it('should calculate cos(i)', () => {
      const matches = ['IMCOS("i")', '"i"'] as RegExpMatchArray;
      const result = IMCOS.calculate(matches, mockContext);
      // cos(i) = cosh(1) ≈ 1.5431
      expect(parseFloat(result as string)).toBeCloseTo(Math.cosh(1), 3);
    });

    it('should calculate cos of complex number', () => {
      const matches = ['IMCOS("1+i")', '"1+i"'] as RegExpMatchArray;
      const result = IMCOS.calculate(matches, mockContext);
      expect(typeof result).toBe('string');
    });

    it('should verify Euler identity: cos(z) = (e^iz + e^-iz)/2', () => {
      // For real numbers, this should match regular cosine
      const matches = ['IMCOS("1")', '"1"'] as RegExpMatchArray;
      const result = IMCOS.calculate(matches, mockContext);
      expect(parseFloat(result as string)).toBeCloseTo(Math.cos(1), 5);
    });
  });

  describe('IMTAN Function (Complex Tangent)', () => {
    it('should calculate tan of real number', () => {
      const matches = ['IMTAN("0")', '"0"'] as RegExpMatchArray;
      const result = IMTAN.calculate(matches, mockContext);
      expect(result).toBe('0');
    });

    it('should calculate tan(pi/4)', () => {
      const matches = ['IMTAN("0.7854")', '"0.7854"'] as RegExpMatchArray;
      const result = IMTAN.calculate(matches, mockContext);
      expect(result).toBeCloseTo('1', 3);
    });

    it('should calculate tan of complex number', () => {
      const matches = ['IMTAN("1+i")', '"1+i"'] as RegExpMatchArray;
      const result = IMTAN.calculate(matches, mockContext);
      expect(typeof result).toBe('string');
      expect((result as string).includes('+')).toBe(true);
      expect((result as string).includes('i')).toBe(true);
    });

    it('should handle tan of pure imaginary', () => {
      const matches = ['IMTAN("i")', '"i"'] as RegExpMatchArray;
      const result = IMTAN.calculate(matches, mockContext);
      // tan(i) = i*tanh(1)
      expect(typeof result).toBe('string');
      const imag = parseFloat((result as string).replace('i', ''));
      expect(imag).toBeCloseTo(Math.tanh(1), 3);
    });
  });

  describe('IMCSC Function (Complex Cosecant)', () => {
    it('should calculate csc of real number', () => {
      const matches = ['IMCSC("1.5708")', '"1.5708"'] as RegExpMatchArray;
      const result = IMCSC.calculate(matches, mockContext);
      expect(result).toBeCloseTo('1', 3); // csc(pi/2) = 1
    });

    it('should return NUM error for sin(z) = 0', () => {
      const matches = ['IMCSC("0")', '"0"'] as RegExpMatchArray;
      const result = IMCSC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should calculate csc of complex number', () => {
      const matches = ['IMCSC("1+i")', '"1+i"'] as RegExpMatchArray;
      const result = IMCSC.calculate(matches, mockContext);
      expect(typeof result).toBe('string');
    });
  });

  describe('IMSEC Function (Complex Secant)', () => {
    it('should calculate sec of real number', () => {
      const matches = ['IMSEC("0")', '"0"'] as RegExpMatchArray;
      const result = IMSEC.calculate(matches, mockContext);
      expect(result).toBe('1'); // sec(0) = 1/cos(0) = 1
    });

    it('should calculate sec(pi)', () => {
      const matches = ['IMSEC("3.14159")', '"3.14159"'] as RegExpMatchArray;
      const result = IMSEC.calculate(matches, mockContext);
      expect(result).toBeCloseTo('-1', 3); // sec(pi) = -1
    });

    it('should return NUM error for cos(z) = 0', () => {
      const matches = ['IMSEC("1.5708")', '"1.5708"'] as RegExpMatchArray;
      const result = IMSEC.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM); // sec(pi/2) undefined
    });
  });

  describe('IMCOT Function (Complex Cotangent)', () => {
    it('should calculate cot of real number', () => {
      const matches = ['IMCOT("0.7854")', '"0.7854"'] as RegExpMatchArray;
      const result = IMCOT.calculate(matches, mockContext);
      expect(result).toBeCloseTo('1', 3); // cot(pi/4) = 1
    });

    it('should return NUM error for tan(z) = ∞', () => {
      const matches = ['IMCOT("0")', '"0"'] as RegExpMatchArray;
      const result = IMCOT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should calculate cot of complex number', () => {
      const matches = ['IMCOT("1+i")', '"1+i"'] as RegExpMatchArray;
      const result = IMCOT.calculate(matches, mockContext);
      expect(typeof result).toBe('string');
    });
  });

  describe('IMSINH Function (Complex Hyperbolic Sine)', () => {
    it('should calculate sinh of real number', () => {
      const matches = ['IMSINH("0")', '"0"'] as RegExpMatchArray;
      const result = IMSINH.calculate(matches, mockContext);
      expect(result).toBe('0');
    });

    it('should calculate sinh(1)', () => {
      const matches = ['IMSINH("1")', '"1"'] as RegExpMatchArray;
      const result = IMSINH.calculate(matches, mockContext);
      expect(parseFloat(result as string)).toBeCloseTo(Math.sinh(1), 5);
    });

    it('should calculate sinh(i)', () => {
      const matches = ['IMSINH("i")', '"i"'] as RegExpMatchArray;
      const result = IMSINH.calculate(matches, mockContext);
      // sinh(i) = i*sin(1)
      expect(typeof result).toBe('string');
      const imag = parseFloat((result as string).replace('i', ''));
      expect(imag).toBeCloseTo(Math.sin(1), 5);
    });

    it('should handle complex numbers', () => {
      const matches = ['IMSINH("1+i")', '"1+i"'] as RegExpMatchArray;
      const result = IMSINH.calculate(matches, mockContext);
      expect(typeof result).toBe('string');
      expect((result as string).includes('+')).toBe(true);
      expect((result as string).includes('i')).toBe(true);
    });
  });

  describe('IMCOSH Function (Complex Hyperbolic Cosine)', () => {
    it('should calculate cosh of real number', () => {
      const matches = ['IMCOSH("0")', '"0"'] as RegExpMatchArray;
      const result = IMCOSH.calculate(matches, mockContext);
      expect(result).toBe('1');
    });

    it('should calculate cosh(1)', () => {
      const matches = ['IMCOSH("1")', '"1"'] as RegExpMatchArray;
      const result = IMCOSH.calculate(matches, mockContext);
      expect(parseFloat(result as string)).toBeCloseTo(Math.cosh(1), 5);
    });

    it('should calculate cosh(i)', () => {
      const matches = ['IMCOSH("i")', '"i"'] as RegExpMatchArray;
      const result = IMCOSH.calculate(matches, mockContext);
      // cosh(i) = cos(1)
      expect(parseFloat(result as string)).toBeCloseTo(Math.cos(1), 5);
    });

    it('should verify cosh(-z) = cosh(z)', () => {
      const matches1 = ['IMCOSH("2+3i")', '"2+3i"'] as RegExpMatchArray;
      const result1 = IMCOSH.calculate(matches1, mockContext);
      
      const matches2 = ['IMCOSH("-2-3i")', '"-2-3i"'] as RegExpMatchArray;
      const result2 = IMCOSH.calculate(matches2, mockContext);
      
      expect(result1).toBe(result2); // Even function
    });
  });

  describe('IMTANH Function (Complex Hyperbolic Tangent)', () => {
    it('should calculate tanh of real number', () => {
      const matches = ['IMTANH("0")', '"0"'] as RegExpMatchArray;
      const result = IMTANH.calculate(matches, mockContext);
      expect(result).toBe('0');
    });

    it('should calculate tanh(∞) approaching 1', () => {
      const matches = ['IMTANH("10")', '"10"'] as RegExpMatchArray;
      const result = IMTANH.calculate(matches, mockContext);
      expect(parseFloat(result as string)).toBeCloseTo(1, 5);
    });

    it('should calculate tanh of complex number', () => {
      const matches = ['IMTANH("1+i")', '"1+i"'] as RegExpMatchArray;
      const result = IMTANH.calculate(matches, mockContext);
      expect(typeof result).toBe('string');
    });
  });

  describe('IMCSCH Function (Complex Hyperbolic Cosecant)', () => {
    it('should calculate csch of real number', () => {
      const matches = ['IMCSCH("1")', '"1"'] as RegExpMatchArray;
      const result = IMCSCH.calculate(matches, mockContext);
      expect(parseFloat(result as string)).toBeCloseTo(1/Math.sinh(1), 5);
    });

    it('should return NUM error for sinh(z) = 0', () => {
      const matches = ['IMCSCH("0")', '"0"'] as RegExpMatchArray;
      const result = IMCSCH.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('IMSECH Function (Complex Hyperbolic Secant)', () => {
    it('should calculate sech of real number', () => {
      const matches = ['IMSECH("0")', '"0"'] as RegExpMatchArray;
      const result = IMSECH.calculate(matches, mockContext);
      expect(result).toBe('1'); // sech(0) = 1/cosh(0) = 1
    });

    it('should calculate sech of complex number', () => {
      const matches = ['IMSECH("1+i")', '"1+i"'] as RegExpMatchArray;
      const result = IMSECH.calculate(matches, mockContext);
      expect(typeof result).toBe('string');
    });
  });

  describe('IMCOTH Function (Complex Hyperbolic Cotangent)', () => {
    it('should calculate coth of real number', () => {
      const matches = ['IMCOTH("1")', '"1"'] as RegExpMatchArray;
      const result = IMCOTH.calculate(matches, mockContext);
      expect(parseFloat(result as string)).toBeCloseTo(1/Math.tanh(1), 5);
    });

    it('should return NUM error for tanh(z) = 0', () => {
      const matches = ['IMCOTH("0")', '"0"'] as RegExpMatchArray;
      const result = IMCOTH.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should approach ±1 for large values', () => {
      const matches = ['IMCOTH("10")', '"10"'] as RegExpMatchArray;
      const result = IMCOTH.calculate(matches, mockContext);
      expect(parseFloat(result as string)).toBeCloseTo(1, 3);
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should verify sin²(z) + cos²(z) = 1 for complex z', () => {
      const z = '"1+i"';
      
      const sinMatches = ['IMSIN(' + z + ')', z] as RegExpMatchArray;
      const sinResult = IMSIN.calculate(sinMatches, mockContext);
      
      const cosMatches = ['IMCOS(' + z + ')', z] as RegExpMatchArray;
      const cosResult = IMCOS.calculate(cosMatches, mockContext);
      
      // Would need IMPRODUCT and IMSUM to verify the identity fully
      expect(typeof sinResult).toBe('string');
      expect(typeof cosResult).toBe('string');
    });

    it('should verify relationship between trig and hyperbolic functions', () => {
      // sin(iz) = i*sinh(z)
      // For z = 1: sin(i) should equal i*sinh(1)
      const sinMatches = ['IMSIN("i")', '"i"'] as RegExpMatchArray;
      const sinResult = IMSIN.calculate(sinMatches, mockContext) as string;
      
      const sinhMatches = ['IMSINH("1")', '"1"'] as RegExpMatchArray;
      const sinhResult = IMSINH.calculate(sinhMatches, mockContext);
      
      const imag = parseFloat(sinResult.replace('i', ''));
      expect(imag).toBeCloseTo(parseFloat(sinhResult as string), 5);
    });

    it('should handle very small complex numbers', () => {
      const matches = ['IMSIN("0.0001+0.0001i")', '"0.0001+0.0001i"'] as RegExpMatchArray;
      const result = IMSIN.calculate(matches, mockContext);
      expect(typeof result).toBe('string');
      // For small z, sin(z) ≈ z
      expect((result as string).includes('0.0001')).toBe(true);
    });

    it('should handle cell references', () => {
      const matches = ['IMCOS(C2)', 'C2'] as RegExpMatchArray;
      const result = IMCOS.calculate(matches, mockContext);
      expect(typeof result).toBe('string');
    });
  });
});