import { describe, it, expect } from 'vitest';
import type { FormulaContext } from '../../shared/types';
import {
  SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2,
  SINH, COSH, TANH, ASINH, ACOSH, ATANH,
  DEGREES, RADIANS, PI,
  CSC, SEC, COT, ACOT,
  CSCH, SECH, COTH, ACOTH
} from '../trigonometricLogic';
import { FormulaError } from '../../shared/types';

// Helper function to create FormulaContext
function createContext(data: any[][]): FormulaContext {
  return {
    data,
    row: 0,
    col: 0
  };
}

describe('Trigonometric Functions', () => {
  const mockContext = createContext([
    [0, Math.PI / 2, Math.PI, 0.5, 1, -1, 'text', 2, 3, 4, 45, 90, 180]  // Row 1 (A1-A13)
  ]);

  describe('Basic Trigonometric Functions', () => {
    describe('SIN', () => {
      it('should calculate sine of angle in radians', () => {
        const matches = ['SIN(A1)', 'A1'] as RegExpMatchArray;
        expect(SIN.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['SIN(A2)', 'A2'] as RegExpMatchArray;
        expect(SIN.calculate(matches2, mockContext)).toBeCloseTo(1);
        
        const matches3 = ['SIN(A3)', 'A3'] as RegExpMatchArray;
        expect(SIN.calculate(matches3, mockContext)).toBeCloseTo(0);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['SIN(A7)', 'A7'] as RegExpMatchArray;
        expect(SIN.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });

    describe('COS', () => {
      it('should calculate cosine of angle in radians', () => {
        const matches = ['COS(A1)', 'A1'] as RegExpMatchArray;
        expect(COS.calculate(matches, mockContext)).toBe(1);
        
        const matches2 = ['COS(A2)', 'A2'] as RegExpMatchArray;
        expect(COS.calculate(matches2, mockContext)).toBeCloseTo(0);
        
        const matches3 = ['COS(A3)', 'A3'] as RegExpMatchArray;
        expect(COS.calculate(matches3, mockContext)).toBeCloseTo(-1);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['COS(A7)', 'A7'] as RegExpMatchArray;
        expect(COS.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });

    describe('TAN', () => {
      it('should calculate tangent of angle in radians', () => {
        const matches = ['TAN(A1)', 'A1'] as RegExpMatchArray;
        expect(TAN.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['TAN(0.785398)', '0.785398'] as RegExpMatchArray;
        expect(TAN.calculate(matches2, mockContext)).toBeCloseTo(1); // tan(π/4) ≈ 1
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['TAN(A7)', 'A7'] as RegExpMatchArray;
        expect(TAN.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });
  });

  describe('Inverse Trigonometric Functions', () => {
    describe('ASIN', () => {
      it('should calculate arcsine', () => {
        const matches = ['ASIN(A1)', 'A1'] as RegExpMatchArray;
        expect(ASIN.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['ASIN(A5)', 'A5'] as RegExpMatchArray;
        expect(ASIN.calculate(matches2, mockContext)).toBeCloseTo(Math.PI / 2);
        
        const matches3 = ['ASIN(A6)', 'A6'] as RegExpMatchArray;
        expect(ASIN.calculate(matches3, mockContext)).toBeCloseTo(-Math.PI / 2);
      });

      it('should return NUM error for values outside [-1, 1]', () => {
        const matches = ['ASIN(A8)', 'A8'] as RegExpMatchArray;
        expect(ASIN.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });

      it('should return NUM error for non-numeric input', () => {
        const matches = ['ASIN(A7)', 'A7'] as RegExpMatchArray;
        expect(ASIN.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('ACOS', () => {
      it('should calculate arccosine', () => {
        const matches = ['ACOS(A1)', 'A1'] as RegExpMatchArray;
        expect(ACOS.calculate(matches, mockContext)).toBeCloseTo(Math.PI / 2);
        
        const matches2 = ['ACOS(A5)', 'A5'] as RegExpMatchArray;
        expect(ACOS.calculate(matches2, mockContext)).toBe(0);
        
        const matches3 = ['ACOS(A6)', 'A6'] as RegExpMatchArray;
        expect(ACOS.calculate(matches3, mockContext)).toBeCloseTo(Math.PI);
      });

      it('should return NUM error for values outside [-1, 1]', () => {
        const matches = ['ACOS(A8)', 'A8'] as RegExpMatchArray;
        expect(ACOS.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('ATAN', () => {
      it('should calculate arctangent', () => {
        const matches = ['ATAN(A1)', 'A1'] as RegExpMatchArray;
        expect(ATAN.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['ATAN(A5)', 'A5'] as RegExpMatchArray;
        expect(ATAN.calculate(matches2, mockContext)).toBeCloseTo(Math.PI / 4);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['ATAN(A7)', 'A7'] as RegExpMatchArray;
        expect(ATAN.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });

    describe('ATAN2', () => {
      it('should calculate atan2', () => {
        const matches = ['ATAN2(A5, A5)', 'A5', 'A5'] as RegExpMatchArray;
        expect(ATAN2.calculate(matches, mockContext)).toBeCloseTo(Math.PI / 4);
        
        const matches2 = ['ATAN2(A1, A5)', 'A1', 'A5'] as RegExpMatchArray;
        expect(ATAN2.calculate(matches2, mockContext)).toBe(0);
      });

      it('should return DIV0 error for both x and y being 0', () => {
        const matches = ['ATAN2(A1, A1)', 'A1', 'A1'] as RegExpMatchArray;
        expect(ATAN2.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['ATAN2(A7, A5)', 'A7', 'A5'] as RegExpMatchArray;
        expect(ATAN2.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });
  });

  describe('Hyperbolic Functions', () => {
    describe('SINH', () => {
      it('should calculate hyperbolic sine', () => {
        const matches = ['SINH(A1)', 'A1'] as RegExpMatchArray;
        expect(SINH.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['SINH(A5)', 'A5'] as RegExpMatchArray;
        expect(SINH.calculate(matches2, mockContext)).toBeCloseTo(Math.sinh(1));
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['SINH(A7)', 'A7'] as RegExpMatchArray;
        expect(SINH.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });

    describe('COSH', () => {
      it('should calculate hyperbolic cosine', () => {
        const matches = ['COSH(A1)', 'A1'] as RegExpMatchArray;
        expect(COSH.calculate(matches, mockContext)).toBe(1);
        
        const matches2 = ['COSH(A5)', 'A5'] as RegExpMatchArray;
        expect(COSH.calculate(matches2, mockContext)).toBeCloseTo(Math.cosh(1));
      });
    });

    describe('TANH', () => {
      it('should calculate hyperbolic tangent', () => {
        const matches = ['TANH(A1)', 'A1'] as RegExpMatchArray;
        expect(TANH.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['TANH(A5)', 'A5'] as RegExpMatchArray;
        expect(TANH.calculate(matches2, mockContext)).toBeCloseTo(Math.tanh(1));
      });
    });
  });

  describe('Inverse Hyperbolic Functions', () => {
    describe('ASINH', () => {
      it('should calculate inverse hyperbolic sine', () => {
        const matches = ['ASINH(A1)', 'A1'] as RegExpMatchArray;
        expect(ASINH.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['ASINH(A5)', 'A5'] as RegExpMatchArray;
        expect(ASINH.calculate(matches2, mockContext)).toBeCloseTo(Math.asinh(1));
      });
    });

    describe('ACOSH', () => {
      it('should calculate inverse hyperbolic cosine', () => {
        const matches = ['ACOSH(A5)', 'A5'] as RegExpMatchArray;
        expect(ACOSH.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['ACOSH(A8)', 'A8'] as RegExpMatchArray;
        expect(ACOSH.calculate(matches2, mockContext)).toBeCloseTo(Math.acosh(2));
      });

      it('should return NUM error for values < 1', () => {
        const matches = ['ACOSH(A4)', 'A4'] as RegExpMatchArray;
        expect(ACOSH.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('ATANH', () => {
      it('should calculate inverse hyperbolic tangent', () => {
        const matches = ['ATANH(A1)', 'A1'] as RegExpMatchArray;
        expect(ATANH.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['ATANH(A4)', 'A4'] as RegExpMatchArray;
        expect(ATANH.calculate(matches2, mockContext)).toBeCloseTo(Math.atanh(0.5));
      });

      it('should return NUM error for values outside (-1, 1)', () => {
        const matches = ['ATANH(A5)', 'A5'] as RegExpMatchArray;
        expect(ATANH.calculate(matches, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['ATANH(A6)', 'A6'] as RegExpMatchArray;
        expect(ATANH.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
      });
    });
  });

  describe('Angle Conversion Functions', () => {
    describe('DEGREES', () => {
      it('should convert radians to degrees', () => {
        const matches = ['DEGREES(A1)', 'A1'] as RegExpMatchArray;
        expect(DEGREES.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['DEGREES(A2)', 'A2'] as RegExpMatchArray;
        expect(DEGREES.calculate(matches2, mockContext)).toBeCloseTo(90);
        
        const matches3 = ['DEGREES(A3)', 'A3'] as RegExpMatchArray;
        expect(DEGREES.calculate(matches3, mockContext)).toBeCloseTo(180);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['DEGREES(A7)', 'A7'] as RegExpMatchArray;
        expect(DEGREES.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });

    describe('RADIANS', () => {
      it('should convert degrees to radians', () => {
        const matches = ['RADIANS(A1)', 'A1'] as RegExpMatchArray;
        expect(RADIANS.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['RADIANS(A12)', 'A12'] as RegExpMatchArray;
        expect(RADIANS.calculate(matches2, mockContext)).toBeCloseTo(Math.PI / 2);
        
        const matches3 = ['RADIANS(A13)', 'A13'] as RegExpMatchArray;
        expect(RADIANS.calculate(matches3, mockContext)).toBeCloseTo(Math.PI);
      });
    });
  });

  describe('Constants', () => {
    describe('PI', () => {
      it('should return PI value', () => {
        const matches = ['PI()'] as RegExpMatchArray;
        expect(PI.calculate(matches, mockContext)).toBe(Math.PI);
      });
    });
  });

  describe('Reciprocal Trigonometric Functions', () => {
    describe('CSC', () => {
      it('should calculate cosecant', () => {
        const matches = ['CSC(A2)', 'A2'] as RegExpMatchArray;
        expect(CSC.calculate(matches, mockContext)).toBeCloseTo(1); // csc(π/2) = 1
      });

      it('should return DIV0 error when sin is zero', () => {
        const matches = ['CSC(A1)', 'A1'] as RegExpMatchArray;
        expect(CSC.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
      });
    });

    describe('SEC', () => {
      it('should calculate secant', () => {
        const matches = ['SEC(A1)', 'A1'] as RegExpMatchArray;
        expect(SEC.calculate(matches, mockContext)).toBe(1); // sec(0) = 1
        
        const matches2 = ['SEC(A3)', 'A3'] as RegExpMatchArray;
        expect(SEC.calculate(matches2, mockContext)).toBeCloseTo(-1); // sec(π) = -1
      });

      it('should return DIV0 error when cos is zero', () => {
        const matches = ['SEC(A2)', 'A2'] as RegExpMatchArray;
        expect(SEC.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
      });
    });

    describe('COT', () => {
      it('should calculate cotangent', () => {
        const matches = ['COT(0.785398)', '0.785398'] as RegExpMatchArray;
        expect(COT.calculate(matches, mockContext)).toBeCloseTo(1); // cot(π/4) ≈ 1
      });

      it('should return DIV0 error when tan is zero', () => {
        const matches = ['COT(A1)', 'A1'] as RegExpMatchArray;
        expect(COT.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
      });
    });

    describe('ACOT', () => {
      it('should calculate inverse cotangent', () => {
        const matches = ['ACOT(A5)', 'A5'] as RegExpMatchArray;
        expect(ACOT.calculate(matches, mockContext)).toBeCloseTo(Math.PI / 4);
        
        const matches2 = ['ACOT(A1)', 'A1'] as RegExpMatchArray;
        expect(ACOT.calculate(matches2, mockContext)).toBeCloseTo(Math.PI / 2);
      });
    });
  });

  describe('Hyperbolic Reciprocal Functions', () => {
    describe('CSCH', () => {
      it('should calculate hyperbolic cosecant', () => {
        const matches = ['CSCH(A5)', 'A5'] as RegExpMatchArray;
        expect(CSCH.calculate(matches, mockContext)).toBeCloseTo(1 / Math.sinh(1));
      });

      it('should return DIV0 error for zero', () => {
        const matches = ['CSCH(A1)', 'A1'] as RegExpMatchArray;
        expect(CSCH.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
      });
    });

    describe('SECH', () => {
      it('should calculate hyperbolic secant', () => {
        const matches = ['SECH(A1)', 'A1'] as RegExpMatchArray;
        expect(SECH.calculate(matches, mockContext)).toBe(1);
        
        const matches2 = ['SECH(A5)', 'A5'] as RegExpMatchArray;
        expect(SECH.calculate(matches2, mockContext)).toBeCloseTo(1 / Math.cosh(1));
      });
    });

    describe('COTH', () => {
      it('should calculate hyperbolic cotangent', () => {
        const matches = ['COTH(A5)', 'A5'] as RegExpMatchArray;
        expect(COTH.calculate(matches, mockContext)).toBeCloseTo(1 / Math.tanh(1));
      });

      it('should return DIV0 error for zero', () => {
        const matches = ['COTH(A1)', 'A1'] as RegExpMatchArray;
        expect(COTH.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
      });
    });

    describe('ACOTH', () => {
      it('should calculate inverse hyperbolic cotangent', () => {
        const matches = ['ACOTH(A8)', 'A8'] as RegExpMatchArray;
        const expected = 0.5 * Math.log(3); // 0.5 * ln((2+1)/(2-1))
        expect(ACOTH.calculate(matches, mockContext)).toBeCloseTo(expected);
      });

      it('should return NUM error for values in [-1, 1]', () => {
        const matches = ['ACOTH(A4)', 'A4'] as RegExpMatchArray;
        expect(ACOTH.calculate(matches, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['ACOTH(A5)', 'A5'] as RegExpMatchArray;
        expect(ACOTH.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
      });
    });
  });
});