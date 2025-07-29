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
function createContext(data: (string | number | boolean | null)[][]): FormulaContext {
  return {
    data: data.map(row => row.map(cell => ({ value: cell }))),
    row: 0,
    col: 0
  };
}

describe('Trigonometric Functions', () => {
  const mockContext = createContext([
    [0, Math.PI / 2, Math.PI, 0.5, 1, -1, 'text', 2, 3, 4, 45, 90, 180]  // Row 1 (A1-M1)
  ]);

  describe('Basic Trigonometric Functions', () => {
    describe('SIN', () => {
      it('should calculate sine of angle in radians', () => {
        const matches = ['SIN(A1)', 'A1'] as RegExpMatchArray;
        expect(SIN.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['SIN(B1)', 'B1'] as RegExpMatchArray;
        expect(SIN.calculate(matches2, mockContext)).toBeCloseTo(1);
        
        const matches3 = ['SIN(C1)', 'C1'] as RegExpMatchArray;
        expect(SIN.calculate(matches3, mockContext)).toBeCloseTo(0);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['SIN(G1)', 'G1'] as RegExpMatchArray;
        expect(SIN.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });

    describe('COS', () => {
      it('should calculate cosine of angle in radians', () => {
        const matches = ['COS(A1)', 'A1'] as RegExpMatchArray;
        expect(COS.calculate(matches, mockContext)).toBe(1);
        
        const matches2 = ['COS(B1)', 'B1'] as RegExpMatchArray;
        expect(COS.calculate(matches2, mockContext)).toBeCloseTo(0);
        
        const matches3 = ['COS(C1)', 'C1'] as RegExpMatchArray;
        expect(COS.calculate(matches3, mockContext)).toBeCloseTo(-1);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['COS(G1)', 'G1'] as RegExpMatchArray;
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
        const matches = ['TAN(G1)', 'G1'] as RegExpMatchArray;
        expect(TAN.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });
  });

  describe('Inverse Trigonometric Functions', () => {
    describe('ASIN', () => {
      it('should calculate arcsine', () => {
        const matches = ['ASIN(A1)', 'A1'] as RegExpMatchArray;
        expect(ASIN.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['ASIN(E1)', 'E1'] as RegExpMatchArray;
        expect(ASIN.calculate(matches2, mockContext)).toBeCloseTo(Math.PI / 2);
        
        const matches3 = ['ASIN(F1)', 'F1'] as RegExpMatchArray;
        expect(ASIN.calculate(matches3, mockContext)).toBeCloseTo(-Math.PI / 2);
      });

      it('should return NUM error for values outside [-1, 1]', () => {
        const matches = ['ASIN(H1)', 'H1'] as RegExpMatchArray;
        expect(ASIN.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });

      it('should return NUM error for non-numeric input', () => {
        const matches = ['ASIN(G1)', 'G1'] as RegExpMatchArray;
        expect(ASIN.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('ACOS', () => {
      it('should calculate arccosine', () => {
        const matches = ['ACOS(A1)', 'A1'] as RegExpMatchArray;
        expect(ACOS.calculate(matches, mockContext)).toBeCloseTo(Math.PI / 2);
        
        const matches2 = ['ACOS(E1)', 'E1'] as RegExpMatchArray;
        expect(ACOS.calculate(matches2, mockContext)).toBe(0);
        
        const matches3 = ['ACOS(F1)', 'F1'] as RegExpMatchArray;
        expect(ACOS.calculate(matches3, mockContext)).toBeCloseTo(Math.PI);
      });

      it('should return NUM error for values outside [-1, 1]', () => {
        const matches = ['ACOS(H1)', 'H1'] as RegExpMatchArray;
        expect(ACOS.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('ATAN', () => {
      it('should calculate arctangent', () => {
        const matches = ['ATAN(A1)', 'A1'] as RegExpMatchArray;
        expect(ATAN.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['ATAN(E1)', 'E1'] as RegExpMatchArray;
        expect(ATAN.calculate(matches2, mockContext)).toBeCloseTo(Math.PI / 4);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['ATAN(G1)', 'G1'] as RegExpMatchArray;
        expect(ATAN.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });

    describe('ATAN2', () => {
      it('should calculate atan2', () => {
        const matches = ['ATAN2(E1, E1)', 'E1', 'E1'] as RegExpMatchArray;
        expect(ATAN2.calculate(matches, mockContext)).toBeCloseTo(Math.PI / 4);
        
        const matches2 = ['ATAN2(A1, E1)', 'A1', 'E1'] as RegExpMatchArray;
        expect(ATAN2.calculate(matches2, mockContext)).toBeCloseTo(Math.PI / 2);
      });

      it('should return DIV0 error for both x and y being 0', () => {
        const matches = ['ATAN2(A1, A1)', 'A1', 'A1'] as RegExpMatchArray;
        expect(ATAN2.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['ATAN2(G1, E1)', 'G1', 'E1'] as RegExpMatchArray;
        expect(ATAN2.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });
  });

  describe('Hyperbolic Functions', () => {
    describe('SINH', () => {
      it('should calculate hyperbolic sine', () => {
        const matches = ['SINH(A1)', 'A1'] as RegExpMatchArray;
        expect(SINH.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['SINH(E1)', 'E1'] as RegExpMatchArray;
        expect(SINH.calculate(matches2, mockContext)).toBeCloseTo(Math.sinh(1));
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['SINH(G1)', 'G1'] as RegExpMatchArray;
        expect(SINH.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });

    describe('COSH', () => {
      it('should calculate hyperbolic cosine', () => {
        const matches = ['COSH(A1)', 'A1'] as RegExpMatchArray;
        expect(COSH.calculate(matches, mockContext)).toBe(1);
        
        const matches2 = ['COSH(E1)', 'E1'] as RegExpMatchArray;
        expect(COSH.calculate(matches2, mockContext)).toBeCloseTo(Math.cosh(1));
      });
    });

    describe('TANH', () => {
      it('should calculate hyperbolic tangent', () => {
        const matches = ['TANH(A1)', 'A1'] as RegExpMatchArray;
        expect(TANH.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['TANH(E1)', 'E1'] as RegExpMatchArray;
        expect(TANH.calculate(matches2, mockContext)).toBeCloseTo(Math.tanh(1));
      });
    });
  });

  describe('Inverse Hyperbolic Functions', () => {
    describe('ASINH', () => {
      it('should calculate inverse hyperbolic sine', () => {
        const matches = ['ASINH(A1)', 'A1'] as RegExpMatchArray;
        expect(ASINH.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['ASINH(E1)', 'E1'] as RegExpMatchArray;
        expect(ASINH.calculate(matches2, mockContext)).toBeCloseTo(Math.asinh(1));
      });
    });

    describe('ACOSH', () => {
      it('should calculate inverse hyperbolic cosine', () => {
        const matches = ['ACOSH(E1)', 'E1'] as RegExpMatchArray;
        expect(ACOSH.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['ACOSH(H1)', 'H1'] as RegExpMatchArray;
        expect(ACOSH.calculate(matches2, mockContext)).toBeCloseTo(Math.acosh(2));
      });

      it('should return NUM error for values < 1', () => {
        const matches = ['ACOSH(D1)', 'D1'] as RegExpMatchArray;
        expect(ACOSH.calculate(matches, mockContext)).toBe(FormulaError.NUM);
      });
    });

    describe('ATANH', () => {
      it('should calculate inverse hyperbolic tangent', () => {
        const matches = ['ATANH(A1)', 'A1'] as RegExpMatchArray;
        expect(ATANH.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['ATANH(D1)', 'D1'] as RegExpMatchArray;
        expect(ATANH.calculate(matches2, mockContext)).toBeCloseTo(Math.atanh(0.5));
      });

      it('should return NUM error for values outside (-1, 1)', () => {
        const matches = ['ATANH(E1)', 'E1'] as RegExpMatchArray;
        expect(ATANH.calculate(matches, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['ATANH(F1)', 'F1'] as RegExpMatchArray;
        expect(ATANH.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
      });
    });
  });

  describe('Angle Conversion Functions', () => {
    describe('DEGREES', () => {
      it('should convert radians to degrees', () => {
        const matches = ['DEGREES(A1)', 'A1'] as RegExpMatchArray;
        expect(DEGREES.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['DEGREES(B1)', 'B1'] as RegExpMatchArray;
        expect(DEGREES.calculate(matches2, mockContext)).toBeCloseTo(90);
        
        const matches3 = ['DEGREES(C1)', 'C1'] as RegExpMatchArray;
        expect(DEGREES.calculate(matches3, mockContext)).toBeCloseTo(180);
      });

      it('should return VALUE error for non-numeric input', () => {
        const matches = ['DEGREES(G1)', 'G1'] as RegExpMatchArray;
        expect(DEGREES.calculate(matches, mockContext)).toBe(FormulaError.VALUE);
      });
    });

    describe('RADIANS', () => {
      it('should convert degrees to radians', () => {
        const matches = ['RADIANS(A1)', 'A1'] as RegExpMatchArray;
        expect(RADIANS.calculate(matches, mockContext)).toBe(0);
        
        const matches2 = ['RADIANS(L1)', 'L1'] as RegExpMatchArray;
        expect(RADIANS.calculate(matches2, mockContext)).toBeCloseTo(Math.PI / 2);
        
        const matches3 = ['RADIANS(M1)', 'M1'] as RegExpMatchArray;
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
        const matches = ['CSC(B1)', 'B1'] as RegExpMatchArray;
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
        
        const matches2 = ['SEC(C1)', 'C1'] as RegExpMatchArray;
        expect(SEC.calculate(matches2, mockContext)).toBeCloseTo(-1); // sec(π) = -1
      });

      it('should return DIV0 error when cos is zero', () => {
        const matches = ['SEC(B1)', 'B1'] as RegExpMatchArray;
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
        const matches = ['ACOT(E1)', 'E1'] as RegExpMatchArray;
        expect(ACOT.calculate(matches, mockContext)).toBeCloseTo(Math.PI / 4);
        
        const matches2 = ['ACOT(A1)', 'A1'] as RegExpMatchArray;
        expect(ACOT.calculate(matches2, mockContext)).toBeCloseTo(Math.PI / 2);
      });
    });
  });

  describe('Hyperbolic Reciprocal Functions', () => {
    describe('CSCH', () => {
      it('should calculate hyperbolic cosecant', () => {
        const matches = ['CSCH(E1)', 'E1'] as RegExpMatchArray;
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
        
        const matches2 = ['SECH(E1)', 'E1'] as RegExpMatchArray;
        expect(SECH.calculate(matches2, mockContext)).toBeCloseTo(1 / Math.cosh(1));
      });
    });

    describe('COTH', () => {
      it('should calculate hyperbolic cotangent', () => {
        const matches = ['COTH(E1)', 'E1'] as RegExpMatchArray;
        expect(COTH.calculate(matches, mockContext)).toBeCloseTo(1 / Math.tanh(1));
      });

      it('should return DIV0 error for zero', () => {
        const matches = ['COTH(A1)', 'A1'] as RegExpMatchArray;
        expect(COTH.calculate(matches, mockContext)).toBe(FormulaError.DIV0);
      });
    });

    describe('ACOTH', () => {
      it('should calculate inverse hyperbolic cotangent', () => {
        const matches = ['ACOTH(H1)', 'H1'] as RegExpMatchArray;
        const expected = 0.5 * Math.log(3); // 0.5 * ln((2+1)/(2-1))
        expect(ACOTH.calculate(matches, mockContext)).toBeCloseTo(expected);
      });

      it('should return NUM error for values in [-1, 1]', () => {
        const matches = ['ACOTH(D1)', 'D1'] as RegExpMatchArray;
        expect(ACOTH.calculate(matches, mockContext)).toBe(FormulaError.NUM);
        
        const matches2 = ['ACOTH(E1)', 'E1'] as RegExpMatchArray;
        expect(ACOTH.calculate(matches2, mockContext)).toBe(FormulaError.NUM);
      });
    });
  });
});