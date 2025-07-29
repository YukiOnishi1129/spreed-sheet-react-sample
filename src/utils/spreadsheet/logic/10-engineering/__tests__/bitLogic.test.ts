import { describe, it, expect } from 'vitest';
import {
  BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT
} from '../bitLogic';
import { FormulaError } from '../../shared/types';
import type { FormulaContext } from '../../shared/types';

// Helper function to create FormulaContext
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

describe('Bit Operation Functions', () => {
  const mockContext = createContext([
    [0, 1, 2, 3, 4, 5, 6, 7, 8], // small integers
    [10, 15, 20, 25, 30, 35, 40, 45, 50], // medium integers
    [100, 255, 1024, 2048, 4096], // larger integers
    [-1, -5, -10, -15], // negative integers
    [1.5, 2.7, 3.9], // decimal numbers
    [281474976710655], // maximum allowed value (2^48 - 1)
  ]);

  describe('BITAND Function (Bitwise AND)', () => {
    it('should calculate bitwise AND of two numbers', () => {
      const matches = ['BITAND(5, 3)', '5', '3'] as RegExpMatchArray;
      const result = BITAND.calculate(matches, mockContext);
      expect(result).toBe(1); // 5 (101) & 3 (011) = 1 (001)
    });

    it('should handle AND with zero', () => {
      const matches = ['BITAND(15, 0)', '15', '0'] as RegExpMatchArray;
      const result = BITAND.calculate(matches, mockContext);
      expect(result).toBe(0); // Any number AND 0 = 0
    });

    it('should handle AND with same number', () => {
      const matches = ['BITAND(7, 7)', '7', '7'] as RegExpMatchArray;
      const result = BITAND.calculate(matches, mockContext);
      expect(result).toBe(7); // Number AND itself = number
    });

    it('should handle power of 2 masking', () => {
      const matches = ['BITAND(255, 15)', '255', '15'] as RegExpMatchArray;
      const result = BITAND.calculate(matches, mockContext);
      expect(result).toBe(15); // 11111111 & 00001111 = 00001111
    });

    it('should truncate decimal numbers', () => {
      const matches = ['BITAND(5.7, 3.2)', '5.7', '3.2'] as RegExpMatchArray;
      const result = BITAND.calculate(matches, mockContext);
      expect(result).toBe(1); // Truncates to BITAND(5, 3)
    });

    it('should return NUM error for negative numbers', () => {
      const matches = ['BITAND(-1, 5)', '-1', '5'] as RegExpMatchArray;
      const result = BITAND.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for numbers > 2^48-1', () => {
      const matches = ['BITAND(281474976710656, 1)', '281474976710656', '1'] as RegExpMatchArray;
      const result = BITAND.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should handle cell references', () => {
      const matches = ['BITAND(B1, C1)', 'B1', 'C1'] as RegExpMatchArray;
      const result = BITAND.calculate(matches, mockContext);
      expect(result).toBe(10); // 10 & 15 = 10
    });
  });

  describe('BITOR Function (Bitwise OR)', () => {
    it('should calculate bitwise OR of two numbers', () => {
      const matches = ['BITOR(5, 3)', '5', '3'] as RegExpMatchArray;
      const result = BITOR.calculate(matches, mockContext);
      expect(result).toBe(7); // 5 (101) | 3 (011) = 7 (111)
    });

    it('should handle OR with zero', () => {
      const matches = ['BITOR(15, 0)', '15', '0'] as RegExpMatchArray;
      const result = BITOR.calculate(matches, mockContext);
      expect(result).toBe(15); // Any number OR 0 = number
    });

    it('should handle OR with same number', () => {
      const matches = ['BITOR(7, 7)', '7', '7'] as RegExpMatchArray;
      const result = BITOR.calculate(matches, mockContext);
      expect(result).toBe(7); // Number OR itself = number
    });

    it('should combine bit patterns', () => {
      const matches = ['BITOR(240, 15)', '240', '15'] as RegExpMatchArray;
      const result = BITOR.calculate(matches, mockContext);
      expect(result).toBe(255); // 11110000 | 00001111 = 11111111
    });

    it('should handle large numbers', () => {
      const matches = ['BITOR(1024, 2048)', '1024', '2048'] as RegExpMatchArray;
      const result = BITOR.calculate(matches, mockContext);
      expect(result).toBe(3072); // 10000000000 | 100000000000 = 110000000000
    });

    it('should return NUM error for negative numbers', () => {
      const matches = ['BITOR(5, -3)', '5', '-3'] as RegExpMatchArray;
      const result = BITOR.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return VALUE error for invalid input', () => {
      const matches = ['BITOR("text", 5)', '"text"', '5'] as RegExpMatchArray;
      const result = BITOR.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe('BITXOR Function (Bitwise XOR)', () => {
    it('should calculate bitwise XOR of two numbers', () => {
      const matches = ['BITXOR(5, 3)', '5', '3'] as RegExpMatchArray;
      const result = BITXOR.calculate(matches, mockContext);
      expect(result).toBe(6); // 5 (101) ^ 3 (011) = 6 (110)
    });

    it('should handle XOR with zero', () => {
      const matches = ['BITXOR(15, 0)', '15', '0'] as RegExpMatchArray;
      const result = BITXOR.calculate(matches, mockContext);
      expect(result).toBe(15); // Any number XOR 0 = number
    });

    it('should handle XOR with same number', () => {
      const matches = ['BITXOR(7, 7)', '7', '7'] as RegExpMatchArray;
      const result = BITXOR.calculate(matches, mockContext);
      expect(result).toBe(0); // Number XOR itself = 0
    });

    it('should toggle bits', () => {
      const matches = ['BITXOR(255, 170)', '255', '170'] as RegExpMatchArray;
      const result = BITXOR.calculate(matches, mockContext);
      expect(result).toBe(85); // 11111111 ^ 10101010 = 01010101
    });

    it('should be symmetric', () => {
      const matches1 = ['BITXOR(12, 5)', '12', '5'] as RegExpMatchArray;
      const result1 = BITXOR.calculate(matches1, mockContext);
      
      const matches2 = ['BITXOR(5, 12)', '5', '12'] as RegExpMatchArray;
      const result2 = BITXOR.calculate(matches2, mockContext);
      
      expect(result1).toBe(result2); // XOR is commutative
    });

    it('should return NUM error for numbers > 2^48-1', () => {
      const matches = ['BITXOR(281474976710656, 1)', '281474976710656', '1'] as RegExpMatchArray;
      const result = BITXOR.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });
  });

  describe('BITLSHIFT Function (Bitwise Left Shift)', () => {
    it('should shift bits left', () => {
      const matches = ['BITLSHIFT(4, 2)', '4', '2'] as RegExpMatchArray;
      const result = BITLSHIFT.calculate(matches, mockContext);
      expect(result).toBe(16); // 4 << 2 = 16 (100 -> 10000)
    });

    it('should handle zero shift', () => {
      const matches = ['BITLSHIFT(15, 0)', '15', '0'] as RegExpMatchArray;
      const result = BITLSHIFT.calculate(matches, mockContext);
      expect(result).toBe(15); // No shift
    });

    it('should multiply by powers of 2', () => {
      const matches = ['BITLSHIFT(1, 10)', '1', '10'] as RegExpMatchArray;
      const result = BITLSHIFT.calculate(matches, mockContext);
      expect(result).toBe(1024); // 1 << 10 = 2^10 = 1024
    });

    it('should handle large shifts within bounds', () => {
      const matches = ['BITLSHIFT(1, 47)', '1', '47'] as RegExpMatchArray;
      const result = BITLSHIFT.calculate(matches, mockContext);
      expect(result).toBe(140737488355328); // 2^47
    });

    it('should return NUM error for negative shift', () => {
      const matches = ['BITLSHIFT(4, -1)', '4', '-1'] as RegExpMatchArray;
      const result = BITLSHIFT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for shift > 53', () => {
      const matches = ['BITLSHIFT(1, 54)', '1', '54'] as RegExpMatchArray;
      const result = BITLSHIFT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error if result > 2^48-1', () => {
      const matches = ['BITLSHIFT(1, 48)', '1', '48'] as RegExpMatchArray;
      const result = BITLSHIFT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should truncate decimal shifts', () => {
      const matches = ['BITLSHIFT(4, 2.9)', '4', '2.9'] as RegExpMatchArray;
      const result = BITLSHIFT.calculate(matches, mockContext);
      expect(result).toBe(16); // Truncates to shift of 2
    });
  });

  describe('BITRSHIFT Function (Bitwise Right Shift)', () => {
    it('should shift bits right', () => {
      const matches = ['BITRSHIFT(16, 2)', '16', '2'] as RegExpMatchArray;
      const result = BITRSHIFT.calculate(matches, mockContext);
      expect(result).toBe(4); // 16 >> 2 = 4 (10000 -> 100)
    });

    it('should handle zero shift', () => {
      const matches = ['BITRSHIFT(15, 0)', '15', '0'] as RegExpMatchArray;
      const result = BITRSHIFT.calculate(matches, mockContext);
      expect(result).toBe(15); // No shift
    });

    it('should divide by powers of 2', () => {
      const matches = ['BITRSHIFT(1024, 10)', '1024', '10'] as RegExpMatchArray;
      const result = BITRSHIFT.calculate(matches, mockContext);
      expect(result).toBe(1); // 1024 >> 10 = 1024 / 2^10 = 1
    });

    it('should truncate fractional results', () => {
      const matches = ['BITRSHIFT(15, 2)', '15', '2'] as RegExpMatchArray;
      const result = BITRSHIFT.calculate(matches, mockContext);
      expect(result).toBe(3); // 15 >> 2 = 3 (1111 -> 11)
    });

    it('should handle shifts equal to bit length', () => {
      const matches = ['BITRSHIFT(255, 8)', '255', '8'] as RegExpMatchArray;
      const result = BITRSHIFT.calculate(matches, mockContext);
      expect(result).toBe(0); // All bits shifted out
    });

    it('should return NUM error for negative shift', () => {
      const matches = ['BITRSHIFT(16, -2)', '16', '-2'] as RegExpMatchArray;
      const result = BITRSHIFT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should return NUM error for shift > 53', () => {
      const matches = ['BITRSHIFT(1024, 54)', '1024', '54'] as RegExpMatchArray;
      const result = BITRSHIFT.calculate(matches, mockContext);
      expect(result).toBe(FormulaError.NUM);
    });

    it('should handle cell references', () => {
      const matches = ['BITRSHIFT(C3, A4)', 'C3', 'A4'] as RegExpMatchArray;
      const result = BITRSHIFT.calculate(matches, mockContext);
      expect(result).toBe(64); // 1024 >> 4 = 64
    });
  });

  describe('Edge Cases and Integration', () => {
    it('should verify bit operation identities', () => {
      // De Morgan's law: ~(A & B) = ~A | ~B
      // In limited bit context: (A XOR all_ones) = bit complement
      const a = 170; // 10101010
      const b = 85;  // 01010101
      
      const andMatches = ['BITAND(170, 85)', '170', '85'] as RegExpMatchArray;
      const andResult = BITAND.calculate(andMatches, mockContext);
      expect(andResult).toBe(0); // No common bits
      
      const orMatches = ['BITOR(170, 85)', '170', '85'] as RegExpMatchArray;
      const orResult = BITOR.calculate(orMatches, mockContext);
      expect(orResult).toBe(255); // All bits set in byte
    });

    it('should verify shift operations are inverse', () => {
      const value = 100;
      const shift = 5;
      
      const leftMatches = ['BITLSHIFT(100, 5)', '100', '5'] as RegExpMatchArray;
      const leftResult = BITLSHIFT.calculate(leftMatches, mockContext) as number;
      
      const rightMatches = ['BITRSHIFT(' + leftResult + ', 5)', leftResult.toString(), '5'] as RegExpMatchArray;
      const rightResult = BITRSHIFT.calculate(rightMatches, mockContext);
      
      expect(rightResult).toBe(value);
    });

    it('should handle maximum safe values', () => {
      const maxSafe = 281474976710655; // 2^48 - 1
      const matches = ['BITAND(281474976710655, 281474976710655)', 
        '281474976710655', '281474976710655'] as RegExpMatchArray;
      const result = BITAND.calculate(matches, mockContext);
      expect(result).toBe(maxSafe);
    });

    it('should verify XOR encryption property', () => {
      const data = 42;
      const key = 123;
      
      // Encrypt
      const encryptMatches = ['BITXOR(42, 123)', '42', '123'] as RegExpMatchArray;
      const encrypted = BITXOR.calculate(encryptMatches, mockContext) as number;
      
      // Decrypt (XOR with same key)
      const decryptMatches = ['BITXOR(' + encrypted + ', 123)', encrypted.toString(), '123'] as RegExpMatchArray;
      const decrypted = BITXOR.calculate(decryptMatches, mockContext);
      
      expect(decrypted).toBe(data);
    });

    it('should handle cell references in all functions', () => {
      // Test with actual cell values
      const matches = ['BITOR(A2, B2)', 'A2', 'B2'] as RegExpMatchArray;
      const result = BITOR.calculate(matches, mockContext);
      expect(result).toBe(15); // 1 | 10 = 11
    });
  });
});