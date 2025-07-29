/* eslint-disable @typescript-eslint/no-non-null-assertion */
import { describe, it, expect } from 'vitest';
import { 
  MULTIPLY_OPERATOR, DIVIDE_OPERATOR, ADD_OPERATOR, SUBTRACT_OPERATOR,
  GREATER_THAN_OR_EQUAL, LESS_THAN_OR_EQUAL, GREATER_THAN, LESS_THAN,
  EQUAL, NOT_EQUAL
} from '../basicOperatorLogic';
import type { FormulaContext } from '../../shared/types';

// モックコンテキスト作成
const createContext = (data: (string | number | boolean | null)[][]): FormulaContext => ({
  data: data.map(row => row.map(cell => ({ value: cell }))),
  row: 0,
  col: 0
});

// テスト用のマッチ配列作成
const createMatches = (formula: string, ...groups: string[]): RegExpMatchArray => {
  const matches = [formula, ...groups] as RegExpMatchArray;
  matches.index = 0;
  matches.input = formula;
  return matches;
};

describe('基本演算子のテスト', () => {
  describe('算術演算子', () => {
    describe('乗算演算子 (*)', () => {
      it('数値同士の乗算が正しく動作する', () => {
        const context = createContext([[5], [3]]);
        
        // 5 * 3 = 15
        const result = MULTIPLY_OPERATOR.calculate!(createMatches('5*3', '5', '3'), context);
        expect(result).toBe(15);
      });

      it('セル参照を使った乗算が正しく動作する', () => {
        const context = createContext([[10, 4]]);
        
        // A1 * B1 = 40
        const result = MULTIPLY_OPERATOR.calculate!(createMatches('A1*B1', 'A1', 'B1'), context);
        expect(result).toBe(40);
      });

      it('小数点の乗算が正しく動作する', () => {
        const context = createContext([[2.5], [4.2]]);
        
        // 2.5 * 4.2 = 10.5
        const result = MULTIPLY_OPERATOR.calculate!(createMatches('2.5*4.2', '2.5', '4.2'), context);
        expect(result).toBeCloseTo(10.5, 5);
      });

      it('負の数の乗算が正しく動作する', () => {
        const context = createContext([[-5], [3]]);
        
        // -5 * 3 = -15
        const result = MULTIPLY_OPERATOR.calculate!(createMatches('-5*3', '-5', '3'), context);
        expect(result).toBe(-15);
      });

      it('無効な値でエラーを返す', () => {
        const context = createContext([['text', 3]]);
        
        // text * 3 = #VALUE!
        const result = MULTIPLY_OPERATOR.calculate!(createMatches('A1*B1', 'A1', 'B1'), context);
        expect(result).toBe('#VALUE!');
      });
    });

    describe('除算演算子 (/)', () => {
      it('数値同士の除算が正しく動作する', () => {
        const context = createContext([[15], [3]]);
        
        // 15 / 3 = 5
        const result = DIVIDE_OPERATOR.calculate!(createMatches('15/3', '15', '3'), context);
        expect(result).toBe(5);
      });

      it('セル参照を使った除算が正しく動作する', () => {
        const context = createContext([[20, 4]]);
        
        // A1 / B1 = 5
        const result = DIVIDE_OPERATOR.calculate!(createMatches('A1/B1', 'A1', 'B1'), context);
        expect(result).toBe(5);
      });

      it('小数点の除算が正しく動作する', () => {
        const context = createContext([[10.5], [2.5]]);
        
        // 10.5 / 2.5 = 4.2
        const result = DIVIDE_OPERATOR.calculate!(createMatches('10.5/2.5', '10.5', '2.5'), context);
        expect(result).toBeCloseTo(4.2, 5);
      });

      it('ゼロ除算でエラーを返す', () => {
        const context = createContext([[10, 0]]);
        
        // 10 / 0 = #DIV/0!
        const result = DIVIDE_OPERATOR.calculate!(createMatches('10/0', '10', '0'), context);
        expect(result).toBe('#DIV/0!');
      });

      it('無効な値でエラーを返す', () => {
        const context = createContext([['text', 3]]);
        
        // text / 3 = #VALUE!
        const result = DIVIDE_OPERATOR.calculate!(createMatches('A1/B1', 'A1', 'B1'), context);
        expect(result).toBe('#VALUE!');
      });
    });

    describe('加算演算子 (+)', () => {
      it('数値同士の加算が正しく動作する', () => {
        const context = createContext([[5], [3]]);
        
        // 5 + 3 = 8
        const result = ADD_OPERATOR.calculate!(createMatches('5+3', '5', '3'), context);
        expect(result).toBe(8);
      });

      it('セル参照を使った加算が正しく動作する', () => {
        const context = createContext([[10, 15]]);
        
        // A1 + B1 = 25
        const result = ADD_OPERATOR.calculate!(createMatches('A1+B1', 'A1', 'B1'), context);
        expect(result).toBe(25);
      });

      it('小数点の加算が正しく動作する', () => {
        const context = createContext([[2.5], [3.7]]);
        
        // 2.5 + 3.7 = 6.2
        const result = ADD_OPERATOR.calculate!(createMatches('2.5+3.7', '2.5', '3.7'), context);
        expect(result).toBeCloseTo(6.2, 5);
      });

      it('負の数の加算が正しく動作する', () => {
        const context = createContext([[-5], [3]]);
        
        // -5 + 3 = -2
        const result = ADD_OPERATOR.calculate!(createMatches('-5+3', '-5', '3'), context);
        expect(result).toBe(-2);
      });

      it('無効な値でエラーを返す', () => {
        const context = createContext([['text', 3]]);
        
        // text + 3 = #VALUE!
        const result = ADD_OPERATOR.calculate!(createMatches('A1+B1', 'A1', 'B1'), context);
        expect(result).toBe('#VALUE!');
      });
    });

    describe('減算演算子 (-)', () => {
      it('数値同士の減算が正しく動作する', () => {
        const context = createContext([[10], [3]]);
        
        // 10 - 3 = 7
        const result = SUBTRACT_OPERATOR.calculate!(createMatches('10-3', '10', '3'), context);
        expect(result).toBe(7);
      });

      it('セル参照を使った減算が正しく動作する', () => {
        const context = createContext([[20, 5]]);
        
        // A1 - B1 = 15
        const result = SUBTRACT_OPERATOR.calculate!(createMatches('A1-B1', 'A1', 'B1'), context);
        expect(result).toBe(15);
      });

      it('小数点の減算が正しく動作する', () => {
        const context = createContext([[7.5], [2.3]]);
        
        // 7.5 - 2.3 = 5.2
        const result = SUBTRACT_OPERATOR.calculate!(createMatches('7.5-2.3', '7.5', '2.3'), context);
        expect(result).toBeCloseTo(5.2, 5);
      });

      it('負の数の減算が正しく動作する', () => {
        const context = createContext([[5], [-3]]);
        
        // 5 - (-3) = 8
        const result = SUBTRACT_OPERATOR.calculate!(createMatches('5--3', '5', '-3'), context);
        expect(result).toBe(8);
      });

      it('無効な値でエラーを返す', () => {
        const context = createContext([['text', 3]]);
        
        // text - 3 = #VALUE!
        const result = SUBTRACT_OPERATOR.calculate!(createMatches('A1-B1', 'A1', 'B1'), context);
        expect(result).toBe('#VALUE!');
      });
    });
  });

  describe('比較演算子', () => {
    describe('以上演算子 (>=)', () => {
      it('数値の比較が正しく動作する', () => {
        const context = createContext([[10], [5]]);
        
        // 10 >= 5 = true
        expect(GREATER_THAN_OR_EQUAL.calculate!(createMatches('10>=5', '10', '5'), context)).toBe(true);
        
        // 5 >= 10 = false
        expect(GREATER_THAN_OR_EQUAL.calculate!(createMatches('5>=10', '5', '10'), context)).toBe(false);
        
        // 5 >= 5 = true
        expect(GREATER_THAN_OR_EQUAL.calculate!(createMatches('5>=5', '5', '5'), context)).toBe(true);
      });

      it('セル参照を使った比較が正しく動作する', () => {
        const context = createContext([[10, 5]]);
        
        // A1 >= B1 = true
        const result = GREATER_THAN_OR_EQUAL.calculate!(createMatches('A1>=B1', 'A1', 'B1'), context);
        expect(result).toBe(true);
      });

      it('無効な値でfalseを返す', () => {
        const context = createContext([['text', 5]]);
        
        // text >= 5 = false
        const result = GREATER_THAN_OR_EQUAL.calculate!(createMatches('A1>=B1', 'A1', 'B1'), context);
        expect(result).toBe(false);
      });
    });

    describe('以下演算子 (<=)', () => {
      it('数値の比較が正しく動作する', () => {
        const context = createContext([[5], [10]]);
        
        // 5 <= 10 = true
        expect(LESS_THAN_OR_EQUAL.calculate!(createMatches('5<=10', '5', '10'), context)).toBe(true);
        
        // 10 <= 5 = false
        expect(LESS_THAN_OR_EQUAL.calculate!(createMatches('10<=5', '10', '5'), context)).toBe(false);
        
        // 5 <= 5 = true
        expect(LESS_THAN_OR_EQUAL.calculate!(createMatches('5<=5', '5', '5'), context)).toBe(true);
      });

      it('セル参照を使った比較が正しく動作する', () => {
        const context = createContext([[5, 10]]);
        
        // A1 <= B1 = true
        const result = LESS_THAN_OR_EQUAL.calculate!(createMatches('A1<=B1', 'A1', 'B1'), context);
        expect(result).toBe(true);
      });

      it('無効な値でfalseを返す', () => {
        const context = createContext([['text', 5]]);
        
        // text <= 5 = false
        const result = LESS_THAN_OR_EQUAL.calculate!(createMatches('A1<=B1', 'A1', 'B1'), context);
        expect(result).toBe(false);
      });
    });

    describe('より大きい演算子 (>)', () => {
      it('数値の比較が正しく動作する', () => {
        const context = createContext([[10], [5]]);
        
        // 10 > 5 = true
        expect(GREATER_THAN.calculate!(createMatches('10>5', '10', '5'), context)).toBe(true);
        
        // 5 > 10 = false
        expect(GREATER_THAN.calculate!(createMatches('5>10', '5', '10'), context)).toBe(false);
        
        // 5 > 5 = false
        expect(GREATER_THAN.calculate!(createMatches('5>5', '5', '5'), context)).toBe(false);
      });

      it('セル参照を使った比較が正しく動作する', () => {
        const context = createContext([[10, 5]]);
        
        // A1 > B1 = true
        const result = GREATER_THAN.calculate!(createMatches('A1>B1', 'A1', 'B1'), context);
        expect(result).toBe(true);
      });

      it('無効な値でfalseを返す', () => {
        const context = createContext([['text', 5]]);
        
        // text > 5 = false
        const result = GREATER_THAN.calculate!(createMatches('A1>B1', 'A1', 'B1'), context);
        expect(result).toBe(false);
      });
    });

    describe('より小さい演算子 (<)', () => {
      it('数値の比較が正しく動作する', () => {
        const context = createContext([[5], [10]]);
        
        // 5 < 10 = true
        expect(LESS_THAN.calculate!(createMatches('5<10', '5', '10'), context)).toBe(true);
        
        // 10 < 5 = false
        expect(LESS_THAN.calculate!(createMatches('10<5', '10', '5'), context)).toBe(false);
        
        // 5 < 5 = false
        expect(LESS_THAN.calculate!(createMatches('5<5', '5', '5'), context)).toBe(false);
      });

      it('セル参照を使った比較が正しく動作する', () => {
        const context = createContext([[5, 10]]);
        
        // A1 < B1 = true
        const result = LESS_THAN.calculate!(createMatches('A1<B1', 'A1', 'B1'), context);
        expect(result).toBe(true);
      });

      it('無効な値でfalseを返す', () => {
        const context = createContext([['text', 5]]);
        
        // text < 5 = false
        const result = LESS_THAN.calculate!(createMatches('A1<B1', 'A1', 'B1'), context);
        expect(result).toBe(false);
      });
    });

    describe('等しい演算子 (=)', () => {
      it('数値の比較が正しく動作する', () => {
        const context = createContext([[5], [5]]);
        
        // 5 = 5 = true
        expect(EQUAL.calculate!(createMatches('5=5', '5', '5'), context)).toBe(true);
        
        // 5 = 10 = false
        expect(EQUAL.calculate!(createMatches('5=10', '5', '10'), context)).toBe(false);
      });

      it('文字列の比較が正しく動作する', () => {
        const context = createContext([['hello', 'world']]);
        
        // "hello" = "hello" = true
        expect(EQUAL.calculate!(createMatches('hello=hello', 'hello', 'hello'), context)).toBe(true);
        
        // "hello" = "world" = false
        expect(EQUAL.calculate!(createMatches('hello=world', 'hello', 'world'), context)).toBe(false);
      });

      it('セル参照を使った比較が正しく動作する', () => {
        const context = createContext([[5, 5]]);
        
        // A1 = B1 = true
        const result = EQUAL.calculate!(createMatches('A1=B1', 'A1', 'B1'), context);
        expect(result).toBe(true);
      });

      it('異なる型でも正しく比較する', () => {
        const context = createContext([[5, '5']]);
        
        // 5 = "5" = false (型が異なる)
        const result = EQUAL.calculate!(createMatches('A1=B1', 'A1', 'B1'), context);
        expect(result).toBe(false);
      });
    });

    describe('等しくない演算子 (<>)', () => {
      it('数値の比較が正しく動作する', () => {
        const context = createContext([[5], [10]]);
        
        // 5 <> 10 = true
        expect(NOT_EQUAL.calculate!(createMatches('5<>10', '5', '10'), context)).toBe(true);
        
        // 5 <> 5 = false
        expect(NOT_EQUAL.calculate!(createMatches('5<>5', '5', '5'), context)).toBe(false);
      });

      it('文字列の比較が正しく動作する', () => {
        const context = createContext([['hello', 'world']]);
        
        // "hello" <> "world" = true
        expect(NOT_EQUAL.calculate!(createMatches('hello<>world', 'hello', 'world'), context)).toBe(true);
        
        // "hello" <> "hello" = false
        expect(NOT_EQUAL.calculate!(createMatches('hello<>hello', 'hello', 'hello'), context)).toBe(false);
      });

      it('セル参照を使った比較が正しく動作する', () => {
        const context = createContext([[5, 10]]);
        
        // A1 <> B1 = true
        const result = NOT_EQUAL.calculate!(createMatches('A1<>B1', 'A1', 'B1'), context);
        expect(result).toBe(true);
      });

      it('異なる型で正しく比較する', () => {
        const context = createContext([[5, '5']]);
        
        // 5 <> "5" = true (型が異なる)
        const result = NOT_EQUAL.calculate!(createMatches('A1<>B1', 'A1', 'B1'), context);
        expect(result).toBe(true);
      });
    });
  });

  describe('複雑なケースのテスト', () => {
    it('空白セルとの演算が正しく動作する', () => {
      const context = createContext([[5, '']]);
      
      // 5 + "" = 5 (空文字列は0として扱われる)
      expect(ADD_OPERATOR.calculate!(createMatches('A1+B1', 'A1', 'B1'), context)).toBe(5);
      
      // 5 = "" = false
      expect(EQUAL.calculate!(createMatches('A1=B1', 'A1', 'B1'), context)).toBe(false);
    });

    it('nullとの演算が正しく動作する', () => {
      const context = createContext([[5, null]]);
      
      // 5 + null = #VALUE!
      expect(ADD_OPERATOR.calculate!(createMatches('A1+B1', 'A1', 'B1'), context)).toBe('#VALUE!');
      
      // 5 = null = false
      expect(EQUAL.calculate!(createMatches('A1=B1', 'A1', 'B1'), context)).toBe(false);
    });

    it('スペースを含む式が正しく処理される', () => {
      const context = createContext([[5], [3]]);
      
      // " 5 " + " 3 " = 8 (スペースはトリムされる)
      const result = ADD_OPERATOR.calculate!(createMatches(' 5 + 3 ', ' 5 ', ' 3 '), context);
      expect(result).toBe(8);
    });

    it('複数桁の数値が正しく処理される', () => {
      const context = createContext([[123, 456]]);
      
      // 123 + 456 = 579
      const result = ADD_OPERATOR.calculate!(createMatches('123+456', '123', '456'), context);
      expect(result).toBe(579);
    });
  });
});