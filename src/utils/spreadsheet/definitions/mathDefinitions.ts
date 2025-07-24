// 数学・三角関数の定義
import { type FunctionDefinition } from '../types/types';
import { COLOR_SCHEMES } from '../types/colorSchemes';

export const MATH_FUNCTIONS: Record<string, FunctionDefinition> = {
  SUM: {
    name: 'SUM',
    syntax: 'SUM(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '合計する最初の数値または範囲' },
      { name: 'number2', desc: '合計する追加の数値または範囲', optional: true }
    ],
    description: '数値の合計を計算します',
    category: 'math',
    examples: ['=SUM(A1:A10)', '=SUM(1,2,3,4,5)', '=SUM(A1:A5,C1:C5)'],
    colorScheme: COLOR_SCHEMES.math
  },
  
  SUMIF: {
    name: 'SUMIF',
    syntax: 'SUMIF(range, criteria, [sum_range])',
    params: [
      { name: 'range', desc: '条件を評価する範囲' },
      { name: 'criteria', desc: '条件（">10"、"=商品A"など）' },
      { name: 'sum_range', desc: '合計する範囲（省略時はrangeを使用）', optional: true }
    ],
    description: '条件に一致するセルの合計を計算します',
    category: 'math',
    examples: ['=SUMIF(A1:A10,">100")', '=SUMIF(B1:B10,"商品A",C1:C10)'],
    colorScheme: COLOR_SCHEMES.math
  },

  SUMIFS: {
    name: 'SUMIFS',
    syntax: 'SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    params: [
      { name: 'sum_range', desc: '合計する範囲' },
      { name: 'criteria_range1', desc: '1つ目の条件範囲' },
      { name: 'criteria1', desc: '1つ目の条件' },
      { name: 'criteria_range2', desc: '2つ目の条件範囲', optional: true },
      { name: 'criteria2', desc: '2つ目の条件', optional: true }
    ],
    description: '複数条件に一致するセルの合計を計算します',
    category: 'math',
    examples: ['=SUMIFS(C1:C10,A1:A10,">100",B1:B10,"商品A")'],
    colorScheme: COLOR_SCHEMES.math
  },

  PRODUCT: {
    name: 'PRODUCT',
    syntax: 'PRODUCT(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '積を求める最初の数値' },
      { name: 'number2', desc: '積を求める追加の数値', optional: true }
    ],
    description: '数値の積を計算します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  SQRT: {
    name: 'SQRT',
    syntax: 'SQRT(number)',
    params: [
      { name: 'number', desc: '平方根を求める数値' }
    ],
    description: '平方根を計算します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  POWER: {
    name: 'POWER',
    syntax: 'POWER(number, power)',
    params: [
      { name: 'number', desc: '底となる数値' },
      { name: 'power', desc: '指数' }
    ],
    description: 'べき乗を計算します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  ABS: {
    name: 'ABS',
    syntax: 'ABS(number)',
    params: [
      { name: 'number', desc: '絶対値を求める数値' }
    ],
    description: '絶対値を返します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  ROUND: {
    name: 'ROUND',
    syntax: 'ROUND(number, num_digits)',
    params: [
      { name: 'number', desc: '四捨五入する数値' },
      { name: 'num_digits', desc: '桁数（正:小数点以下、負:整数部分）' }
    ],
    description: '指定桁数で四捨五入します',
    category: 'math',
    examples: ['=ROUND(3.14159,2)', '=ROUND(123.456,-1)'],
    colorScheme: COLOR_SCHEMES.math
  },

  ROUNDUP: {
    name: 'ROUNDUP',
    syntax: 'ROUNDUP(number, num_digits)',
    params: [
      { name: 'number', desc: '切り上げる数値' },
      { name: 'num_digits', desc: '桁数' }
    ],
    description: '指定桁数で切り上げます',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  ROUNDDOWN: {
    name: 'ROUNDDOWN',
    syntax: 'ROUNDDOWN(number, num_digits)',
    params: [
      { name: 'number', desc: '切り下げる数値' },
      { name: 'num_digits', desc: '桁数' }
    ],
    description: '指定桁数で切り下げます',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  SEQUENCE: {
    name: 'SEQUENCE',
    syntax: 'SEQUENCE(rows, [columns], [start], [step])',
    params: [
      { name: 'rows', desc: '行数' },
      { name: 'columns', desc: '列数', optional: true },
      { name: 'start', desc: '開始値', optional: true },
      { name: 'step', desc: 'ステップ', optional: true }
    ],
    description: '連続した数値の配列を生成します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },
  
  RANDARRAY: {
    name: 'RANDARRAY',
    syntax: 'RANDARRAY([rows], [columns], [min], [max], [whole_number])',
    params: [
      { name: 'rows', desc: '行数', optional: true },
      { name: 'columns', desc: '列数', optional: true },
      { name: 'min', desc: '最小値', optional: true },
      { name: 'max', desc: '最大値', optional: true },
      { name: 'whole_number', desc: '整数フラグ', optional: true }
    ],
    description: 'ランダムな数値の配列を生成します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  }
};