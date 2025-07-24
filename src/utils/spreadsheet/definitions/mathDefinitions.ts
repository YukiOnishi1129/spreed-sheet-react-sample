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

  COUNTIF: {
    name: 'COUNTIF',
    syntax: 'COUNTIF(range, criteria)',
    params: [
      { name: 'range', desc: '条件を評価する範囲' },
      { name: 'criteria', desc: '条件（">10"、"=商品A"など）' }
    ],
    description: '条件に一致するセルの個数を数えます',
    category: 'math',
    examples: ['=COUNTIF(A1:A10,">100")', '=COUNTIF(B1:B10,"商品A")'],
    colorScheme: COLOR_SCHEMES.math
  },

  COUNTIFS: {
    name: 'COUNTIFS',
    syntax: 'COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    params: [
      { name: 'criteria_range1', desc: '1つ目の条件範囲' },
      { name: 'criteria1', desc: '1つ目の条件' },
      { name: 'criteria_range2', desc: '2つ目の条件範囲', optional: true },
      { name: 'criteria2', desc: '2つ目の条件', optional: true }
    ],
    description: '複数条件に一致するセルの個数を数えます',
    category: 'math',
    examples: ['=COUNTIFS(A1:A10,">100",B1:B10,"商品A")'],
    colorScheme: COLOR_SCHEMES.math
  },

  AVERAGEIF: {
    name: 'AVERAGEIF',
    syntax: 'AVERAGEIF(range, criteria, [average_range])',
    params: [
      { name: 'range', desc: '条件を評価する範囲' },
      { name: 'criteria', desc: '条件（">10"、"=商品A"など）' },
      { name: 'average_range', desc: '平均する範囲（省略時はrangeを使用）', optional: true }
    ],
    description: '条件に一致するセルの平均を計算します',
    category: 'math',
    examples: ['=AVERAGEIF(A1:A10,">100")', '=AVERAGEIF(B1:B10,"商品A",C1:C10)'],
    colorScheme: COLOR_SCHEMES.math
  },

  AVERAGEIFS: {
    name: 'AVERAGEIFS',
    syntax: 'AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    params: [
      { name: 'average_range', desc: '平均する範囲' },
      { name: 'criteria_range1', desc: '1つ目の条件範囲' },
      { name: 'criteria1', desc: '1つ目の条件' },
      { name: 'criteria_range2', desc: '2つ目の条件範囲', optional: true },
      { name: 'criteria2', desc: '2つ目の条件', optional: true }
    ],
    description: '複数条件に一致するセルの平均を計算します',
    category: 'math',
    examples: ['=AVERAGEIFS(C1:C10,A1:A10,">100",B1:B10,"商品A")'],
    colorScheme: COLOR_SCHEMES.math
  },

  RAND: {
    name: 'RAND',
    syntax: 'RAND()',
    params: [],
    description: '0以上1未満の乱数を返します',
    category: 'math',
    examples: ['=RAND()'],
    colorScheme: COLOR_SCHEMES.math
  },

  RANDBETWEEN: {
    name: 'RANDBETWEEN',
    syntax: 'RANDBETWEEN(bottom, top)',
    params: [
      { name: 'bottom', desc: '最小値' },
      { name: 'top', desc: '最大値' }
    ],
    description: '指定された範囲内の整数の乱数を返します',
    category: 'math',
    examples: ['=RANDBETWEEN(1,100)', '=RANDBETWEEN(-10,10)'],
    colorScheme: COLOR_SCHEMES.math
  },

  MOD: {
    name: 'MOD',
    syntax: 'MOD(number, divisor)',
    params: [
      { name: 'number', desc: '被除数' },
      { name: 'divisor', desc: '除数' }
    ],
    description: '除算の余りを返します',
    category: 'math',
    examples: ['=MOD(10,3)', '=MOD(A1,B1)'],
    colorScheme: COLOR_SCHEMES.math
  },

  CEILING: {
    name: 'CEILING',
    syntax: 'CEILING(number, significance)',
    params: [
      { name: 'number', desc: '切り上げる数値' },
      { name: 'significance', desc: '基準値' }
    ],
    description: '基準値の倍数に切り上げます',
    category: 'math',
    examples: ['=CEILING(4.3,1)', '=CEILING(4.3,0.5)'],
    colorScheme: COLOR_SCHEMES.math
  },

  CEILING_MATH: {
    name: 'CEILING.MATH',
    syntax: 'CEILING.MATH(number, [significance], [mode])',
    params: [
      { name: 'number', desc: '切り上げる数値' },
      { name: 'significance', desc: '基準値', optional: true },
      { name: 'mode', desc: '負の数の丸めモード', optional: true }
    ],
    description: '基準値の倍数に切り上げます（数学的な切り上げ）',
    category: 'math',
    examples: ['=CEILING.MATH(4.3)', '=CEILING.MATH(4.3,0.5)'],
    colorScheme: COLOR_SCHEMES.math
  },

  FLOOR: {
    name: 'FLOOR',
    syntax: 'FLOOR(number, significance)',
    params: [
      { name: 'number', desc: '切り下げる数値' },
      { name: 'significance', desc: '基準値' }
    ],
    description: '基準値の倍数に切り下げます',
    category: 'math',
    examples: ['=FLOOR(4.7,1)', '=FLOOR(4.7,0.5)'],
    colorScheme: COLOR_SCHEMES.math
  },

  FLOOR_MATH: {
    name: 'FLOOR.MATH',
    syntax: 'FLOOR.MATH(number, [significance], [mode])',
    params: [
      { name: 'number', desc: '切り下げる数値' },
      { name: 'significance', desc: '基準値', optional: true },
      { name: 'mode', desc: '負の数の丸めモード', optional: true }
    ],
    description: '基準値の倍数に切り下げます（数学的な切り下げ）',
    category: 'math',
    examples: ['=FLOOR.MATH(4.7)', '=FLOOR.MATH(4.7,0.5)'],
    colorScheme: COLOR_SCHEMES.math
  },

  TRUNC: {
    name: 'TRUNC',
    syntax: 'TRUNC(number, [num_digits])',
    params: [
      { name: 'number', desc: '切り捨てる数値' },
      { name: 'num_digits', desc: '残す桁数', optional: true }
    ],
    description: '数値の小数部を切り捨てて整数または指定桁数にします',
    category: 'math',
    examples: ['=TRUNC(4.7)', '=TRUNC(4.789,2)'],
    colorScheme: COLOR_SCHEMES.math
  },

  INT: {
    name: 'INT',
    syntax: 'INT(number)',
    params: [
      { name: 'number', desc: '切り捨てる数値' }
    ],
    description: '数値を切り捨てて整数にします',
    category: 'math',
    examples: ['=INT(4.7)', '=INT(-4.7)'],
    colorScheme: COLOR_SCHEMES.math
  },

  SIGN: {
    name: 'SIGN',
    syntax: 'SIGN(number)',
    params: [
      { name: 'number', desc: '符号を調べる数値' }
    ],
    description: '数値の符号を返します（正=1、ゼロ=0、負=-1）',
    category: 'math',
    examples: ['=SIGN(10)', '=SIGN(-5)', '=SIGN(0)'],
    colorScheme: COLOR_SCHEMES.math
  },

  EVEN: {
    name: 'EVEN',
    syntax: 'EVEN(number)',
    params: [
      { name: 'number', desc: '偶数に切り上げる数値' }
    ],
    description: '最も近い偶数に切り上げます',
    category: 'math',
    examples: ['=EVEN(3)', '=EVEN(4.5)'],
    colorScheme: COLOR_SCHEMES.math
  },

  ODD: {
    name: 'ODD',
    syntax: 'ODD(number)',
    params: [
      { name: 'number', desc: '奇数に切り上げる数値' }
    ],
    description: '最も近い奇数に切り上げます',
    category: 'math',
    examples: ['=ODD(2)', '=ODD(4.5)'],
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