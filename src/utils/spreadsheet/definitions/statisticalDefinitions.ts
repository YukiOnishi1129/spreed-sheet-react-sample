// 統計関数の定義
import { type FunctionDefinition } from '../types/types';
import { COLOR_SCHEMES } from '../types/colorSchemes';

export const STATISTICAL_FUNCTIONS: Record<string, FunctionDefinition> = {
  AVERAGE: {
    name: 'AVERAGE',
    syntax: 'AVERAGE(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '平均を求める最初の数値または範囲' },
      { name: 'number2', desc: '平均を求める追加の数値または範囲', optional: true }
    ],
    description: '数値の平均を計算します',
    category: 'statistical',
    colorScheme: COLOR_SCHEMES.statistical
  },

  COUNTA: {
    name: 'COUNTA',
    syntax: 'COUNTA(value1, [value2], ...)',
    params: [
      { name: 'value1', desc: '値をカウントする最初の範囲' },
      { name: 'value2', desc: '値をカウントする追加の範囲', optional: true }
    ],
    description: '空白でないセルの個数を数えます',
    category: 'statistical',
    examples: ['=COUNTA(A1:A10)', '=COUNTA(A1:A10,C1:C10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  COUNTBLANK: {
    name: 'COUNTBLANK',
    syntax: 'COUNTBLANK(range)',
    params: [
      { name: 'range', desc: '空白セルをカウントする範囲' }
    ],
    description: '空白セルの個数を数えます',
    category: 'statistical',
    examples: ['=COUNTBLANK(A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  MEDIAN: {
    name: 'MEDIAN',
    syntax: 'MEDIAN(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '中央値を求める最初の数値または範囲' },
      { name: 'number2', desc: '中央値を求める追加の数値または範囲', optional: true }
    ],
    description: '数値の中央値を返します',
    category: 'statistical',
    examples: ['=MEDIAN(A1:A10)', '=MEDIAN(1,2,3,4,5)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  MODE: {
    name: 'MODE',
    syntax: 'MODE(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '最頻値を求める最初の数値または範囲' },
      { name: 'number2', desc: '最頻値を求める追加の数値または範囲', optional: true }
    ],
    description: '最も頻繁に出現する値を返します（旧バージョン）',
    category: 'statistical',
    examples: ['=MODE(A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  MODE_SNGL: {
    name: 'MODE.SNGL',
    syntax: 'MODE.SNGL(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '最頻値を求める最初の数値または範囲' },
      { name: 'number2', desc: '最頻値を求める追加の数値または範囲', optional: true }
    ],
    description: '最も頻繁に出現する値を返します',
    category: 'statistical',
    examples: ['=MODE.SNGL(A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  MODE_MULT: {
    name: 'MODE.MULT',
    syntax: 'MODE.MULT(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '最頻値を求める最初の数値または範囲' },
      { name: 'number2', desc: '最頻値を求める追加の数値または範囲', optional: true }
    ],
    description: '最も頻繁に出現する値を縦方向の配列で返します',
    category: 'statistical',
    examples: ['=MODE.MULT(A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  COUNT: {
    name: 'COUNT',
    syntax: 'COUNT(value1, [value2], ...)',
    params: [
      { name: 'value1', desc: '数値をカウントする最初の値または範囲' },
      { name: 'value2', desc: '数値をカウントする追加の値または範囲', optional: true }
    ],
    description: '数値が入力されているセルの個数を数えます',
    category: 'statistical',
    colorScheme: COLOR_SCHEMES.statistical
  },

  MAX: {
    name: 'MAX',
    syntax: 'MAX(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '最大値を求める最初の数値または範囲' },
      { name: 'number2', desc: '最大値を求める追加の数値または範囲', optional: true }
    ],
    description: '最大値を返します',
    category: 'statistical',
    colorScheme: COLOR_SCHEMES.statistical
  },

  MIN: {
    name: 'MIN',
    syntax: 'MIN(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '最小値を求める最初の数値または範囲' },
      { name: 'number2', desc: '最小値を求める追加の数値または範囲', optional: true }
    ],
    description: '最小値を返します',
    category: 'statistical',
    colorScheme: COLOR_SCHEMES.statistical
  },

  STDEV: {
    name: 'STDEV',
    syntax: 'STDEV(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '標準偏差を求める最初の数値または範囲' },
      { name: 'number2', desc: '標準偏差を求める追加の数値または範囲', optional: true }
    ],
    description: '標本標準偏差を計算します（旧バージョン）',
    category: 'statistical',
    examples: ['=STDEV(A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  STDEV_S: {
    name: 'STDEV.S',
    syntax: 'STDEV.S(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '標準偏差を求める最初の数値または範囲' },
      { name: 'number2', desc: '標準偏差を求める追加の数値または範囲', optional: true }
    ],
    description: '標本標準偏差を計算します',
    category: 'statistical',
    examples: ['=STDEV.S(A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  STDEV_P: {
    name: 'STDEV.P',
    syntax: 'STDEV.P(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '標準偏差を求める最初の数値または範囲' },
      { name: 'number2', desc: '標準偏差を求める追加の数値または範囲', optional: true }
    ],
    description: '母標準偏差を計算します',
    category: 'statistical',
    examples: ['=STDEV.P(A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  VAR: {
    name: 'VAR',
    syntax: 'VAR(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '分散を求める最初の数値または範囲' },
      { name: 'number2', desc: '分散を求める追加の数値または範囲', optional: true }
    ],
    description: '標本分散を計算します（旧バージョン）',
    category: 'statistical',
    examples: ['=VAR(A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  VAR_S: {
    name: 'VAR.S',
    syntax: 'VAR.S(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '分散を求める最初の数値または範囲' },
      { name: 'number2', desc: '分散を求める追加の数値または範囲', optional: true }
    ],
    description: '標本分散を計算します',
    category: 'statistical',
    examples: ['=VAR.S(A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  VAR_P: {
    name: 'VAR.P',
    syntax: 'VAR.P(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '分散を求める最初の数値または範囲' },
      { name: 'number2', desc: '分散を求める追加の数値または範囲', optional: true }
    ],
    description: '母分散を計算します',
    category: 'statistical',
    examples: ['=VAR.P(A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  RANK: {
    name: 'RANK',
    syntax: 'RANK(number, ref, [order])',
    params: [
      { name: 'number', desc: '順位を求める値' },
      { name: 'ref', desc: '数値の配列または範囲' },
      { name: 'order', desc: '順位の順序（0=降順、1=昇順）', optional: true }
    ],
    description: '数値の順位を返します（旧バージョン）',
    category: 'statistical',
    examples: ['=RANK(A1,A1:A10)', '=RANK(A1,A1:A10,0)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  RANK_EQ: {
    name: 'RANK.EQ',
    syntax: 'RANK.EQ(number, ref, [order])',
    params: [
      { name: 'number', desc: '順位を求める値' },
      { name: 'ref', desc: '数値の配列または範囲' },
      { name: 'order', desc: '順位の順序（0=降順、1=昇順）', optional: true }
    ],
    description: '数値の順位を返します（同じ値には同じ順位）',
    category: 'statistical',
    examples: ['=RANK.EQ(A1,A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  RANK_AVG: {
    name: 'RANK.AVG',
    syntax: 'RANK.AVG(number, ref, [order])',
    params: [
      { name: 'number', desc: '順位を求める値' },
      { name: 'ref', desc: '数値の配列または範囲' },
      { name: 'order', desc: '順位の順序（0=降順、1=昇順）', optional: true }
    ],
    description: '数値の順位を返します（同じ値には平均順位）',
    category: 'statistical',
    examples: ['=RANK.AVG(A1,A1:A10)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  LARGE: {
    name: 'LARGE',
    syntax: 'LARGE(array, k)',
    params: [
      { name: 'array', desc: 'データの配列または範囲' },
      { name: 'k', desc: '何番目に大きい値を取得するか' }
    ],
    description: 'k番目に大きい値を返します',
    category: 'statistical',
    examples: ['=LARGE(A1:A10,1)', '=LARGE(A1:A10,3)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  SMALL: {
    name: 'SMALL',
    syntax: 'SMALL(array, k)',
    params: [
      { name: 'array', desc: 'データの配列または範囲' },
      { name: 'k', desc: '何番目に小さい値を取得するか' }
    ],
    description: 'k番目に小さい値を返します',
    category: 'statistical',
    examples: ['=SMALL(A1:A10,1)', '=SMALL(A1:A10,3)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  PERCENTILE: {
    name: 'PERCENTILE',
    syntax: 'PERCENTILE(array, k)',
    params: [
      { name: 'array', desc: 'データの配列または範囲' },
      { name: 'k', desc: 'パーセンタイル値（0～1）' }
    ],
    description: 'k番目のパーセンタイルを返します（旧バージョン）',
    category: 'statistical',
    examples: ['=PERCENTILE(A1:A10,0.9)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  PERCENTILE_INC: {
    name: 'PERCENTILE.INC',
    syntax: 'PERCENTILE.INC(array, k)',
    params: [
      { name: 'array', desc: 'データの配列または範囲' },
      { name: 'k', desc: 'パーセンタイル値（0～1）' }
    ],
    description: 'k番目のパーセンタイルを返します（0と1を含む）',
    category: 'statistical',
    examples: ['=PERCENTILE.INC(A1:A10,0.9)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  PERCENTILE_EXC: {
    name: 'PERCENTILE.EXC',
    syntax: 'PERCENTILE.EXC(array, k)',
    params: [
      { name: 'array', desc: 'データの配列または範囲' },
      { name: 'k', desc: 'パーセンタイル値（0～1の間）' }
    ],
    description: 'k番目のパーセンタイルを返します（0と1を除く）',
    category: 'statistical',
    examples: ['=PERCENTILE.EXC(A1:A10,0.9)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  QUARTILE: {
    name: 'QUARTILE',
    syntax: 'QUARTILE(array, quart)',
    params: [
      { name: 'array', desc: 'データの配列または範囲' },
      { name: 'quart', desc: '四分位数（0=最小値、1=第1四分位、2=中央値、3=第3四分位、4=最大値）' }
    ],
    description: '四分位数を返します（旧バージョン）',
    category: 'statistical',
    examples: ['=QUARTILE(A1:A10,1)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  QUARTILE_INC: {
    name: 'QUARTILE.INC',
    syntax: 'QUARTILE.INC(array, quart)',
    params: [
      { name: 'array', desc: 'データの配列または範囲' },
      { name: 'quart', desc: '四分位数（0=最小値、1=第1四分位、2=中央値、3=第3四分位、4=最大値）' }
    ],
    description: '四分位数を返します（0と4を含む）',
    category: 'statistical',
    examples: ['=QUARTILE.INC(A1:A10,1)'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  QUARTILE_EXC: {
    name: 'QUARTILE.EXC',
    syntax: 'QUARTILE.EXC(array, quart)',
    params: [
      { name: 'array', desc: 'データの配列または範囲' },
      { name: 'quart', desc: '四分位数（1=第1四分位、2=中央値、3=第3四分位）' }
    ],
    description: '四分位数を返します（0と4を除く）',
    category: 'statistical',
    examples: ['=QUARTILE.EXC(A1:A10,1)'],
    colorScheme: COLOR_SCHEMES.statistical
  }
};