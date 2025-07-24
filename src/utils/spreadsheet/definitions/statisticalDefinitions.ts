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
  }
};