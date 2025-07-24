// 文字列操作関数の定義
import { type FunctionDefinition } from '../types';
import { COLOR_SCHEMES } from '../colorSchemes';

export const TEXT_FUNCTIONS: Record<string, FunctionDefinition> = {
  CONCATENATE: {
    name: 'CONCATENATE',
    syntax: 'CONCATENATE(text1, [text2], ...)',
    params: [
      { name: 'text1', desc: '結合する最初のテキスト' },
      { name: 'text2', desc: '結合する2番目のテキスト', optional: true }
    ],
    description: '複数のテキストを結合します',
    category: 'text',
    examples: ['=CONCATENATE("Hello"," ","World")', '=CONCATENATE(A1,B1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  LEFT: {
    name: 'LEFT',
    syntax: 'LEFT(text, [num_chars])',
    params: [
      { name: 'text', desc: '文字列' },
      { name: 'num_chars', desc: '取り出す文字数（省略時は1）', optional: true }
    ],
    description: '文字列の左端から指定した文字数を取り出します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  },

  RIGHT: {
    name: 'RIGHT',
    syntax: 'RIGHT(text, [num_chars])',
    params: [
      { name: 'text', desc: '文字列' },
      { name: 'num_chars', desc: '取り出す文字数（省略時は1）', optional: true }
    ],
    description: '文字列の右端から指定した文字数を取り出します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  },

  MID: {
    name: 'MID',
    syntax: 'MID(text, start_num, num_chars)',
    params: [
      { name: 'text', desc: '文字列' },
      { name: 'start_num', desc: '開始位置' },
      { name: 'num_chars', desc: '取り出す文字数' }
    ],
    description: '文字列の指定した位置から指定した文字数を取り出します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  },

  LEN: {
    name: 'LEN',
    syntax: 'LEN(text)',
    params: [
      { name: 'text', desc: '長さを求める文字列' }
    ],
    description: '文字列の長さを返します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  }
};