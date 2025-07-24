// その他の関数の定義
import { type FunctionDefinition } from '../types';
import { COLOR_SCHEMES } from '../colorSchemes';

export const OTHER_FUNCTIONS: Record<string, FunctionDefinition> = {
  LAMBDA: {
    name: 'LAMBDA',
    syntax: 'LAMBDA(parameter1, [parameter2], ..., calculation)',
    params: [
      { name: 'parameter1', desc: '最初のパラメータ' },
      { name: 'parameter2', desc: '追加のパラメータ', optional: true },
      { name: 'calculation', desc: '計算式' }
    ],
    description: 'カスタム関数を作成します',
    category: 'other',
    colorScheme: COLOR_SCHEMES.other
  },
  
  LET: {
    name: 'LET',
    syntax: 'LET(name1, value1, [name2, value2], ..., calculation)',
    params: [
      { name: 'name1', desc: '変数名1' },
      { name: 'value1', desc: '変数値1' },
      { name: 'name2', desc: '変数名2', optional: true },
      { name: 'value2', desc: '変数値2', optional: true },
      { name: 'calculation', desc: '計算式' }
    ],
    description: '計算内で変数を定義します',
    category: 'other',
    colorScheme: COLOR_SCHEMES.other
  }
};