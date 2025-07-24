// 論理関数の定義
import { type FunctionDefinition } from '../types/types';
import { COLOR_SCHEMES } from '../types/colorSchemes';

export const LOGICAL_FUNCTIONS: Record<string, FunctionDefinition> = {
  IF: {
    name: 'IF',
    syntax: 'IF(logical_test, value_if_true, value_if_false)',
    params: [
      { name: 'logical_test', desc: '判定する条件' },
      { name: 'value_if_true', desc: '条件がTRUEの場合の値' },
      { name: 'value_if_false', desc: '条件がFALSEの場合の値' }
    ],
    description: '条件によって異なる値を返します',
    category: 'logical',
    examples: ['=IF(A1>10,"大","小")', '=IF(B1="合格",100,0)'],
    colorScheme: COLOR_SCHEMES.logical
  },

  AND: {
    name: 'AND',
    syntax: 'AND(logical1, [logical2], ...)',
    params: [
      { name: 'logical1', desc: '1つ目の論理値または条件' },
      { name: 'logical2', desc: '2つ目の論理値または条件', optional: true }
    ],
    description: 'すべての条件がTRUEの場合にTRUEを返します',
    category: 'logical',
    colorScheme: COLOR_SCHEMES.logical
  },

  OR: {
    name: 'OR',
    syntax: 'OR(logical1, [logical2], ...)',
    params: [
      { name: 'logical1', desc: '1つ目の論理値または条件' },
      { name: 'logical2', desc: '2つ目の論理値または条件', optional: true }
    ],
    description: 'いずれかの条件がTRUEの場合にTRUEを返します',
    category: 'logical',
    colorScheme: COLOR_SCHEMES.logical
  },

  NOT: {
    name: 'NOT',
    syntax: 'NOT(logical)',
    params: [
      { name: 'logical', desc: '反転する論理値' }
    ],
    description: '論理値を反転します',
    category: 'logical',
    colorScheme: COLOR_SCHEMES.logical
  },

  TRUE: {
    name: 'TRUE',
    syntax: 'TRUE()',
    params: [],
    description: '論理値TRUEを返します',
    category: 'logical',
    examples: ['=TRUE()', '=IF(A1>0,TRUE(),FALSE())'],
    colorScheme: COLOR_SCHEMES.logical
  },

  FALSE: {
    name: 'FALSE',
    syntax: 'FALSE()',
    params: [],
    description: '論理値FALSEを返します',
    category: 'logical',
    examples: ['=FALSE()', '=IF(A1<0,TRUE(),FALSE())'],
    colorScheme: COLOR_SCHEMES.logical
  },

  IFERROR: {
    name: 'IFERROR',
    syntax: 'IFERROR(value, value_if_error)',
    params: [
      { name: 'value', desc: '評価する値または式' },
      { name: 'value_if_error', desc: 'エラーの場合に返す値' }
    ],
    description: 'エラーの場合に指定した値を返します',
    category: 'logical',
    examples: ['=IFERROR(A1/B1,0)', '=IFERROR(VLOOKUP(A1,B:C,2,FALSE),"見つかりません")'],
    colorScheme: COLOR_SCHEMES.logical
  },

  IFNA: {
    name: 'IFNA',
    syntax: 'IFNA(value, value_if_na)',
    params: [
      { name: 'value', desc: '評価する値または式' },
      { name: 'value_if_na', desc: '#N/Aエラーの場合に返す値' }
    ],
    description: '#N/Aエラーの場合に指定した値を返します',
    category: 'logical',
    examples: ['=IFNA(VLOOKUP(A1,B:C,2,FALSE),"該当なし")', '=IFNA(MATCH(A1,B:B,0),0)'],
    colorScheme: COLOR_SCHEMES.logical
  },

  IFS: {
    name: 'IFS',
    syntax: 'IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)',
    params: [
      { name: 'logical_test1', desc: '1つ目の条件' },
      { name: 'value_if_true1', desc: '1つ目の条件がTRUEの場合の値' },
      { name: 'logical_test2', desc: '2つ目の条件', optional: true },
      { name: 'value_if_true2', desc: '2つ目の条件がTRUEの場合の値', optional: true }
    ],
    description: '複数の条件を順に評価し、最初にTRUEになった値を返します',
    category: 'logical',
    examples: ['=IFS(A1>=90,"優",A1>=80,"良",A1>=70,"可",TRUE,"不可")', '=IFS(A1="月",1,A1="火",2,A1="水",3)'],
    colorScheme: COLOR_SCHEMES.logical
  },

  XOR: {
    name: 'XOR',
    syntax: 'XOR(logical1, [logical2], ...)',
    params: [
      { name: 'logical1', desc: '1つ目の論理値または条件' },
      { name: 'logical2', desc: '2つ目の論理値または条件', optional: true }
    ],
    description: '奇数個の条件がTRUEの場合にTRUEを返します（排他的論理和）',
    category: 'logical',
    examples: ['=XOR(TRUE,FALSE)', '=XOR(A1>0,B1>0,C1>0)'],
    colorScheme: COLOR_SCHEMES.logical
  }
};