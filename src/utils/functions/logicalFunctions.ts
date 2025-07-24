// 論理関数の定義
import { type FunctionDefinition } from '../types';
import { COLOR_SCHEMES } from '../colorSchemes';

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
  }
};