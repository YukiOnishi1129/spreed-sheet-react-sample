// 情報関数の定義
import { type FunctionDefinition, COLOR_SCHEMES } from '../functionDefinitions';

export const INFORMATION_FUNCTIONS: Record<string, FunctionDefinition> = {
  ISBLANK: {
    name: 'ISBLANK',
    syntax: 'ISBLANK(value)',
    params: [
      { name: 'value', desc: 'テストする値' }
    ],
    description: 'セルが空白かどうかを判定します',
    category: 'information',
    colorScheme: COLOR_SCHEMES.information
  },

  ISERROR: {
    name: 'ISERROR',
    syntax: 'ISERROR(value)',
    params: [
      { name: 'value', desc: 'テストする値' }
    ],
    description: '値がエラーかどうかを判定します',
    category: 'information',
    colorScheme: COLOR_SCHEMES.information
  },

  ISNUMBER: {
    name: 'ISNUMBER',
    syntax: 'ISNUMBER(value)',
    params: [
      { name: 'value', desc: 'テストする値' }
    ],
    description: '値が数値かどうかを判定します',
    category: 'information',
    colorScheme: COLOR_SCHEMES.information
  }
};