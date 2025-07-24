// データベース関数の定義
import { type FunctionDefinition } from '../types';
import { COLOR_SCHEMES } from '../colorSchemes';

export const DATABASE_FUNCTIONS: Record<string, FunctionDefinition> = {
  DSUM: {
    name: 'DSUM',
    syntax: 'DSUM(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: '合計するフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致するレコードの合計を計算します',
    category: 'database',
    examples: ['=DSUM(A1:F10,"売上",H1:H2)', '=DSUM(A1:F10,3,H1:I2)'],
    colorScheme: COLOR_SCHEMES.database
  },

  DAVERAGE: {
    name: 'DAVERAGE',
    syntax: 'DAVERAGE(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: '平均するフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致するレコードの平均を計算します',
    category: 'database',
    examples: ['=DAVERAGE(A1:F10,"売上",H1:H2)'],
    colorScheme: COLOR_SCHEMES.database
  },

  DCOUNT: {
    name: 'DCOUNT',
    syntax: 'DCOUNT(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: 'カウントするフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致する数値セルの個数を数えます',
    category: 'database',
    examples: ['=DCOUNT(A1:F10,"売上",H1:H2)'],
    colorScheme: COLOR_SCHEMES.database
  },

  DCOUNTA: {
    name: 'DCOUNTA',
    syntax: 'DCOUNTA(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: 'カウントするフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致する空白以外のセルの個数を数えます',
    category: 'database',
    examples: ['=DCOUNTA(A1:F10,"商品名",H1:H2)'],
    colorScheme: COLOR_SCHEMES.database
  },

  DMAX: {
    name: 'DMAX',
    syntax: 'DMAX(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: '最大値を求めるフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致するレコードの最大値を求めます',
    category: 'database',
    examples: ['=DMAX(A1:F10,"売上",H1:H2)'],
    colorScheme: COLOR_SCHEMES.database
  },

  DMIN: {
    name: 'DMIN',
    syntax: 'DMIN(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: '最小値を求めるフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致するレコードの最小値を求めます',
    category: 'database',
    examples: ['=DMIN(A1:F10,"売上",H1:H2)'],
    colorScheme: COLOR_SCHEMES.database
  },

  DPRODUCT: {
    name: 'DPRODUCT',
    syntax: 'DPRODUCT(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: '積を求めるフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致するレコードの積を計算します',
    category: 'database',
    examples: ['=DPRODUCT(A1:F10,"数量",H1:H2)'],
    colorScheme: COLOR_SCHEMES.database
  },

  DSTDEV: {
    name: 'DSTDEV',
    syntax: 'DSTDEV(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: '標準偏差を求めるフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致するレコードの標準偏差（標本）を計算します',
    category: 'database',
    examples: ['=DSTDEV(A1:F10,"売上",H1:H2)'],
    colorScheme: COLOR_SCHEMES.database
  },

  DSTDEVP: {
    name: 'DSTDEVP',
    syntax: 'DSTDEVP(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: '標準偏差を求めるフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致するレコードの標準偏差（母集団）を計算します',
    category: 'database',
    examples: ['=DSTDEVP(A1:F10,"売上",H1:H2)'],
    colorScheme: COLOR_SCHEMES.database
  },

  DVAR: {
    name: 'DVAR',
    syntax: 'DVAR(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: '分散を求めるフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致するレコードの分散（標本）を計算します',
    category: 'database',
    examples: ['=DVAR(A1:F10,"売上",H1:H2)'],
    colorScheme: COLOR_SCHEMES.database
  },

  DVARP: {
    name: 'DVARP',
    syntax: 'DVARP(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: '分散を求めるフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致するレコードの分散（母集団）を計算します',
    category: 'database',
    examples: ['=DVARP(A1:F10,"売上",H1:H2)'],
    colorScheme: COLOR_SCHEMES.database
  },

  DGET: {
    name: 'DGET',
    syntax: 'DGET(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲（列見出しを含む）' },
      { name: 'field', desc: '取得するフィールド名または列番号' },
      { name: 'criteria', desc: '条件範囲（見出しと条件を含む）' }
    ],
    description: 'データベースから条件に一致する単一の値を取得します',
    category: 'database',
    examples: ['=DGET(A1:F10,"商品名",H1:H2)'],
    colorScheme: COLOR_SCHEMES.database
  }
};