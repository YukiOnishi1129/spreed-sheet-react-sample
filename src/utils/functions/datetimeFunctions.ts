// 日付・時刻関数の定義
import { type FunctionDefinition } from '../types';
import { COLOR_SCHEMES } from '../colorSchemes';

export const DATETIME_FUNCTIONS: Record<string, FunctionDefinition> = {
  TODAY: {
    name: 'TODAY',
    syntax: 'TODAY()',
    params: [],
    description: '今日の日付を返します',
    category: 'datetime',
    colorScheme: COLOR_SCHEMES.datetime
  },

  NOW: {
    name: 'NOW',
    syntax: 'NOW()',
    params: [],
    description: '現在の日付と時刻を返します',
    category: 'datetime',
    colorScheme: COLOR_SCHEMES.datetime
  },

  DATE: {
    name: 'DATE',
    syntax: 'DATE(year, month, day)',
    params: [
      { name: 'year', desc: '年' },
      { name: 'month', desc: '月' },
      { name: 'day', desc: '日' }
    ],
    description: '年、月、日から日付を作成します',
    category: 'datetime',
    colorScheme: COLOR_SCHEMES.datetime
  },

  YEAR: {
    name: 'YEAR',
    syntax: 'YEAR(serial_number)',
    params: [
      { name: 'serial_number', desc: '年を取得したい日付' }
    ],
    description: '日付から年を取得します',
    category: 'datetime',
    colorScheme: COLOR_SCHEMES.datetime
  },

  MONTH: {
    name: 'MONTH',
    syntax: 'MONTH(serial_number)',
    params: [
      { name: 'serial_number', desc: '月を取得したい日付' }
    ],
    description: '日付から月を取得します',
    category: 'datetime',
    colorScheme: COLOR_SCHEMES.datetime
  },

  DAY: {
    name: 'DAY',
    syntax: 'DAY(serial_number)',
    params: [
      { name: 'serial_number', desc: '日を取得したい日付' }
    ],
    description: '日付から日を取得します',
    category: 'datetime',
    colorScheme: COLOR_SCHEMES.datetime
  },

  DATEDIF: {
    name: 'DATEDIF',
    syntax: 'DATEDIF(start_date, end_date, unit)',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'end_date', desc: '終了日' },
      { name: 'unit', desc: '単位（"Y":年、"M":月、"D":日など）' }
    ],
    description: '2つの日付間の差を計算します',
    category: 'datetime',
    examples: ['=DATEDIF(A1,B1,"D")', '=DATEDIF(A1,TODAY(),"Y")'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  NETWORKDAYS: {
    name: 'NETWORKDAYS',
    syntax: 'NETWORKDAYS(start_date, end_date, [holidays])',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'end_date', desc: '終了日' },
      { name: 'holidays', desc: '休日のリスト', optional: true }
    ],
    description: '2つの日付間の稼働日数を計算します',
    category: 'datetime',
    examples: ['=NETWORKDAYS(A1,B1)', '=NETWORKDAYS(A1,B1,C1:C10)'],
    colorScheme: COLOR_SCHEMES.datetime
  }
};