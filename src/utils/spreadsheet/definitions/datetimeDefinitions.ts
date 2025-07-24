// 日付・時刻関数の定義
import { type FunctionDefinition } from '../types/types';
import { COLOR_SCHEMES } from '../types/colorSchemes';

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
  },

  WEEKDAY: {
    name: 'WEEKDAY',
    syntax: 'WEEKDAY(serial_number, [return_type])',
    params: [
      { name: 'serial_number', desc: '日付' },
      { name: 'return_type', desc: '戻り値の種類（1-3）', optional: true }
    ],
    description: '日付の曜日を数値で返します',
    category: 'datetime',
    examples: ['=WEEKDAY("2024-01-01")', '=WEEKDAY(TODAY(),2)'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  DAYS: {
    name: 'DAYS',
    syntax: 'DAYS(end_date, start_date)',
    params: [
      { name: 'end_date', desc: '終了日' },
      { name: 'start_date', desc: '開始日' }
    ],
    description: '2つの日付間の日数を計算します',
    category: 'datetime',
    examples: ['=DAYS("2024-12-31","2024-01-01")', '=DAYS(B1,A1)'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  EDATE: {
    name: 'EDATE',
    syntax: 'EDATE(start_date, months)',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'months', desc: '加算する月数' }
    ],
    description: '指定した月数だけ前後の日付を返します',
    category: 'datetime',
    examples: ['=EDATE("2024-01-15",3)', '=EDATE(TODAY(),-6)'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  EOMONTH: {
    name: 'EOMONTH',
    syntax: 'EOMONTH(start_date, months)',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'months', desc: '加算する月数' }
    ],
    description: '指定した月数だけ前後の月末日を返します',
    category: 'datetime',
    examples: ['=EOMONTH("2024-01-15",0)', '=EOMONTH(TODAY(),1)'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  TIME: {
    name: 'TIME',
    syntax: 'TIME(hour, minute, second)',
    params: [
      { name: 'hour', desc: '時（0-23）' },
      { name: 'minute', desc: '分（0-59）' },
      { name: 'second', desc: '秒（0-59）' }
    ],
    description: '時、分、秒から時刻を作成します',
    category: 'datetime',
    examples: ['=TIME(14,30,0)', '=TIME(9,15,30)'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  HOUR: {
    name: 'HOUR',
    syntax: 'HOUR(serial_number)',
    params: [
      { name: 'serial_number', desc: '時刻' }
    ],
    description: '時刻から時を抽出します',
    category: 'datetime',
    examples: ['=HOUR("14:30:00")', '=HOUR(NOW())'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  MINUTE: {
    name: 'MINUTE',
    syntax: 'MINUTE(serial_number)',
    params: [
      { name: 'serial_number', desc: '時刻' }
    ],
    description: '時刻から分を抽出します',
    category: 'datetime',
    examples: ['=MINUTE("14:30:00")', '=MINUTE(NOW())'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  SECOND: {
    name: 'SECOND',
    syntax: 'SECOND(serial_number)',
    params: [
      { name: 'serial_number', desc: '時刻' }
    ],
    description: '時刻から秒を抽出します',
    category: 'datetime',
    examples: ['=SECOND("14:30:45")', '=SECOND(NOW())'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  WEEKNUM: {
    name: 'WEEKNUM',
    syntax: 'WEEKNUM(serial_number, [return_type])',
    params: [
      { name: 'serial_number', desc: '日付' },
      { name: 'return_type', desc: '週の開始曜日（1または2）', optional: true }
    ],
    description: '日付が年の第何週かを返します',
    category: 'datetime',
    examples: ['=WEEKNUM("2024-03-15")', '=WEEKNUM(TODAY(),2)'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  DAYS360: {
    name: 'DAYS360',
    syntax: 'DAYS360(start_date, end_date, [method])',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'end_date', desc: '終了日' },
      { name: 'method', desc: '計算方法（TRUE=ヨーロッパ方式）', optional: true }
    ],
    description: '360日法で2つの日付間の日数を計算します',
    category: 'datetime',
    examples: ['=DAYS360("2024-01-01","2024-12-31")', '=DAYS360(A1,B1,TRUE)'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  YEARFRAC: {
    name: 'YEARFRAC',
    syntax: 'YEARFRAC(start_date, end_date, [basis])',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'end_date', desc: '終了日' },
      { name: 'basis', desc: '日数計算の基準（0-4）', optional: true }
    ],
    description: '2つの日付間の年数を小数で返します',
    category: 'datetime',
    examples: ['=YEARFRAC("2024-01-01","2024-07-01")', '=YEARFRAC(A1,B1,1)'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  DATEVALUE: {
    name: 'DATEVALUE',
    syntax: 'DATEVALUE(date_text)',
    params: [
      { name: 'date_text', desc: '日付を表す文字列' }
    ],
    description: '日付文字列をシリアル値に変換します',
    category: 'datetime',
    examples: ['=DATEVALUE("2024-01-01")', '=DATEVALUE("2024年1月1日")'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  TIMEVALUE: {
    name: 'TIMEVALUE',
    syntax: 'TIMEVALUE(time_text)',
    params: [
      { name: 'time_text', desc: '時刻を表す文字列' }
    ],
    description: '時刻文字列をシリアル値に変換します',
    category: 'datetime',
    examples: ['=TIMEVALUE("14:30:00")', '=TIMEVALUE("2:30 PM")'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  ISOWEEKNUM: {
    name: 'ISOWEEKNUM',
    syntax: 'ISOWEEKNUM(date)',
    params: [
      { name: 'date', desc: '日付' }
    ],
    description: 'ISO 8601に基づく週番号を返します',
    category: 'datetime',
    examples: ['=ISOWEEKNUM("2024-01-01")', '=ISOWEEKNUM(TODAY())'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  'NETWORKDAYS.INTL': {
    name: 'NETWORKDAYS.INTL',
    syntax: 'NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'end_date', desc: '終了日' },
      { name: 'weekend', desc: '週末パラメータ', optional: true },
      { name: 'holidays', desc: '除外する休日のリスト', optional: true }
    ],
    description: 'カスタム週末を指定して稼働日数を計算します',
    category: 'datetime',
    examples: ['=NETWORKDAYS.INTL("2024-01-01","2024-01-31",11)', '=NETWORKDAYS.INTL(A1,B1,"0000011")'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  WORKDAY: {
    name: 'WORKDAY',
    syntax: 'WORKDAY(start_date, days, [holidays])',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'days', desc: '加算する稼働日数' },
      { name: 'holidays', desc: '除外する休日のリスト', optional: true }
    ],
    description: '指定した稼働日数だけ前後の日付を返します',
    category: 'datetime',
    examples: ['=WORKDAY("2024-01-01",10)', '=WORKDAY(TODAY(),5,C1:C5)'],
    colorScheme: COLOR_SCHEMES.datetime
  },

  'WORKDAY.INTL': {
    name: 'WORKDAY.INTL',
    syntax: 'WORKDAY.INTL(start_date, days, [weekend], [holidays])',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'days', desc: '加算する稼働日数' },
      { name: 'weekend', desc: '週末パラメータ', optional: true },
      { name: 'holidays', desc: '除外する休日のリスト', optional: true }
    ],
    description: 'カスタム週末を指定して稼働日を計算します',
    category: 'datetime',
    examples: ['=WORKDAY.INTL("2024-01-01",10,11)', '=WORKDAY.INTL(TODAY(),5,"0000011")'],
    colorScheme: COLOR_SCHEMES.datetime
  }
};