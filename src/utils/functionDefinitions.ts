// Excel/Google Sheets関数の詳細定義
// spreadsheet-functions-reference.mdに基づく完全な関数リスト

export interface FunctionParam {
  name: string;
  desc: string;
  optional?: boolean;
}

export interface FunctionDefinition {
  name: string;
  syntax: string;
  params: FunctionParam[];
  description: string;
  category: string;
  examples?: string[];
  colorScheme?: {
    bg: string;
    fc: string;
  };
}

// カテゴリ別の色定義
export const COLOR_SCHEMES = {
  math: { bg: "#FFE0B2", fc: "#D84315" },         // 数学・三角関数 - オレンジ
  statistical: { bg: "#FFE0B2", fc: "#D84315" },  // 統計関数 - オレンジ
  text: { bg: "#FCE4EC", fc: "#C2185B" },         // 文字列関数 - ピンク
  datetime: { bg: "#F3E5F5", fc: "#7B1FA2" },     // 日付・時刻関数 - 紫
  logical: { bg: "#E8F5E8", fc: "#2E7D32" },      // 論理関数 - 緑
  lookup: { bg: "#E3F2FD", fc: "#1976D2" },       // 検索・参照関数 - 青
  information: { bg: "#F5F5F5", fc: "#616161" },   // 情報関数 - グレー
  database: { bg: "#FFF8E1", fc: "#F57C00" },     // データベース関数 - 薄い黄
  financial: { bg: "#FFFDE7", fc: "#F57F17" },    // 財務関数 - 黄色
  engineering: { bg: "#E0F2F1", fc: "#00796B" },  // エンジニアリング関数 - 青緑
  cube: { bg: "#F5F5F5", fc: "#616161" },         // キューブ関数 - グレー
  web: { bg: "#E8EAF6", fc: "#3F51B5" },          // Web関数 - インディゴ
  googlesheets: { bg: "#E8F5E9", fc: "#43A047" }, // Google Sheets専用 - 緑
  other: { bg: "#F5F5F5", fc: "#616161" }         // その他 - グレー
};

export const FUNCTION_DEFINITIONS: Record<string, FunctionDefinition> = {
  // 01. 数学・三角関数
  SUM: {
    name: 'SUM',
    syntax: 'SUM(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '合計する最初の数値または範囲' },
      { name: 'number2', desc: '合計する追加の数値または範囲', optional: true }
    ],
    description: '数値の合計を計算します',
    category: 'math',
    examples: ['=SUM(A1:A10)', '=SUM(1,2,3,4,5)', '=SUM(A1:A5,C1:C5)'],
    colorScheme: COLOR_SCHEMES.math
  },
  
  SUMIF: {
    name: 'SUMIF',
    syntax: 'SUMIF(range, criteria, [sum_range])',
    params: [
      { name: 'range', desc: '条件を評価する範囲' },
      { name: 'criteria', desc: '条件（">10"、"=商品A"など）' },
      { name: 'sum_range', desc: '合計する範囲（省略時はrangeを使用）', optional: true }
    ],
    description: '条件に一致するセルの合計を計算します',
    category: 'math',
    examples: ['=SUMIF(A1:A10,">100")', '=SUMIF(B1:B10,"商品A",C1:C10)'],
    colorScheme: COLOR_SCHEMES.math
  },

  SUMIFS: {
    name: 'SUMIFS',
    syntax: 'SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    params: [
      { name: 'sum_range', desc: '合計する範囲' },
      { name: 'criteria_range1', desc: '1つ目の条件範囲' },
      { name: 'criteria1', desc: '1つ目の条件' },
      { name: 'criteria_range2', desc: '2つ目の条件範囲', optional: true },
      { name: 'criteria2', desc: '2つ目の条件', optional: true }
    ],
    description: '複数条件に一致するセルの合計を計算します',
    category: 'math',
    examples: ['=SUMIFS(C1:C10,A1:A10,">100",B1:B10,"商品A")'],
    colorScheme: COLOR_SCHEMES.math
  },

  PRODUCT: {
    name: 'PRODUCT',
    syntax: 'PRODUCT(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '積を求める最初の数値' },
      { name: 'number2', desc: '積を求める追加の数値', optional: true }
    ],
    description: '数値の積を計算します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  SQRT: {
    name: 'SQRT',
    syntax: 'SQRT(number)',
    params: [
      { name: 'number', desc: '平方根を求める数値' }
    ],
    description: '平方根を計算します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  POWER: {
    name: 'POWER',
    syntax: 'POWER(number, power)',
    params: [
      { name: 'number', desc: '底となる数値' },
      { name: 'power', desc: '指数' }
    ],
    description: 'べき乗を計算します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  ABS: {
    name: 'ABS',
    syntax: 'ABS(number)',
    params: [
      { name: 'number', desc: '絶対値を求める数値' }
    ],
    description: '絶対値を返します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  ROUND: {
    name: 'ROUND',
    syntax: 'ROUND(number, num_digits)',
    params: [
      { name: 'number', desc: '四捨五入する数値' },
      { name: 'num_digits', desc: '桁数（正:小数点以下、負:整数部分）' }
    ],
    description: '指定桁数で四捨五入します',
    category: 'math',
    examples: ['=ROUND(3.14159,2)', '=ROUND(123.456,-1)'],
    colorScheme: COLOR_SCHEMES.math
  },

  ROUNDUP: {
    name: 'ROUNDUP',
    syntax: 'ROUNDUP(number, num_digits)',
    params: [
      { name: 'number', desc: '切り上げる数値' },
      { name: 'num_digits', desc: '桁数' }
    ],
    description: '指定桁数で切り上げます',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  ROUNDDOWN: {
    name: 'ROUNDDOWN',
    syntax: 'ROUNDDOWN(number, num_digits)',
    params: [
      { name: 'number', desc: '切り下げる数値' },
      { name: 'num_digits', desc: '桁数' }
    ],
    description: '指定桁数で切り下げます',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  CEILING: {
    name: 'CEILING',
    syntax: 'CEILING(number, significance)',
    params: [
      { name: 'number', desc: '切り上げる数値' },
      { name: 'significance', desc: '基準値' }
    ],
    description: '基準値の倍数に切り上げます',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  FLOOR: {
    name: 'FLOOR',
    syntax: 'FLOOR(number, significance)',
    params: [
      { name: 'number', desc: '切り下げる数値' },
      { name: 'significance', desc: '基準値' }
    ],
    description: '基準値の倍数に切り下げます',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  MOD: {
    name: 'MOD',
    syntax: 'MOD(number, divisor)',
    params: [
      { name: 'number', desc: '被除数' },
      { name: 'divisor', desc: '除数' }
    ],
    description: '剰余を返します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  INT: {
    name: 'INT',
    syntax: 'INT(number)',
    params: [
      { name: 'number', desc: '整数部分を取り出す数値' }
    ],
    description: '整数部分を返します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  RAND: {
    name: 'RAND',
    syntax: 'RAND()',
    params: [],
    description: '0以上1未満の乱数を生成します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  RANDBETWEEN: {
    name: 'RANDBETWEEN',
    syntax: 'RANDBETWEEN(bottom, top)',
    params: [
      { name: 'bottom', desc: '最小値' },
      { name: 'top', desc: '最大値' }
    ],
    description: '指定範囲の整数乱数を生成します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  // 三角関数
  SIN: {
    name: 'SIN',
    syntax: 'SIN(number)',
    params: [
      { name: 'number', desc: '角度（ラジアン）' }
    ],
    description: '正弦を計算します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  COS: {
    name: 'COS',
    syntax: 'COS(number)',
    params: [
      { name: 'number', desc: '角度（ラジアン）' }
    ],
    description: '余弦を計算します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  TAN: {
    name: 'TAN',
    syntax: 'TAN(number)',
    params: [
      { name: 'number', desc: '角度（ラジアン）' }
    ],
    description: '正接を計算します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  PI: {
    name: 'PI',
    syntax: 'PI()',
    params: [],
    description: '円周率を返します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  // 02. 統計関数
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

  AVERAGEIF: {
    name: 'AVERAGEIF',
    syntax: 'AVERAGEIF(range, criteria, [average_range])',
    params: [
      { name: 'range', desc: '条件を評価する範囲' },
      { name: 'criteria', desc: '条件' },
      { name: 'average_range', desc: '平均を計算する範囲', optional: true }
    ],
    description: '条件に一致するセルの平均を計算します',
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

  COUNTA: {
    name: 'COUNTA',
    syntax: 'COUNTA(value1, [value2], ...)',
    params: [
      { name: 'value1', desc: 'カウントする最初の値または範囲' },
      { name: 'value2', desc: 'カウントする追加の値または範囲', optional: true }
    ],
    description: '空白でないセルの個数を数えます',
    category: 'statistical',
    colorScheme: COLOR_SCHEMES.statistical
  },

  COUNTIF: {
    name: 'COUNTIF',
    syntax: 'COUNTIF(range, criteria)',
    params: [
      { name: 'range', desc: 'カウント対象の範囲' },
      { name: 'criteria', desc: 'カウント条件' }
    ],
    description: '条件に一致するセルの個数を数えます',
    category: 'statistical',
    examples: ['=COUNTIF(A1:A10,">5")', '=COUNTIF(B1:B10,"合格")'],
    colorScheme: COLOR_SCHEMES.statistical
  },

  COUNTIFS: {
    name: 'COUNTIFS',
    syntax: 'COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    params: [
      { name: 'criteria_range1', desc: '1つ目の条件範囲' },
      { name: 'criteria1', desc: '1つ目の条件' },
      { name: 'criteria_range2', desc: '2つ目の条件範囲', optional: true },
      { name: 'criteria2', desc: '2つ目の条件', optional: true }
    ],
    description: '複数条件に一致するセルの個数を数えます',
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

  MEDIAN: {
    name: 'MEDIAN',
    syntax: 'MEDIAN(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '中央値を求める最初の数値' },
      { name: 'number2', desc: '中央値を求める追加の数値', optional: true }
    ],
    description: '中央値を返します',
    category: 'statistical',
    colorScheme: COLOR_SCHEMES.statistical
  },

  MODE: {
    name: 'MODE',
    syntax: 'MODE(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '最頻値を求める最初の数値' },
      { name: 'number2', desc: '最頻値を求める追加の数値', optional: true }
    ],
    description: '最頻値を返します',
    category: 'statistical',
    colorScheme: COLOR_SCHEMES.statistical
  },

  STDEV: {
    name: 'STDEV',
    syntax: 'STDEV(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '標準偏差を求める最初の数値' },
      { name: 'number2', desc: '標準偏差を求める追加の数値', optional: true }
    ],
    description: '標準偏差を計算します（標本）',
    category: 'statistical',
    colorScheme: COLOR_SCHEMES.statistical
  },

  VAR: {
    name: 'VAR',
    syntax: 'VAR(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '分散を求める最初の数値' },
      { name: 'number2', desc: '分散を求める追加の数値', optional: true }
    ],
    description: '分散を計算します（標本）',
    category: 'statistical',
    colorScheme: COLOR_SCHEMES.statistical
  },

  // 03. 文字列操作関数
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
  },

  UPPER: {
    name: 'UPPER',
    syntax: 'UPPER(text)',
    params: [
      { name: 'text', desc: '大文字に変換する文字列' }
    ],
    description: '文字列を大文字に変換します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  },

  LOWER: {
    name: 'LOWER',
    syntax: 'LOWER(text)',
    params: [
      { name: 'text', desc: '小文字に変換する文字列' }
    ],
    description: '文字列を小文字に変換します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  },

  TRIM: {
    name: 'TRIM',
    syntax: 'TRIM(text)',
    params: [
      { name: 'text', desc: 'トリムする文字列' }
    ],
    description: '文字列から余分なスペースを削除します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  },

  TEXT: {
    name: 'TEXT',
    syntax: 'TEXT(value, format_text)',
    params: [
      { name: 'value', desc: '書式を適用する数値' },
      { name: 'format_text', desc: '書式パターン（例："0,000"、"yyyy/mm/dd"）' }
    ],
    description: '数値を指定した書式の文字列に変換します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  },

  FIND: {
    name: 'FIND',
    syntax: 'FIND(find_text, within_text, [start_num])',
    params: [
      { name: 'find_text', desc: '検索する文字列' },
      { name: 'within_text', desc: '検索対象の文字列' },
      { name: 'start_num', desc: '開始位置', optional: true }
    ],
    description: '文字列内で別の文字列を検索し、位置を返します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  },

  REPLACE: {
    name: 'REPLACE',
    syntax: 'REPLACE(old_text, start_num, num_chars, new_text)',
    params: [
      { name: 'old_text', desc: '元の文字列' },
      { name: 'start_num', desc: '置換開始位置' },
      { name: 'num_chars', desc: '置換する文字数' },
      { name: 'new_text', desc: '新しい文字列' }
    ],
    description: '文字列の一部を別の文字列に置換します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  },

  SUBSTITUTE: {
    name: 'SUBSTITUTE',
    syntax: 'SUBSTITUTE(text, old_text, new_text, [instance_num])',
    params: [
      { name: 'text', desc: '文字列' },
      { name: 'old_text', desc: '置換前の文字列' },
      { name: 'new_text', desc: '置換後の文字列' },
      { name: 'instance_num', desc: '何番目を置換するか', optional: true }
    ],
    description: '文字列内の特定の文字列を置換します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  },

  REPT: {
    name: 'REPT',
    syntax: 'REPT(text, number_times)',
    params: [
      { name: 'text', desc: '繰り返すテキスト' },
      { name: 'number_times', desc: '繰り返し回数' }
    ],
    description: 'テキストを指定した回数繰り返します',
    category: 'text',
    colorScheme: COLOR_SCHEMES.text
  },

  // 04. 日付・時刻関数
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
      { name: 'return_type', desc: '戻り値の種類（1:日曜=1、2:月曜=1など）', optional: true }
    ],
    description: '日付から曜日番号を返します',
    category: 'datetime',
    colorScheme: COLOR_SCHEMES.datetime
  },

  EDATE: {
    name: 'EDATE',
    syntax: 'EDATE(start_date, months)',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'months', desc: '追加する月数（負の値も可能）' }
    ],
    description: '指定した月数後の日付を返します',
    category: 'datetime',
    colorScheme: COLOR_SCHEMES.datetime
  },

  EOMONTH: {
    name: 'EOMONTH',
    syntax: 'EOMONTH(start_date, months)',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'months', desc: '月数' }
    ],
    description: '指定した月数後の月末日を返します',
    category: 'datetime',
    colorScheme: COLOR_SCHEMES.datetime
  },

  // 05. 論理関数
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

  IFERROR: {
    name: 'IFERROR',
    syntax: 'IFERROR(value, value_if_error)',
    params: [
      { name: 'value', desc: '評価する値や数式' },
      { name: 'value_if_error', desc: 'エラー時に表示する代替値' }
    ],
    description: 'エラーの場合に指定した値を返します',
    category: 'logical',
    examples: ['=IFERROR(A1/B1,0)', '=IFERROR(VLOOKUP(A1,B:C,2,0),"見つかりません")'],
    colorScheme: COLOR_SCHEMES.logical
  },

  IFNA: {
    name: 'IFNA',
    syntax: 'IFNA(value, value_if_na)',
    params: [
      { name: 'value', desc: '評価する値や数式' },
      { name: 'value_if_na', desc: '#N/Aエラー時に表示する代替値' }
    ],
    description: '#N/Aエラーの場合に指定した値を表示します',
    category: 'logical',
    colorScheme: COLOR_SCHEMES.logical
  },

  TRUE: {
    name: 'TRUE',
    syntax: 'TRUE()',
    params: [],
    description: '論理値TRUEを返します',
    category: 'logical',
    colorScheme: COLOR_SCHEMES.logical
  },

  FALSE: {
    name: 'FALSE',
    syntax: 'FALSE()',
    params: [],
    description: '論理値FALSEを返します',
    category: 'logical',
    colorScheme: COLOR_SCHEMES.logical
  },

  // 06. 検索・参照関数
  VLOOKUP: {
    name: 'VLOOKUP',
    syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    params: [
      { name: 'lookup_value', desc: '検索したい値が入っているセル' },
      { name: 'table_array', desc: '検索対象のデータ範囲' },
      { name: 'col_index_num', desc: '取得したいデータの列番号' },
      { name: 'range_lookup', desc: '完全一致は0またはFALSE（省略可能）', optional: true }
    ],
    description: '縦方向（列）でデータを検索し、指定した列から値を返します',
    category: 'lookup',
    examples: ['=VLOOKUP(A1,B:D,3,0)', '=VLOOKUP("商品A",A1:C10,2,FALSE)'],
    colorScheme: COLOR_SCHEMES.lookup
  },

  HLOOKUP: {
    name: 'HLOOKUP',
    syntax: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])',
    params: [
      { name: 'lookup_value', desc: '検索したい値が入っているセル' },
      { name: 'table_array', desc: '検索対象のデータ範囲' },
      { name: 'row_index_num', desc: '取得したいデータの行番号' },
      { name: 'range_lookup', desc: '完全一致は0またはFALSE（省略可能）', optional: true }
    ],
    description: '横方向（行）でデータを検索し、指定した行から値を返します',
    category: 'lookup',
    colorScheme: COLOR_SCHEMES.lookup
  },

  INDEX: {
    name: 'INDEX',
    syntax: 'INDEX(array, row_num, [column_num])',
    params: [
      { name: 'array', desc: '配列または範囲' },
      { name: 'row_num', desc: '行番号' },
      { name: 'column_num', desc: '列番号', optional: true }
    ],
    description: '指定した位置の値を返します',
    category: 'lookup',
    examples: ['=INDEX(A1:C10,5,2)', '=INDEX(B:B,MATCH("検索値",A:A,0))'],
    colorScheme: COLOR_SCHEMES.lookup
  },

  MATCH: {
    name: 'MATCH',
    syntax: 'MATCH(lookup_value, lookup_array, [match_type])',
    params: [
      { name: 'lookup_value', desc: '検索値' },
      { name: 'lookup_array', desc: '検索範囲' },
      { name: 'match_type', desc: '照合の型（0:完全一致）', optional: true }
    ],
    description: '検索値の位置を返します',
    category: 'lookup',
    colorScheme: COLOR_SCHEMES.lookup
  },

  LOOKUP: {
    name: 'LOOKUP',
    syntax: 'LOOKUP(lookup_value, lookup_vector, [result_vector])',
    params: [
      { name: 'lookup_value', desc: '検索値' },
      { name: 'lookup_vector', desc: '検索ベクトル' },
      { name: 'result_vector', desc: '結果ベクトル', optional: true }
    ],
    description: 'ベクトルまたは配列を検索します',
    category: 'lookup',
    colorScheme: COLOR_SCHEMES.lookup
  },

  XLOOKUP: {
    name: 'XLOOKUP',
    syntax: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])',
    params: [
      { name: 'lookup_value', desc: '検索値' },
      { name: 'lookup_array', desc: '検索範囲' },
      { name: 'return_array', desc: '戻り値の範囲' },
      { name: 'if_not_found', desc: '見つからない場合の値', optional: true },
      { name: 'match_mode', desc: '一致モード', optional: true },
      { name: 'search_mode', desc: '検索モード', optional: true }
    ],
    description: '高度な検索機能（Excel 2021以降）',
    category: 'lookup',
    colorScheme: COLOR_SCHEMES.lookup
  },

  FILTER: {
    name: 'FILTER',
    syntax: 'FILTER(array, include, [if_empty])',
    params: [
      { name: 'array', desc: 'フィルタリングする配列' },
      { name: 'include', desc: '条件（TRUE/FALSEの配列）' },
      { name: 'if_empty', desc: '空の場合の値', optional: true }
    ],
    description: '条件に基づいてデータをフィルタリングします',
    category: 'lookup',
    colorScheme: COLOR_SCHEMES.lookup
  },

  UNIQUE: {
    name: 'UNIQUE',
    syntax: 'UNIQUE(array, [by_col], [exactly_once])',
    params: [
      { name: 'array', desc: '一意の値を抽出する配列' },
      { name: 'by_col', desc: '列で比較するか', optional: true },
      { name: 'exactly_once', desc: '1回のみ出現する値のみ', optional: true }
    ],
    description: '一意の値を抽出します',
    category: 'lookup',
    colorScheme: COLOR_SCHEMES.lookup
  },

  SORT: {
    name: 'SORT',
    syntax: 'SORT(array, [sort_index], [sort_order], [by_col])',
    params: [
      { name: 'array', desc: '並べ替える配列' },
      { name: 'sort_index', desc: '並べ替えインデックス', optional: true },
      { name: 'sort_order', desc: '並べ替え順序（1:昇順、-1:降順）', optional: true },
      { name: 'by_col', desc: '列で並べ替えるか', optional: true }
    ],
    description: '配列を並べ替えます',
    category: 'lookup',
    colorScheme: COLOR_SCHEMES.lookup
  },

  // 07. 情報関数
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
  },

  ISTEXT: {
    name: 'ISTEXT',
    syntax: 'ISTEXT(value)',
    params: [
      { name: 'value', desc: 'テストする値' }
    ],
    description: '値がテキストかどうかを判定します',
    category: 'information',
    colorScheme: COLOR_SCHEMES.information
  },

  TYPE: {
    name: 'TYPE',
    syntax: 'TYPE(value)',
    params: [
      { name: 'value', desc: 'データ型を調べる値' }
    ],
    description: '値のデータ型を返します',
    category: 'information',
    colorScheme: COLOR_SCHEMES.information
  },

  // 08. データベース関数
  DSUM: {
    name: 'DSUM',
    syntax: 'DSUM(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲' },
      { name: 'field', desc: 'フィールド（列）' },
      { name: 'criteria', desc: '条件範囲' }
    ],
    description: '条件に一致するレコードの合計を計算します',
    category: 'database',
    colorScheme: COLOR_SCHEMES.database
  },

  DAVERAGE: {
    name: 'DAVERAGE',
    syntax: 'DAVERAGE(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲' },
      { name: 'field', desc: 'フィールド（列）' },
      { name: 'criteria', desc: '条件範囲' }
    ],
    description: '条件に一致するレコードの平均を計算します',
    category: 'database',
    colorScheme: COLOR_SCHEMES.database
  },

  DCOUNT: {
    name: 'DCOUNT',
    syntax: 'DCOUNT(database, field, criteria)',
    params: [
      { name: 'database', desc: 'データベース範囲' },
      { name: 'field', desc: 'フィールド（列）' },
      { name: 'criteria', desc: '条件範囲' }
    ],
    description: '条件に一致する数値セルの個数をカウントします',
    category: 'database',
    colorScheme: COLOR_SCHEMES.database
  },

  // 09. 財務関数
  PMT: {
    name: 'PMT',
    syntax: 'PMT(rate, nper, pv, [fv], [type])',
    params: [
      { name: 'rate', desc: '期間あたりの利率' },
      { name: 'nper', desc: '支払い回数の合計' },
      { name: 'pv', desc: '現在価値（ローン額）' },
      { name: 'fv', desc: '将来価値', optional: true },
      { name: 'type', desc: '支払期日（0:期末、1:期首）', optional: true }
    ],
    description: '定期的な支払額を計算します',
    category: 'financial',
    examples: ['=PMT(0.05/12,60,-10000)', '=PMT(B1/12,B2,-B3)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  PV: {
    name: 'PV',
    syntax: 'PV(rate, nper, pmt, [fv], [type])',
    params: [
      { name: 'rate', desc: '期間あたりの利率' },
      { name: 'nper', desc: '支払い回数' },
      { name: 'pmt', desc: '定期支払額' },
      { name: 'fv', desc: '将来価値', optional: true },
      { name: 'type', desc: '支払期日', optional: true }
    ],
    description: '現在価値を計算します',
    category: 'financial',
    colorScheme: COLOR_SCHEMES.financial
  },

  FV: {
    name: 'FV',
    syntax: 'FV(rate, nper, pmt, [pv], [type])',
    params: [
      { name: 'rate', desc: '期間あたりの利率' },
      { name: 'nper', desc: '支払い回数' },
      { name: 'pmt', desc: '定期支払額' },
      { name: 'pv', desc: '現在価値', optional: true },
      { name: 'type', desc: '支払期日', optional: true }
    ],
    description: '将来価値を計算します',
    category: 'financial',
    colorScheme: COLOR_SCHEMES.financial
  },

  NPER: {
    name: 'NPER',
    syntax: 'NPER(rate, pmt, pv, [fv], [type])',
    params: [
      { name: 'rate', desc: '期間あたりの利率' },
      { name: 'pmt', desc: '各期の支払額' },
      { name: 'pv', desc: '現在価値' },
      { name: 'fv', desc: '将来価値', optional: true },
      { name: 'type', desc: '支払期日', optional: true }
    ],
    description: '支払回数を計算します',
    category: 'financial',
    examples: ['=NPER(0.05/12,-200,10000)', '=NPER(B1,B2,B3)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  RATE: {
    name: 'RATE',
    syntax: 'RATE(nper, pmt, pv, [fv], [type], [guess])',
    params: [
      { name: 'nper', desc: '支払回数' },
      { name: 'pmt', desc: '定期支払額' },
      { name: 'pv', desc: '現在価値' },
      { name: 'fv', desc: '将来価値', optional: true },
      { name: 'type', desc: '支払期日', optional: true },
      { name: 'guess', desc: '予想利率', optional: true }
    ],
    description: '利率を計算します',
    category: 'financial',
    colorScheme: COLOR_SCHEMES.financial
  },

  NPV: {
    name: 'NPV',
    syntax: 'NPV(rate, value1, [value2], ...)',
    params: [
      { name: 'rate', desc: '割引率' },
      { name: 'value1', desc: '1期目のキャッシュフロー' },
      { name: 'value2', desc: '2期目のキャッシュフロー', optional: true }
    ],
    description: '正味現在価値を計算します',
    category: 'financial',
    colorScheme: COLOR_SCHEMES.financial
  },

  IRR: {
    name: 'IRR',
    syntax: 'IRR(values, [guess])',
    params: [
      { name: 'values', desc: 'キャッシュフローの配列' },
      { name: 'guess', desc: '予想収益率', optional: true }
    ],
    description: '内部収益率を計算します',
    category: 'financial',
    colorScheme: COLOR_SCHEMES.financial
  },

  // 10. エンジニアリング関数
  BIN2DEC: {
    name: 'BIN2DEC',
    syntax: 'BIN2DEC(number)',
    params: [
      { name: 'number', desc: '2進数（最大10桁）' }
    ],
    description: '2進数を10進数に変換します',
    category: 'engineering',
    colorScheme: COLOR_SCHEMES.engineering
  },

  DEC2BIN: {
    name: 'DEC2BIN',
    syntax: 'DEC2BIN(number, [places])',
    params: [
      { name: 'number', desc: '10進数' },
      { name: 'places', desc: '桁数', optional: true }
    ],
    description: '10進数を2進数に変換します',
    category: 'engineering',
    colorScheme: COLOR_SCHEMES.engineering
  },

  HEX2DEC: {
    name: 'HEX2DEC',
    syntax: 'HEX2DEC(number)',
    params: [
      { name: 'number', desc: '16進数' }
    ],
    description: '16進数を10進数に変換します',
    category: 'engineering',
    colorScheme: COLOR_SCHEMES.engineering
  },

  DEC2HEX: {
    name: 'DEC2HEX',
    syntax: 'DEC2HEX(number, [places])',
    params: [
      { name: 'number', desc: '10進数' },
      { name: 'places', desc: '桁数', optional: true }
    ],
    description: '10進数を16進数に変換します',
    category: 'engineering',
    colorScheme: COLOR_SCHEMES.engineering
  },

  // 12. Web関数
  WEBSERVICE: {
    name: 'WEBSERVICE',
    syntax: 'WEBSERVICE(url)',
    params: [
      { name: 'url', desc: 'データを取得するURL' }
    ],
    description: 'Webサービスからデータを取得します',
    category: 'web',
    colorScheme: COLOR_SCHEMES.web
  },

  FILTERXML: {
    name: 'FILTERXML',
    syntax: 'FILTERXML(xml, xpath)',
    params: [
      { name: 'xml', desc: 'XMLデータ' },
      { name: 'xpath', desc: 'XPath式' }
    ],
    description: 'XMLデータから特定の値を抽出します',
    category: 'web',
    colorScheme: COLOR_SCHEMES.web
  },

  ENCODEURL: {
    name: 'ENCODEURL',
    syntax: 'ENCODEURL(text)',
    params: [
      { name: 'text', desc: 'エンコードする文字列' }
    ],
    description: 'URLエンコードを行います',
    category: 'web',
    colorScheme: COLOR_SCHEMES.web
  },

  // 13. Google Sheets専用関数
  ARRAYFORMULA: {
    name: 'ARRAYFORMULA',
    syntax: 'ARRAYFORMULA(array_formula)',
    params: [
      { name: 'array_formula', desc: '配列として評価する数式' }
    ],
    description: '配列数式を作成します（Google Sheets専用）',
    category: 'googlesheets',
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  QUERY: {
    name: 'QUERY',
    syntax: 'QUERY(data, query, [headers])',
    params: [
      { name: 'data', desc: 'クエリ対象のデータ範囲' },
      { name: 'query', desc: 'SQL風のクエリ文字列' },
      { name: 'headers', desc: 'ヘッダー行数', optional: true }
    ],
    description: 'SQL風のクエリでデータを操作します（Google Sheets専用）',
    category: 'googlesheets',
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  IMPORTDATA: {
    name: 'IMPORTDATA',
    syntax: 'IMPORTDATA(url)',
    params: [
      { name: 'url', desc: 'CSVまたはTSVファイルのURL' }
    ],
    description: 'CSVまたはTSVファイルをインポートします（Google Sheets専用）',
    category: 'googlesheets',
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  IMPORTHTML: {
    name: 'IMPORTHTML',
    syntax: 'IMPORTHTML(url, query, index)',
    params: [
      { name: 'url', desc: 'HTMLページのURL' },
      { name: 'query', desc: '"table"または"list"' },
      { name: 'index', desc: 'インデックス番号' }
    ],
    description: 'HTMLテーブルやリストをインポートします（Google Sheets専用）',
    category: 'googlesheets',
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  IMPORTXML: {
    name: 'IMPORTXML',
    syntax: 'IMPORTXML(url, xpath_query)',
    params: [
      { name: 'url', desc: 'XMLのURL' },
      { name: 'xpath_query', desc: 'XPath式' }
    ],
    description: 'XMLデータをインポートします（Google Sheets専用）',
    category: 'googlesheets',
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  IMPORTRANGE: {
    name: 'IMPORTRANGE',
    syntax: 'IMPORTRANGE(spreadsheet_url, range_string)',
    params: [
      { name: 'spreadsheet_url', desc: 'スプレッドシートのURL' },
      { name: 'range_string', desc: '範囲文字列（例："Sheet1!A1:C10"）' }
    ],
    description: '他のスプレッドシートからデータをインポートします（Google Sheets専用）',
    category: 'googlesheets',
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  SPARKLINE: {
    name: 'SPARKLINE',
    syntax: 'SPARKLINE(data, [options])',
    params: [
      { name: 'data', desc: 'グラフ化するデータ' },
      { name: 'options', desc: 'オプション設定', optional: true }
    ],
    description: 'セル内にミニグラフを作成します（Google Sheets専用）',
    category: 'googlesheets',
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  // 14. 新しい関数・その他
  SEQUENCE: {
    name: 'SEQUENCE',
    syntax: 'SEQUENCE(rows, [columns], [start], [step])',
    params: [
      { name: 'rows', desc: '行数' },
      { name: 'columns', desc: '列数', optional: true },
      { name: 'start', desc: '開始値', optional: true },
      { name: 'step', desc: 'ステップ', optional: true }
    ],
    description: '連続した数値の配列を生成します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

  RANDARRAY: {
    name: 'RANDARRAY',
    syntax: 'RANDARRAY([rows], [columns], [min], [max], [whole_number])',
    params: [
      { name: 'rows', desc: '行数', optional: true },
      { name: 'columns', desc: '列数', optional: true },
      { name: 'min', desc: '最小値', optional: true },
      { name: 'max', desc: '最大値', optional: true },
      { name: 'whole_number', desc: '整数フラグ', optional: true }
    ],
    description: 'ランダムな数値の配列を生成します',
    category: 'math',
    colorScheme: COLOR_SCHEMES.math
  },

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

/**
 * 関数名から使用されている個別の関数を抽出
 */
export function extractFunctionsFromName(functionName: string): string[] {
  const functions: string[] = [];
  
  // &で区切られた関数名を抽出
  const parts = functionName.split('&').map(part => part.trim());
  
  for (const part of parts) {
    // 各パートから関数名を抽出（英語の関数名のみ）
    const matches = part.match(/[A-Z][A-Z0-9.]*/g);
    if (matches) {
      functions.push(...matches);
    }
  }
  
  // 重複を除去して返す
  return [...new Set(functions)];
}

/**
 * 使用関数の詳細説明を生成（見やすく整形）
 */
export function generateFunctionDetails(functions: string[]): string {
  let details = '';
  
  for (const func of functions) {
    const definition = FUNCTION_DEFINITIONS[func];
    if (!definition) continue;
    
    // 関数名とシンタックスを表示
    details += `${definition.syntax}\n`;
    
    // 各引数を改行して表示
    for (const param of definition.params) {
      const optional = param.optional ? '（省略可能）' : '';
      details += `- ${param.name}: ${param.desc}${optional}\n`;
    }
    
    // 関数間にスペースを追加
    details += '\n';
  }
  
  return details.trim();
}

/**
 * 関数タイプを取得する関数
 */
export function getFunctionType(functionName: string): string {
  const def = FUNCTION_DEFINITIONS[functionName];
  return def ? def.category : 'other';
}

/**
 * 関数の定義を取得
 */
export function getFunctionDefinition(functionName: string): FunctionDefinition | undefined {
  return FUNCTION_DEFINITIONS[functionName.toUpperCase()];
}

/**
 * カテゴリから関数リストを取得
 */
export function getFunctionsByCategory(category: string): FunctionDefinition[] {
  return Object.values(FUNCTION_DEFINITIONS).filter(def => def.category === category);
}

/**
 * すべての関数名を取得
 */
export function getAllFunctionNames(): string[] {
  return Object.keys(FUNCTION_DEFINITIONS);
}

/**
 * 数式から使用されている関数を抽出
 */
export function extractFunctionsFromFormula(formula: string): string[] {
  const functions: string[] = [];
  const functionPattern = /([A-Z_]+(?:\.[A-Z]+)?)\s*\(/g;
  let match;
  
  while ((match = functionPattern.exec(formula)) !== null) {
    const funcName = match[1];
    if (FUNCTION_DEFINITIONS[funcName]) {
      functions.push(funcName);
    }
  }
  
  return [...new Set(functions)]; // 重複を除去
}