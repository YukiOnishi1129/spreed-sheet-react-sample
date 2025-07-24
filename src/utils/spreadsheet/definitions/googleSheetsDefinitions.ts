// Google Sheets専用関数の定義
import { type FunctionDefinition } from '../types/types';
import { COLOR_SCHEMES } from '../types/colorSchemes';

export const GOOGLESHEETS_FUNCTIONS: Record<string, FunctionDefinition> = {
  // 配列・フィルタ関数
  ARRAYFORMULA: {
    name: 'ARRAYFORMULA',
    syntax: 'ARRAYFORMULA(array_formula)',
    params: [
      { name: 'array_formula', desc: '配列数式' }
    ],
    description: '配列数式を範囲全体に適用します',
    category: 'googlesheets',
    examples: ['=ARRAYFORMULA(A1:A10*B1:B10)', '=ARRAYFORMULA(IF(A1:A10>0,A1:A10,""))'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  QUERY: {
    name: 'QUERY',
    syntax: 'QUERY(data, query, [headers])',
    params: [
      { name: 'data', desc: 'データ範囲' },
      { name: 'query', desc: 'SQLライクなクエリ文字列' },
      { name: 'headers', desc: 'ヘッダー行数', optional: true }
    ],
    description: 'SQLライクなクエリでデータを操作します',
    category: 'googlesheets',
    examples: ['=QUERY(A1:D10,"SELECT A, B WHERE C > 100")', '=QUERY(A:D,"SELECT * WHERE A is not null ORDER BY B DESC")'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  SORTN: {
    name: 'SORTN',
    syntax: 'SORTN(range, [n], [display_ties_mode], [sort_column], [is_ascending])',
    params: [
      { name: 'range', desc: '並べ替える範囲' },
      { name: 'n', desc: '上位何件を表示するか', optional: true },
      { name: 'display_ties_mode', desc: '同値の表示方法', optional: true },
      { name: 'sort_column', desc: '並べ替え基準列', optional: true },
      { name: 'is_ascending', desc: '昇順かどうか', optional: true }
    ],
    description: '上位N件を並べ替えて返します',
    category: 'googlesheets',
    examples: ['=SORTN(A1:B10,5)', '=SORTN(A:C,3,0,2,FALSE)'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  FLATTEN: {
    name: 'FLATTEN',
    syntax: 'FLATTEN(range1, [range2], ...)',
    params: [
      { name: 'range1', desc: '1つ目の範囲' },
      { name: 'range2', desc: '2つ目の範囲', optional: true }
    ],
    description: '複数の範囲を1次元配列に変換します',
    category: 'googlesheets',
    examples: ['=FLATTEN(A1:B5)', '=FLATTEN(A1:A5,C1:C3,E1:E7)'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  // Google固有の関数
  GOOGLEFINANCE: {
    name: 'GOOGLEFINANCE',
    syntax: 'GOOGLEFINANCE(ticker, [attribute], [start_date], [end_date], [interval])',
    params: [
      { name: 'ticker', desc: '銘柄コード' },
      { name: 'attribute', desc: '取得する情報', optional: true },
      { name: 'start_date', desc: '開始日', optional: true },
      { name: 'end_date', desc: '終了日', optional: true },
      { name: 'interval', desc: '間隔', optional: true }
    ],
    description: 'Google Financeから株価情報を取得します',
    category: 'googlesheets',
    examples: ['=GOOGLEFINANCE("GOOGL")', '=GOOGLEFINANCE("AAPL","price")', '=GOOGLEFINANCE("MSFT","high",TODAY()-30,TODAY(),"DAILY")'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  GOOGLETRANSLATE: {
    name: 'GOOGLETRANSLATE',
    syntax: 'GOOGLETRANSLATE(text, source_language, target_language)',
    params: [
      { name: 'text', desc: '翻訳するテキスト' },
      { name: 'source_language', desc: '元の言語コード' },
      { name: 'target_language', desc: '翻訳先の言語コード' }
    ],
    description: 'Google翻訳でテキストを翻訳します',
    category: 'googlesheets',
    examples: ['=GOOGLETRANSLATE("Hello","en","ja")', '=GOOGLETRANSLATE("こんにちは","ja","en")'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  DETECTLANGUAGE: {
    name: 'DETECTLANGUAGE',
    syntax: 'DETECTLANGUAGE(text)',
    params: [
      { name: 'text', desc: '言語を検出するテキスト' }
    ],
    description: 'テキストの言語を検出します',
    category: 'googlesheets',
    examples: ['=DETECTLANGUAGE("Hello World")', '=DETECTLANGUAGE("こんにちは")'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  SPARKLINE: {
    name: 'SPARKLINE',
    syntax: 'SPARKLINE(data, [options])',
    params: [
      { name: 'data', desc: 'データ範囲' },
      { name: 'options', desc: 'オプション設定', optional: true }
    ],
    description: 'スパークライン（小さなグラフ）を作成します',
    category: 'googlesheets',
    examples: ['=SPARKLINE(A1:A10)', '=SPARKLINE(B1:B20,{"charttype","column";"color1","red"})'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  // 文字列処理関数（Google Sheets拡張）
  SPLIT: {
    name: 'SPLIT',
    syntax: 'SPLIT(text, delimiter, [split_by_each], [remove_empty_text])',
    params: [
      { name: 'text', desc: '分割するテキスト' },
      { name: 'delimiter', desc: '区切り文字' },
      { name: 'split_by_each', desc: '各文字で分割するか', optional: true },
      { name: 'remove_empty_text', desc: '空のテキストを除去するか', optional: true }
    ],
    description: '文字列を区切り文字で分割します',
    category: 'googlesheets',
    examples: ['=SPLIT("apple,banana,cherry",",")', '=SPLIT("a-b-c","-")'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  JOIN: {
    name: 'JOIN',
    syntax: 'JOIN(delimiter, value_or_array1, [value_or_array2], ...)',
    params: [
      { name: 'delimiter', desc: '区切り文字' },
      { name: 'value_or_array1', desc: '1つ目の値または配列' },
      { name: 'value_or_array2', desc: '2つ目の値または配列', optional: true }
    ],
    description: '値を指定した区切り文字で結合します',
    category: 'googlesheets',
    examples: ['=JOIN(",",A1:A5)', '=JOIN(" - ","apple","banana","cherry")'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  REGEXEXTRACT: {
    name: 'REGEXEXTRACT',
    syntax: 'REGEXEXTRACT(text, regular_expression)',
    params: [
      { name: 'text', desc: '対象テキスト' },
      { name: 'regular_expression', desc: '正規表現' }
    ],
    description: '正規表現にマッチする最初の部分文字列を抽出します',
    category: 'googlesheets',
    examples: ['=REGEXEXTRACT("abc123def","\\d+")', '=REGEXEXTRACT("Email: john@example.com","[\\w\\.-]+@[\\w\\.-]+")'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  REGEXMATCH: {
    name: 'REGEXMATCH',
    syntax: 'REGEXMATCH(text, regular_expression)',
    params: [
      { name: 'text', desc: '対象テキスト' },
      { name: 'regular_expression', desc: '正規表現' }
    ],
    description: 'テキストが正規表現にマッチするかを判定します',
    category: 'googlesheets',
    examples: ['=REGEXMATCH("123-456-7890","\\d{3}-\\d{3}-\\d{4}")', '=REGEXMATCH("test@email.com",".*@.*\\.com")'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  REGEXREPLACE: {
    name: 'REGEXREPLACE',
    syntax: 'REGEXREPLACE(text, regular_expression, replacement)',
    params: [
      { name: 'text', desc: '対象テキスト' },
      { name: 'regular_expression', desc: '正規表現' },
      { name: 'replacement', desc: '置換文字列' }
    ],
    description: '正規表現にマッチする部分を置換します',
    category: 'googlesheets',
    examples: ['=REGEXREPLACE("Hello World","o","0")', '=REGEXREPLACE("abc123def","\\d+","XXX")'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  // データ型変換関数
  TO_DATE: {
    name: 'TO_DATE',
    syntax: 'TO_DATE(value)',
    params: [
      { name: 'value', desc: '日付に変換する値' }
    ],
    description: '値を日付に変換します',
    category: 'googlesheets',
    examples: ['=TO_DATE(25569)', '=TO_DATE("1/1/2020")'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  TO_DOLLARS: {
    name: 'TO_DOLLARS',
    syntax: 'TO_DOLLARS(value)',
    params: [
      { name: 'value', desc: 'ドル表記に変換する値' }
    ],
    description: '値をドル表記に変換します',
    category: 'googlesheets',
    examples: ['=TO_DOLLARS(1234.56)', '=TO_DOLLARS(A1)'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  TO_PERCENT: {
    name: 'TO_PERCENT',
    syntax: 'TO_PERCENT(value)',
    params: [
      { name: 'value', desc: 'パーセント表記に変換する値' }
    ],
    description: '値をパーセント表記に変換します',
    category: 'googlesheets',
    examples: ['=TO_PERCENT(0.123)', '=TO_PERCENT(0.75)'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  TO_TEXT: {
    name: 'TO_TEXT',
    syntax: 'TO_TEXT(value)',
    params: [
      { name: 'value', desc: 'テキストに変換する値' }
    ],
    description: '値をテキストに変換します',
    category: 'googlesheets',
    examples: ['=TO_TEXT(123)', '=TO_TEXT(TRUE)'],
    colorScheme: COLOR_SCHEMES.googlesheets
  },

  ISBETWEEN: {
    name: 'ISBETWEEN',
    syntax: 'ISBETWEEN(value_to_compare, lower_value, upper_value, [lower_value_inclusive], [upper_value_inclusive])',
    params: [
      { name: 'value_to_compare', desc: '比較する値' },
      { name: 'lower_value', desc: '下限値' },
      { name: 'upper_value', desc: '上限値' },
      { name: 'lower_value_inclusive', desc: '下限値を含むか', optional: true },
      { name: 'upper_value_inclusive', desc: '上限値を含むか', optional: true }
    ],
    description: '値が指定した範囲内にあるかを判定します',
    category: 'googlesheets',
    examples: ['=ISBETWEEN(5,1,10)', '=ISBETWEEN(A1,0,100,TRUE,FALSE)'],
    colorScheme: COLOR_SCHEMES.googlesheets
  }
};