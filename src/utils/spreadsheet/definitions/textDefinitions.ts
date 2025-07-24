// 文字列操作関数の定義
import { type FunctionDefinition } from '../types/types';
import { COLOR_SCHEMES } from '../types/colorSchemes';

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
    examples: ['=LEFT("Hello",2)', '=LEFT(A1,3)'],
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
    examples: ['=RIGHT("Hello",2)', '=RIGHT(A1,3)'],
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
    description: '文字列の途中から指定した文字数を取り出します',
    category: 'text',
    examples: ['=MID("Hello World",7,5)', '=MID(A1,3,2)'],
    colorScheme: COLOR_SCHEMES.text
  },

  LEN: {
    name: 'LEN',
    syntax: 'LEN(text)',
    params: [
      { name: 'text', desc: '文字列' }
    ],
    description: '文字列の文字数を返します',
    category: 'text',
    examples: ['=LEN("Hello")', '=LEN(A1)'],
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
    examples: ['=LOWER("HELLO")', '=LOWER(A1)'],
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
    examples: ['=UPPER("hello")', '=UPPER(A1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  PROPER: {
    name: 'PROPER',
    syntax: 'PROPER(text)',
    params: [
      { name: 'text', desc: '先頭を大文字にする文字列' }
    ],
    description: '文字列の各単語の先頭を大文字にします',
    category: 'text',
    examples: ['=PROPER("hello world")', '=PROPER(A1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  TRIM: {
    name: 'TRIM',
    syntax: 'TRIM(text)',
    params: [
      { name: 'text', desc: 'スペースを削除する文字列' }
    ],
    description: '文字列から余分なスペースを削除します',
    category: 'text',
    examples: ['=TRIM("  Hello  World  ")', '=TRIM(A1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  SUBSTITUTE: {
    name: 'SUBSTITUTE',
    syntax: 'SUBSTITUTE(text, old_text, new_text, [instance_num])',
    params: [
      { name: 'text', desc: '対象の文字列' },
      { name: 'old_text', desc: '置換する文字列' },
      { name: 'new_text', desc: '置換後の文字列' },
      { name: 'instance_num', desc: '何番目を置換するか', optional: true }
    ],
    description: '文字列の一部を置き換えます',
    category: 'text',
    examples: ['=SUBSTITUTE("Hello World","World","Excel")', '=SUBSTITUTE(A1,"old","new")'],
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
    description: '文字列の位置を検索します（大文字小文字を区別）',
    category: 'text',
    examples: ['=FIND("o","Hello")', '=FIND("World",A1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  SEARCH: {
    name: 'SEARCH',
    syntax: 'SEARCH(find_text, within_text, [start_num])',
    params: [
      { name: 'find_text', desc: '検索する文字列' },
      { name: 'within_text', desc: '検索対象の文字列' },
      { name: 'start_num', desc: '開始位置', optional: true }
    ],
    description: '文字列の位置を検索します（大文字小文字を区別しない）',
    category: 'text',
    examples: ['=SEARCH("o","Hello")', '=SEARCH("world",A1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  REPLACE: {
    name: 'REPLACE',
    syntax: 'REPLACE(old_text, start_num, num_chars, new_text)',
    params: [
      { name: 'old_text', desc: '対象の文字列' },
      { name: 'start_num', desc: '開始位置' },
      { name: 'num_chars', desc: '置換する文字数' },
      { name: 'new_text', desc: '新しい文字列' }
    ],
    description: '文字列の指定した位置の文字を置き換えます',
    category: 'text',
    examples: ['=REPLACE("Hello",2,3,"i")', '=REPLACE(A1,1,5,"Hi")'],
    colorScheme: COLOR_SCHEMES.text
  },

  REPT: {
    name: 'REPT',
    syntax: 'REPT(text, number_times)',
    params: [
      { name: 'text', desc: '繰り返す文字列' },
      { name: 'number_times', desc: '繰り返す回数' }
    ],
    description: '文字列を指定回数繰り返します',
    category: 'text',
    examples: ['=REPT("*",5)', '=REPT(A1,3)'],
    colorScheme: COLOR_SCHEMES.text
  },

  TEXT: {
    name: 'TEXT',
    syntax: 'TEXT(value, format_text)',
    params: [
      { name: 'value', desc: '書式設定する値' },
      { name: 'format_text', desc: '表示形式' }
    ],
    description: '数値を指定した表示形式の文字列に変換します',
    category: 'text',
    examples: ['=TEXT(1234.5,"#,##0.00")', '=TEXT(TODAY(),"yyyy/mm/dd")'],
    colorScheme: COLOR_SCHEMES.text
  },

  VALUE: {
    name: 'VALUE',
    syntax: 'VALUE(text)',
    params: [
      { name: 'text', desc: '数値に変換する文字列' }
    ],
    description: '文字列を数値に変換します',
    category: 'text',
    examples: ['=VALUE("123")', '=VALUE(A1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  CHAR: {
    name: 'CHAR',
    syntax: 'CHAR(number)',
    params: [
      { name: 'number', desc: '文字コード（1～255）' }
    ],
    description: '文字コードに対応する文字を返します',
    category: 'text',
    examples: ['=CHAR(65)', '=CHAR(13)'],
    colorScheme: COLOR_SCHEMES.text
  },

  CODE: {
    name: 'CODE',
    syntax: 'CODE(text)',
    params: [
      { name: 'text', desc: '文字列' }
    ],
    description: '文字列の先頭文字の文字コードを返します',
    category: 'text',
    examples: ['=CODE("A")', '=CODE(A1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  EXACT: {
    name: 'EXACT',
    syntax: 'EXACT(text1, text2)',
    params: [
      { name: 'text1', desc: '1つ目の文字列' },
      { name: 'text2', desc: '2つ目の文字列' }
    ],
    description: '2つの文字列が同じかどうかを比較します',
    category: 'text',
    examples: ['=EXACT("Hello","Hello")', '=EXACT(A1,B1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  CLEAN: {
    name: 'CLEAN',
    syntax: 'CLEAN(text)',
    params: [
      { name: 'text', desc: '印刷できない文字を削除する文字列' }
    ],
    description: '文字列から印刷できない文字を削除します',
    category: 'text',
    examples: ['=CLEAN(A1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  CONCAT: {
    name: 'CONCAT',
    syntax: 'CONCAT(text1, [text2], ...)',
    params: [
      { name: 'text1', desc: '結合する最初のテキスト' },
      { name: 'text2', desc: '結合する追加のテキスト', optional: true }
    ],
    description: '複数のテキストを結合します（新バージョン）',
    category: 'text',
    examples: ['=CONCAT("Hello"," ","World")', '=CONCAT(A1:C1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  TEXTJOIN: {
    name: 'TEXTJOIN',
    syntax: 'TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)',
    params: [
      { name: 'delimiter', desc: '区切り文字' },
      { name: 'ignore_empty', desc: '空白セルを無視するか' },
      { name: 'text1', desc: '結合する最初のテキスト' },
      { name: 'text2', desc: '結合する追加のテキスト', optional: true }
    ],
    description: '区切り文字を使って複数のテキストを結合します',
    category: 'text',
    examples: ['=TEXTJOIN(",",TRUE,A1:A5)', '=TEXTJOIN(" ",FALSE,"Hello","World")'],
    colorScheme: COLOR_SCHEMES.text
  },

  SPLIT: {
    name: 'SPLIT',
    syntax: 'SPLIT(text, delimiter, [split_by_each], [remove_empty_text])',
    params: [
      { name: 'text', desc: '分割する文字列' },
      { name: 'delimiter', desc: '区切り文字' },
      { name: 'split_by_each', desc: '各文字で分割するか', optional: true },
      { name: 'remove_empty_text', desc: '空の文字列を削除するか', optional: true }
    ],
    description: '文字列を区切り文字で分割します',
    category: 'text',
    examples: ['=SPLIT("apple,banana,orange",",")', '=SPLIT(A1," ")'],
    colorScheme: COLOR_SCHEMES.text
  },

  T: {
    name: 'T',
    syntax: 'T(value)',
    params: [
      { name: 'value', desc: '評価する値' }
    ],
    description: '値がテキストの場合はそのテキストを、それ以外は空文字列を返します',
    category: 'text',
    examples: ['=T("Hello")', '=T(123)'],
    colorScheme: COLOR_SCHEMES.text
  },

  FIXED: {
    name: 'FIXED',
    syntax: 'FIXED(number, [decimals], [no_commas])',
    params: [
      { name: 'number', desc: '書式設定する数値' },
      { name: 'decimals', desc: '小数点以下の桁数', optional: true },
      { name: 'no_commas', desc: '桁区切りを使用しないか', optional: true }
    ],
    description: '数値を固定小数点形式の文字列に変換します',
    category: 'text',
    examples: ['=FIXED(1234.567,2)', '=FIXED(1234.567,2,TRUE)'],
    colorScheme: COLOR_SCHEMES.text
  },

  NUMBERVALUE: {
    name: 'NUMBERVALUE',
    syntax: 'NUMBERVALUE(text, [decimal_separator], [group_separator])',
    params: [
      { name: 'text', desc: '数値に変換する文字列' },
      { name: 'decimal_separator', desc: '小数点記号', optional: true },
      { name: 'group_separator', desc: '桁区切り記号', optional: true }
    ],
    description: 'ロケール依存の書式を持つ文字列を数値に変換します',
    category: 'text',
    examples: ['=NUMBERVALUE("1.234,56",",",".")'],
    colorScheme: COLOR_SCHEMES.text
  },

  DOLLAR: {
    name: 'DOLLAR',
    syntax: 'DOLLAR(number, [decimals])',
    params: [
      { name: 'number', desc: '書式設定する数値' },
      { name: 'decimals', desc: '小数点以下の桁数', optional: true }
    ],
    description: '数値を通貨形式の文字列に変換します',
    category: 'text',
    examples: ['=DOLLAR(1234.567)', '=DOLLAR(1234.567,2)'],
    colorScheme: COLOR_SCHEMES.text
  },

  UNICHAR: {
    name: 'UNICHAR',
    syntax: 'UNICHAR(number)',
    params: [
      { name: 'number', desc: 'Unicodeコードポイント' }
    ],
    description: 'Unicodeコードポイントに対応する文字を返します',
    category: 'text',
    examples: ['=UNICHAR(65)', '=UNICHAR(12354)'],
    colorScheme: COLOR_SCHEMES.text
  },

  UNICODE: {
    name: 'UNICODE',
    syntax: 'UNICODE(text)',
    params: [
      { name: 'text', desc: '文字列' }
    ],
    description: '文字列の先頭文字のUnicodeコードポイントを返します',
    category: 'text',
    examples: ['=UNICODE("A")', '=UNICODE("あ")'],
    colorScheme: COLOR_SCHEMES.text
  },

  LENB: {
    name: 'LENB',
    syntax: 'LENB(text)',
    params: [
      { name: 'text', desc: 'バイト数を求める文字列' }
    ],
    description: '文字列のバイト数を返します',
    category: 'text',
    examples: ['=LENB("Hello")', '=LENB("こんにちは")'],
    colorScheme: COLOR_SCHEMES.text
  },

  FINDB: {
    name: 'FINDB',
    syntax: 'FINDB(find_text, within_text, [start_num])',
    params: [
      { name: 'find_text', desc: '検索する文字列' },
      { name: 'within_text', desc: '検索対象の文字列' },
      { name: 'start_num', desc: '開始位置（バイト）', optional: true }
    ],
    description: '文字列の位置をバイト単位で検索します',
    category: 'text',
    examples: ['=FINDB("o","Hello")', '=FINDB("ん","こんにちは")'],
    colorScheme: COLOR_SCHEMES.text
  },

  SEARCHB: {
    name: 'SEARCHB',
    syntax: 'SEARCHB(find_text, within_text, [start_num])',
    params: [
      { name: 'find_text', desc: '検索する文字列' },
      { name: 'within_text', desc: '検索対象の文字列' },
      { name: 'start_num', desc: '開始位置（バイト）', optional: true }
    ],
    description: '文字列の位置をバイト単位で検索します（大文字小文字を区別しない）',
    category: 'text',
    examples: ['=SEARCHB("O","Hello")', '=SEARCHB("ん","こんにちは")'],
    colorScheme: COLOR_SCHEMES.text
  },

  REPLACEB: {
    name: 'REPLACEB',
    syntax: 'REPLACEB(old_text, start_num, num_bytes, new_text)',
    params: [
      { name: 'old_text', desc: '対象の文字列' },
      { name: 'start_num', desc: '開始位置（バイト）' },
      { name: 'num_bytes', desc: '置換するバイト数' },
      { name: 'new_text', desc: '新しい文字列' }
    ],
    description: '文字列の指定したバイト位置の文字を置き換えます',
    category: 'text',
    examples: ['=REPLACEB("Hello",2,2,"i")'],
    colorScheme: COLOR_SCHEMES.text
  },

  TEXTBEFORE: {
    name: 'TEXTBEFORE',
    syntax: 'TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])',
    params: [
      { name: 'text', desc: '対象の文字列' },
      { name: 'delimiter', desc: '区切り文字' },
      { name: 'instance_num', desc: '何番目の区切り文字か', optional: true },
      { name: 'match_mode', desc: '大文字小文字の区別', optional: true },
      { name: 'match_end', desc: '文字列の末尾を区切り文字として扱うか', optional: true },
      { name: 'if_not_found', desc: '見つからない場合の値', optional: true }
    ],
    description: '指定した区切り文字より前のテキストを返します',
    category: 'text',
    examples: ['=TEXTBEFORE("apple-banana-orange","-")', '=TEXTBEFORE(A1,",")'],
    colorScheme: COLOR_SCHEMES.text
  },

  TEXTAFTER: {
    name: 'TEXTAFTER',
    syntax: 'TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])',
    params: [
      { name: 'text', desc: '対象の文字列' },
      { name: 'delimiter', desc: '区切り文字' },
      { name: 'instance_num', desc: '何番目の区切り文字か', optional: true },
      { name: 'match_mode', desc: '大文字小文字の区別', optional: true },
      { name: 'match_end', desc: '文字列の先頭を区切り文字として扱うか', optional: true },
      { name: 'if_not_found', desc: '見つからない場合の値', optional: true }
    ],
    description: '指定した区切り文字より後のテキストを返します',
    category: 'text',
    examples: ['=TEXTAFTER("apple-banana-orange","-")', '=TEXTAFTER(A1,",")'],
    colorScheme: COLOR_SCHEMES.text
  },

  TEXTSPLIT: {
    name: 'TEXTSPLIT',
    syntax: 'TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])',
    params: [
      { name: 'text', desc: '分割する文字列' },
      { name: 'col_delimiter', desc: '列の区切り文字' },
      { name: 'row_delimiter', desc: '行の区切り文字', optional: true },
      { name: 'ignore_empty', desc: '空のセルを無視するか', optional: true },
      { name: 'match_mode', desc: '大文字小文字の区別', optional: true },
      { name: 'pad_with', desc: '不足分を埋める値', optional: true }
    ],
    description: 'テキストを行と列に分割します',
    category: 'text',
    examples: ['=TEXTSPLIT("apple,banana;orange,grape",",",";")', '=TEXTSPLIT(A1," ")'],
    colorScheme: COLOR_SCHEMES.text
  },

  ASC: {
    name: 'ASC',
    syntax: 'ASC(text)',
    params: [
      { name: 'text', desc: '変換する文字列' }
    ],
    description: '全角文字を半角文字に変換します',
    category: 'text',
    examples: ['=ASC("ＡＢＣ１２３")', '=ASC("ハローワールド")'],
    colorScheme: COLOR_SCHEMES.text
  },

  JIS: {
    name: 'JIS',
    syntax: 'JIS(text)',
    params: [
      { name: 'text', desc: '変換する文字列' }
    ],
    description: '半角文字を全角文字に変換します',
    category: 'text',
    examples: ['=JIS("ABC123")', '=JIS("ﾊﾛｰﾜｰﾙﾄﾞ")'],
    colorScheme: COLOR_SCHEMES.text
  },

  DBCS: {
    name: 'DBCS',
    syntax: 'DBCS(text)',
    params: [
      { name: 'text', desc: '変換する文字列' }
    ],
    description: '半角文字を全角文字に変換します（JISと同じ）',
    category: 'text',
    examples: ['=DBCS("ABC123")', '=DBCS("ﾊﾛｰﾜｰﾙﾄﾞ")'],
    colorScheme: COLOR_SCHEMES.text
  },

  PHONETIC: {
    name: 'PHONETIC',
    syntax: 'PHONETIC(reference)',
    params: [
      { name: 'reference', desc: 'ふりがなを抽出するセル参照' }
    ],
    description: 'セルからふりがな（ルビ）を抽出します',
    category: 'text',
    examples: ['=PHONETIC(A1)'],
    colorScheme: COLOR_SCHEMES.text
  },

  BAHTTEXT: {
    name: 'BAHTTEXT',
    syntax: 'BAHTTEXT(number)',
    params: [
      { name: 'number', desc: 'タイ語テキストに変換する数値' }
    ],
    description: '数値をタイ語のテキストに変換します',
    category: 'text',
    examples: ['=BAHTTEXT(1234)'],
    colorScheme: COLOR_SCHEMES.text
  }
};