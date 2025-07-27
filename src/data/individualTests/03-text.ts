import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const textTests: IndividualFunctionTest[] = [
  {
    name: 'CONCATENATE',
    category: '03. 文字列',
    description: '文字列を結合',
    data: [
      ['文字列1', '文字列2', '文字列3', '結果'],
      ['Hello', ' ', 'World', '=CONCATENATE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 'Hello World' }
  },
  {
    name: 'CONCAT',
    category: '03. 文字列',
    description: '文字列を結合（新版）',
    data: [
      ['文字列1', '文字列2', '結果'],
      ['ABC', 'DEF', '=CONCAT(A2,B2)']
    ],
    expectedValues: { 'C2': 'ABCDEF' }
  },
  {
    name: 'TEXTJOIN',
    category: '03. 文字列',
    description: '区切り文字で結合',
    data: [
      ['値1', '値2', '値3', '結果'],
      ['A', 'B', 'C', '=TEXTJOIN("-",TRUE,A2:C2)']
    ],
    expectedValues: { 'D2': 'A-B-C' }
  },
  {
    name: 'LEFT',
    category: '03. 文字列',
    description: '左から文字を抽出',
    data: [
      ['文字列', '文字数', '結果'],
      ['Hello World', 5, '=LEFT(A2,B2)']
    ],
    expectedValues: { 'C2': 'Hello' }
  },
  {
    name: 'RIGHT',
    category: '03. 文字列',
    description: '右から文字を抽出',
    data: [
      ['文字列', '文字数', '結果'],
      ['Hello World', 5, '=RIGHT(A2,B2)']
    ],
    expectedValues: { 'C2': 'World' }
  },
  {
    name: 'MID',
    category: '03. 文字列',
    description: '中間の文字を抽出',
    data: [
      ['文字列', '開始位置', '文字数', '結果'],
      ['Hello World', 3, 3, '=MID(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 'llo' }
  },
  {
    name: 'LEN',
    category: '03. 文字列',
    description: '文字数を返す',
    data: [
      ['文字列', '文字数'],
      ['Hello', '=LEN(A2)'],
      ['こんにちは', '=LEN(A3)']
    ],
    expectedValues: { 'B2': 5, 'B3': 5 }
  },
  {
    name: 'FIND',
    category: '03. 文字列',
    description: '文字位置を検索（大小区別）',
    data: [
      ['検索文字列', '対象文字列', '位置'],
      ['o', 'Hello World', '=FIND(A2,B2)']
    ],
    expectedValues: { 'C2': 5 }
  },
  {
    name: 'SEARCH',
    category: '03. 文字列',
    description: '文字位置を検索（大小区別なし）',
    data: [
      ['検索文字列', '対象文字列', '位置'],
      ['world', 'Hello World', '=SEARCH(A2,B2)']
    ],
    expectedValues: { 'C2': 7 }
  },
  {
    name: 'REPLACE',
    category: '03. 文字列',
    description: '文字を置換（位置指定）',
    data: [
      ['文字列', '開始位置', '文字数', '新文字列', '結果'],
      ['Hello World', 7, 5, 'Excel', '=REPLACE(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 'Hello Excel' }
  },
  {
    name: 'SUBSTITUTE',
    category: '03. 文字列',
    description: '文字を置換（文字指定）',
    data: [
      ['文字列', '検索文字列', '置換文字列', '結果'],
      ['Hello World', 'World', 'Excel', '=SUBSTITUTE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 'Hello Excel' }
  },
  {
    name: 'UPPER',
    category: '03. 文字列',
    description: '大文字に変換',
    data: [
      ['文字列', '結果'],
      ['Hello World', '=UPPER(A2)']
    ],
    expectedValues: { 'B2': 'HELLO WORLD' }
  },
  {
    name: 'LOWER',
    category: '03. 文字列',
    description: '小文字に変換',
    data: [
      ['文字列', '結果'],
      ['Hello World', '=LOWER(A2)']
    ],
    expectedValues: { 'B2': 'hello world' }
  },
  {
    name: 'PROPER',
    category: '03. 文字列',
    description: '先頭を大文字に変換',
    data: [
      ['文字列', '結果'],
      ['hello world', '=PROPER(A2)']
    ],
    expectedValues: { 'B2': 'Hello World' }
  },
  {
    name: 'TRIM',
    category: '03. 文字列',
    description: '余分なスペースを削除',
    data: [
      ['文字列', '結果'],
      ['  Hello   World  ', '=TRIM(A2)']
    ],
    expectedValues: { 'B2': 'Hello World' }
  },
  {
    name: 'TEXT',
    category: '03. 文字列',
    description: '数値を書式付き文字列に変換',
    data: [
      ['値', '表示形式', '結果'],
      [1234.56, '#,##0.00', '=TEXT(A2,B2)']
    ],
    expectedValues: { 'C2': '1,234.56' }
  },
  {
    name: 'VALUE',
    category: '03. 文字列',
    description: '文字列を数値に変換',
    data: [
      ['文字列', '結果'],
      ['123.45', '=VALUE(A2)']
    ],
    expectedValues: { 'B2': 123.45 }
  },
  {
    name: 'REPT',
    category: '03. 文字列',
    description: '文字列を繰り返す',
    data: [
      ['文字列', '繰り返し回数', '結果'],
      ['★', 5, '=REPT(A2,B2)']
    ],
    expectedValues: { 'C2': '★★★★★' }
  },
  {
    name: 'CHAR',
    category: '03. 文字列',
    description: '文字コードから文字を返す',
    data: [
      ['文字コード', '文字'],
      [65, '=CHAR(A2)'],
      [97, '=CHAR(A3)']
    ],
    expectedValues: { 'B2': 'A', 'B3': 'a' }
  },
  {
    name: 'CODE',
    category: '03. 文字列',
    description: '文字から文字コードを返す',
    data: [
      ['文字', '文字コード'],
      ['A', '=CODE(A2)'],
      ['a', '=CODE(A3)']
    ],
    expectedValues: { 'B2': 65, 'B3': 97 }
  },
  {
    name: 'EXACT',
    category: '03. 文字列',
    description: '文字列が同一か判定',
    data: [
      ['文字列1', '文字列2', '結果'],
      ['Hello', 'Hello', '=EXACT(A2,B2)'],
      ['Hello', 'hello', '=EXACT(A3,B3)']
    ],
    expectedValues: { 'C2': true, 'C3': false }
  },
  {
    name: 'CLEAN',
    category: '03. 文字列',
    description: '印刷できない文字を削除',
    data: [
      ['テキスト', 'クリーン後'],
      ['Hello\x00World', '=CLEAN(A2)'],
      ['Test\x07String', '=CLEAN(A3)']
    ],
    expectedValues: { 'B2': 'HelloWorld', 'B3': 'TestString' }
  },
  {
    name: 'DOLLAR',
    category: '03. 文字列',
    description: '通貨形式に変換',
    data: [
      ['数値', '小数桁', '結果'],
      [1234.567, 2, '=DOLLAR(A2,B2)'],
      [1234.567, 0, '=DOLLAR(A3,B3)']
    ],
    expectedValues: { 'C2': '$1,234.57', 'C3': '$1,235' }
  },
  {
    name: 'FIXED',
    category: '03. 文字列',
    description: '小数点以下の桁数を固定',
    data: [
      ['数値', '小数桁', 'カンマなし', '結果'],
      [1234.567, 2, 'FALSE', '=FIXED(A2,B2,C2)'],
      [1234.567, 1, 'TRUE', '=FIXED(A3,B3,C3)']
    ],
    expectedValues: { 'D2': '1,234.57', 'D3': '1234.6' }
  },
  {
    name: 'NUMBERVALUE',
    category: '03. 文字列',
    description: 'テキストを数値に変換',
    data: [
      ['テキスト', '小数点', '桁区切り', '結果'],
      ['1,234.56', '.', ',', '=NUMBERVALUE(A2,B2,C2)'],
      ['1.234,56', ',', '.', '=NUMBERVALUE(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 1234.56, 'D3': 1234.56 }
  },
  {
    name: 'TEXTBEFORE',
    category: '03. 文字列',
    description: '指定文字の前のテキストを抽出',
    data: [
      ['テキスト', '区切り文字', '結果'],
      ['hello@example.com', '@', '=TEXTBEFORE(A2,B2)'],
      ['2024-01-15', '-', '=TEXTBEFORE(A3,B3)'],
      ['first,second,third', ',', '=TEXTBEFORE(A4,B4)']
    ],
    expectedValues: { 'C2': 'hello', 'C3': '2024', 'C4': 'first' }
  },
  {
    name: 'TEXTAFTER',
    category: '03. 文字列',
    description: '指定文字の後のテキストを抽出',
    data: [
      ['テキスト', '区切り文字', '結果'],
      ['hello@example.com', '@', '=TEXTAFTER(A2,B2)'],
      ['2024-01-15', '-', '=TEXTAFTER(A3,B3)'],
      ['first,second,third', ',', '=TEXTAFTER(A4,B4)']
    ],
    expectedValues: { 'C2': 'example.com', 'C3': '01-15', 'C4': 'second,third' }
  },
  {
    name: 'TEXTSPLIT',
    category: '03. 文字列',
    description: 'テキストを分割',
    data: [
      ['テキスト', '区切り文字', '分割結果'],
      ['apple,banana,orange', ',', '=TEXTSPLIT(A2,B2)'],
      ['2024-01-15', '-', '=TEXTSPLIT(A3,B3)'],
      ['one two three', ' ', '=TEXTSPLIT(A4,B4)']
    ]
  },
  {
    name: 'UNICHAR',
    category: '03. 文字列',
    description: 'Unicode番号から文字を返す',
    data: [
      ['Unicode番号', '文字'],
      [65, '=UNICHAR(A2)'],
      [8364, '=UNICHAR(A3)'],
      [12354, '=UNICHAR(A4)']
    ],
    expectedValues: { 'B2': 'A', 'B3': '€', 'B4': 'あ' }
  },
  {
    name: 'UNICODE',
    category: '03. 文字列',
    description: '文字からUnicode番号を返す',
    data: [
      ['文字', 'Unicode番号'],
      ['A', '=UNICODE(A2)'],
      ['€', '=UNICODE(A3)'],
      ['あ', '=UNICODE(A4)']
    ],
    expectedValues: { 'B2': 65, 'B3': 8364, 'B4': 12354 }
  },
  {
    name: 'T',
    category: '03. 文字列',
    description: 'テキストのみを返す',
    data: [
      ['値', 'テキスト'],
      ['Hello', '=T(A2)'],
      [123, '=T(A3)'],
      ['=TRUE()', '=T(A4)']
    ],
    expectedValues: { 'B2': 'Hello', 'B3': '', 'B4': '' }
  },
  {
    name: 'ASC',
    category: '03. 文字列',
    description: '全角を半角に変換',
    data: [
      ['全角文字', '半角文字'],
      ['ＨＥＬＬＯ', '=ASC(A2)'],
      ['１２３４５', '=ASC(A3)'],
      ['アイウエオ', '=ASC(A4)']
    ],
    expectedValues: { 'B2': 'HELLO', 'B3': '12345', 'B4': 'ｱｲｳｴｵ' }
  },
  {
    name: 'JIS',
    category: '03. 文字列',
    description: '半角を全角に変換',
    data: [
      ['半角文字', '全角文字'],
      ['HELLO', '=JIS(A2)'],
      ['12345', '=JIS(A3)'],
      ['ｱｲｳｴｵ', '=JIS(A4)']
    ],
    expectedValues: { 'B2': 'ＨＥＬＬＯ', 'B3': '１２３４５', 'B4': 'アイウエオ' }
  },
  {
    name: 'DBCS',
    category: '03. 文字列',
    description: '半角を全角に変換（DBCS）',
    data: [
      ['半角文字', '全角文字'],
      ['abc123', '=DBCS(A2)'],
      ['ｶﾀｶﾅ', '=DBCS(A3)']
    ],
    expectedValues: { 'B2': 'ａｂｃ１２３', 'B3': 'カタカナ' }
  },
  {
    name: 'LENB',
    category: '03. 文字列',
    description: 'バイト数を返す',
    data: [
      ['文字列', 'バイト数'],
      ['Hello', '=LENB(A2)'],
      ['こんにちは', '=LENB(A3)']
    ],
    expectedValues: { 'B2': 5, 'B3': 10 }
  },
  {
    name: 'FINDB',
    category: '03. 文字列',
    description: 'バイト位置を検索',
    data: [
      ['検索文字', '対象文字列', 'バイト位置'],
      ['o', 'Hello', '=FINDB(A2,B2)'],
      ['に', 'こんにちは', '=FINDB(A3,B3)']
    ],
    expectedValues: { 'C2': 5 }
  },
  {
    name: 'SEARCHB',
    category: '03. 文字列',
    description: 'バイト位置を検索（大小区別なし）',
    data: [
      ['検索文字', '対象文字列', 'バイト位置'],
      ['LO', 'Hello', '=SEARCHB(A2,B2)'],
      ['に', 'こんにちは', '=SEARCHB(A3,B3)']
    ]
  },
  {
    name: 'REPLACEB',
    category: '03. 文字列',
    description: 'バイト位置で置換',
    data: [
      ['文字列', '開始位置', 'バイト数', '新文字列', '結果'],
      ['Hello', 3, 2, 'XX', '=REPLACEB(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 'HeXXo' }
  }
];
