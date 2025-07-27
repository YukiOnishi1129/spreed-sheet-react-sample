import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const informationTests: IndividualFunctionTest[] = [
  {
    name: 'ISBLANK',
    category: '09. 情報',
    description: '空白セルか判定',
    data: [
      ['値', '空白判定'],
      ['テキスト', '=ISBLANK(A2)'],
      ['', '=ISBLANK(A3)'],
      [0, '=ISBLANK(A4)']
    ],
    expectedValues: { 'B2': false, 'B3': true, 'B4': false }
  },
  {
    name: 'ISERROR',
    category: '09. 情報',
    description: 'エラー値か判定',
    data: [
      ['値', 'エラー判定'],
      ['=1/0', '=ISERROR(A2)'],
      [100, '=ISERROR(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'ISNA',
    category: '09. 情報',
    description: '#N/Aエラーか判定',
    data: [
      ['値', '#N/A判定'],
      ['=NA()', '=ISNA(A2)'],
      ['#DIV/0!', '=ISNA(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'ISTEXT',
    category: '09. 情報',
    description: '文字列か判定',
    data: [
      ['値', '文字列判定'],
      ['テキスト', '=ISTEXT(A2)'],
      [123, '=ISTEXT(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'ISNUMBER',
    category: '09. 情報',
    description: '数値か判定',
    data: [
      ['値', '数値判定'],
      [123, '=ISNUMBER(A2)'],
      ['123', '=ISNUMBER(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'ISLOGICAL',
    category: '09. 情報',
    description: '論理値か判定',
    data: [
      ['値', '論理値判定'],
      ['=TRUE()', '=ISLOGICAL(A2)'],
      [1, '=ISLOGICAL(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'TYPE',
    category: '09. 情報',
    description: 'データ型を返す',
    data: [
      ['値', 'データ型'],
      [123, '=TYPE(A2)'],
      ['テキスト', '=TYPE(A3)'],
      ['=TRUE()', '=TYPE(A4)']
    ],
    expectedValues: { 'B2': 1, 'B3': 2, 'B4': 4 }
  },
  {
    name: 'N',
    category: '09. 情報',
    description: '数値に変換',
    data: [
      ['値', '数値変換'],
      [7, '=N(A2)'],
      ['7', '=N(A3)'],
      ['=TRUE()', '=N(A4)']
    ],
    expectedValues: { 'B2': 7, 'B3': 0, 'B4': 1 }
  },
  {
    name: 'ERROR.TYPE',
    category: '09. 情報',
    description: 'エラーの種類',
    data: [
      ['エラー', 'エラー番号'],
      ['=1/0', '=ERROR.TYPE(A2)'],
      ['=NA()', '=ERROR.TYPE(A3)']
    ],
    expectedValues: { 'B2': 2, 'B3': 7 }
  },
  {
    name: 'SHEET',
    category: '09. 情報',
    description: 'シート番号を返す',
    data: [
      ['シート番号'],
      ['=SHEET()']
    ]
  },
  {
    name: 'SHEETS',
    category: '09. 情報',
    description: 'シート数を返す',
    data: [
      ['シート数'],
      ['=SHEETS()']
    ]
  },
  {
    name: 'ISFORMULA',
    category: '09. 情報',
    description: '数式か判定',
    data: [
      ['値', '数式判定'],
      ['=A1+1', '=ISFORMULA(A2)'],
      [100, '=ISFORMULA(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'ISREF',
    category: '09. 情報',
    description: '参照か判定',
    data: [
      ['値', '参照判定'],
      ['A1', '=ISREF(A1)'],
      ['テキスト', '=ISREF("テキスト")']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'CELL',
    category: '09. 情報',
    description: 'セル情報を取得',
    data: [
      ['情報タイプ', '参照', '結果'],
      ['type', 'A2', '=CELL(A2,B2)'],
      ['address', 'B3', '=CELL(A3,B3)']
    ]
  },
  {
    name: 'INFO',
    category: '09. 情報',
    description: 'システム情報',
    data: [
      ['情報タイプ', '結果'],
      ['numfile', '=INFO(A2)'],
      ['osversion', '=INFO(A3)']
    ]
  },
  {
    name: 'ISBETWEEN',
    category: '09. 情報',
    description: '値が範囲内か判定',
    data: [
      ['値', '下限', '上限', '範囲内判定'],
      [5, 1, 10, '=ISBETWEEN(A2,B2,C2)'],
      [15, 1, 10, '=ISBETWEEN(A3,B3,C3)'],
      [10, 1, 10, '=ISBETWEEN(A4,B4,C4)']
    ],
    expectedValues: { 'D2': true, 'D3': false, 'D4': true }
  }
];
