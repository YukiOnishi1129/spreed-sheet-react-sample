import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const lookupReferenceTests: IndividualFunctionTest[] = [
  {
    name: 'VLOOKUP',
    category: '06. 検索',
    description: '垂直方向検索',
    data: [
      ['ID', '名前', '部署', '', '検索ID', '結果'],
      [101, '田中', '営業', '', 102, '=VLOOKUP(E2,A2:C4,2,FALSE)'],
      [102, '佐藤', '技術', '', '', ''],
      [103, '鈴木', '人事', '', '', '']
    ],
    expectedValues: { 'F2': '佐藤' }
  },
  {
    name: 'HLOOKUP',
    category: '06. 検索',
    description: '水平方向検索',
    data: [
      ['項目', '1月', '2月', '3月'],
      ['売上', 100, 200, 300],
      ['利益', 10, 20, 30],
      ['', '', '', ''],
      ['検索項目', '2月', '=HLOOKUP(B5,A1:D3,2,FALSE)', '']
    ],
    expectedValues: { 'C5': 200 }
  },
  {
    name: 'XLOOKUP',
    category: '06. 検索',
    description: '柔軟な検索',
    data: [
      ['ID', '名前', '', '検索ID', '結果'],
      [101, '田中', '', 102, '=XLOOKUP(D2,A2:A4,B2:B4,"見つかりません")'],
      [102, '佐藤', '', '', ''],
      [103, '鈴木', '', '', '']
    ],
    expectedValues: { 'E2': '佐藤' }
  },
  {
    name: 'INDEX',
    category: '06. 検索',
    description: '位置から値を取得',
    data: [
      ['データ', '', '', '行', '列', '結果'],
      ['A', 'B', 'C', 2, 2, '=INDEX(A2:C4,D2,E2)'],
      ['D', 'E', 'F', '', '', ''],
      ['G', 'H', 'I', '', '', '']
    ],
    expectedValues: { 'F2': 'E' }
  },
  {
    name: 'MATCH',
    category: '06. 検索',
    description: '値の位置を検索',
    data: [
      ['値', '', '', '検索値', '位置'],
      [10, 20, 30, 20, '=MATCH(D2,A2:C2,0)']
    ],
    expectedValues: { 'E2': 2 }
  },
  {
    name: 'CHOOSE',
    category: '06. 検索',
    description: 'リストから選択',
    data: [
      ['インデックス', '結果'],
      [2, '=CHOOSE(A2,"一","二","三")'],
      [3, '=CHOOSE(A3,"一","二","三")']
    ],
    expectedValues: { 'B2': '二', 'B3': '三' }
  },
  {
    name: 'OFFSET',
    category: '06. 検索',
    description: 'オフセット参照',
    data: [
      ['A1', 'B1', 'C1'],
      ['A2', 'B2', 'C2'],
      ['A3', 'B3', 'C3'],
      ['', '', ''],
      ['結果', '=OFFSET(A1,1,1)', '']
    ],
    expectedValues: { 'B5': 'B2' }
  },
  {
    name: 'INDIRECT',
    category: '06. 検索',
    description: '文字列を参照に変換',
    data: [
      ['値1', '値2', '', '参照文字列', '結果'],
      [100, 200, '', '"A2"', '=INDIRECT(D2)']
    ],
    expectedValues: { 'E2': 100 }
  },
  {
    name: 'ROW',
    category: '06. 検索',
    description: '行番号を返す',
    data: [
      ['行番号'],
      ['=ROW()'],
      ['=ROW()'],
      ['=ROW()']
    ],
    expectedValues: { 'A2': 2, 'A3': 3, 'A4': 4 }
  },
  {
    name: 'COLUMN',
    category: '06. 検索',
    description: '列番号を返す',
    data: [
      ['列1', '列2', '列3'],
      ['=COLUMN()', '=COLUMN()', '=COLUMN()']
    ],
    expectedValues: { 'A2': 1, 'B2': 2, 'C2': 3 }
  },
  {
    name: 'XMATCH',
    category: '06. 検索',
    description: '配列内の項目の位置を返す',
    data: [
      ['データ', '検索値', '位置'],
      ['りんご', 'オレンジ', '=XMATCH(B2,A2:A5)'],
      ['バナナ', '', ''],
      ['オレンジ', '', ''],
      ['ぶどう', '', '']
    ],
    expectedValues: { 'C2': 3 }
  },
  {
    name: 'ROWS',
    category: '06. 検索',
    description: '範囲の行数を返す',
    data: [
      ['データ', '', '行数'],
      [1, 2, '=ROWS(A2:B5)'],
      [3, 4, ''],
      [5, 6, ''],
      [7, 8, '']
    ],
    expectedValues: { 'C2': 4 }
  },
  {
    name: 'COLUMNS',
    category: '06. 検索',
    description: '範囲の列数を返す',
    data: [
      ['データ', '', '', '', '列数'],
      [1, 2, 3, 4, '=COLUMNS(A2:D2)']
    ],
    expectedValues: { 'E2': 4 }
  },
  {
    name: 'ADDRESS',
    category: '06. 検索',
    description: 'セルアドレスを文字列で返す',
    data: [
      ['行番号', '列番号', '絶対参照', 'アドレス'],
      [2, 3, 1, '=ADDRESS(A2,B2,C2)'],
      [5, 1, 4, '=ADDRESS(A3,B3,C3)']
    ],
    expectedValues: { 'D2': '$C$2', 'D3': 'A5' }
  },
  {
    name: 'HYPERLINK',
    category: '06. 検索',
    description: 'ハイパーリンクを作成',
    data: [
      ['URL', '表示テキスト', 'リンク'],
      ['https://example.com', 'Example', '=HYPERLINK(A2,B2)']
    ]
  },
  {
    name: 'LOOKUP',
    category: '06. 検索',
    description: 'ベクトル検索',
    data: [
      ['検索値', '検索範囲', '', '結果範囲', '', '結果'],
      [4.5, 1, 3, 'A', 'B', '=LOOKUP(A2,B2:C2,D2:E2)'],
      ['', 4.5, 6, 'C', 'D', '']
    ],
    expectedValues: { 'F2': 'B' }
  },
  {
    name: 'AREAS',
    category: '06. 検索',
    description: '参照の領域数',
    data: [
      ['数式', '領域数'],
      ['=A1:B2', '=AREAS(A2)'],
      ['=(A1:B2,D1:E2)', '=AREAS((A1:B2,D1:E2))']
    ],
    expectedValues: { 'B2': 1 }
  },
  {
    name: 'FORMULATEXT',
    category: '06. 検索',
    description: '数式を文字列として返す',
    data: [
      ['値', '数式テキスト'],
      ['=SUM(1,2,3)', '=FORMULATEXT(A2)'],
      [100, '=FORMULATEXT(A3)']
    ],
    expectedValues: { 'B2': '=SUM(1,2,3)' }
  },
  {
    name: 'GETPIVOTDATA',
    category: '06. 検索',
    description: 'ピボットテーブルからデータ取得',
    data: [
      ['製品', '地域', '売上'],
      ['A', '東', 100],
      ['A', '西', 150],
      ['B', '東', 200],
      ['', '', ''],
      ['取得値', '=GETPIVOTDATA("売上",A1,"製品","A","地域","東")', '']
    ],
    expectedValues: { 'B6': 100 }
  }
];
