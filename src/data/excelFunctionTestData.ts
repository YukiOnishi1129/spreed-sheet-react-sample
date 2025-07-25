export interface FunctionTestCase {
  name: string;
  category: string;
  description: string;
  data: (string | number | null)[][];
  expectedValues?: { [key: string]: string | number | boolean };
}

export const excelFunctionTests: FunctionTestCase[] = [
  // 数学関数のテスト
  {
    name: 'SUM関数',
    category: '数学',
    description: 'SUM関数の基本的な動作をテスト',
    data: [
      ['数値1', '数値2', '数値3', '合計'],
      [10, 20, 30, '=SUM(A2:C2)'],
      [100, 200, 300, '=SUM(A3:C3)'],
      [-10, -20, 30, '=SUM(A4:C4)'],
      [1.5, 2.5, 3.5, '=SUM(A5:C5)']
    ],
    expectedValues: {
      'D2': 60,
      'D3': 600,
      'D4': 0,
      'D5': 7.5
    }
  },
  {
    name: 'AVERAGE関数',
    category: '数学',
    description: 'AVERAGE関数の基本的な動作をテスト',
    data: [
      ['値1', '値2', '値3', '値4', '平均'],
      [10, 20, 30, 40, '=AVERAGE(A2:D2)'],
      [5, 10, 15, 20, '=AVERAGE(A3:D3)'],
      [100, 0, 50, null, '=AVERAGE(A4:C4)']
    ],
    expectedValues: {
      'E2': 25,
      'E3': 12.5,
      'E4': 50
    }
  },
  {
    name: 'MIN/MAX関数',
    category: '数学',
    description: 'MIN/MAX関数の基本的な動作をテスト',
    data: [
      ['値1', '値2', '値3', '最小値', '最大値'],
      [10, 5, 20, '=MIN(A2:C2)', '=MAX(A2:C2)'],
      [-10, 0, 10, '=MIN(A3:C3)', '=MAX(A3:C3)'],
      [100, 50, 75, '=MIN(A4:C4)', '=MAX(A4:C4)']
    ],
    expectedValues: {
      'D2': 5,
      'E2': 20,
      'D3': -10,
      'E3': 10,
      'D4': 50,
      'E4': 100
    }
  },
  {
    name: 'COUNT/COUNTA関数',
    category: '数学',
    description: 'COUNT/COUNTA関数の基本的な動作をテスト',
    data: [
      ['値1', '値2', '値3', '値4', 'COUNT', 'COUNTA'],
      [10, 20, '', 30, '=COUNT(A2:D2)', '=COUNTA(A2:D2)'],
      ['テキスト', 100, '', 200, '=COUNT(A3:D3)', '=COUNTA(A3:D3)'],
      ['', '', '', '', '=COUNT(A4:D4)', '=COUNTA(A4:D4)']
    ],
    expectedValues: {
      'E2': 3,
      'F2': 3,
      'E3': 2,
      'F3': 3,
      'E4': 0,
      'F4': 0
    }
  },
  // 論理関数のテスト
  {
    name: 'IF関数',
    category: '論理',
    description: 'IF関数の基本的な動作をテスト',
    data: [
      ['値', '条件', '結果'],
      [10, '>5', '=IF(A2>5,"大きい","小さい")'],
      [3, '>5', '=IF(A3>5,"大きい","小さい")'],
      [100, '>=100', '=IF(A4>=100,"合格","不合格")'],
      [-5, '<0', '=IF(A5<0,"負の数","正の数")']
    ],
    expectedValues: {
      'C2': '大きい',
      'C3': '小さい',
      'C4': '合格',
      'C5': '負の数'
    }
  },
  {
    name: 'AND/OR関数',
    category: '論理',
    description: 'AND/OR関数の基本的な動作をテスト',
    data: [
      ['値1', '値2', 'AND結果', 'OR結果'],
      [10, 20, '=AND(A2>5,B2>15)', '=OR(A2>15,B2>15)'],
      [5, 3, '=AND(A3>5,B3>5)', '=OR(A3>5,B3>5)'],
      [0, 10, '=AND(A4>-1,B4>5)', '=OR(A4>5,B4>5)']
    ],
    expectedValues: {
      'C2': true,
      'D2': true,
      'C3': false,
      'D3': false,
      'C4': true,
      'D4': true
    }
  },
  // 文字列関数のテスト
  {
    name: 'CONCATENATE/CONCAT関数',
    category: '文字列',
    description: '文字列結合関数の基本的な動作をテスト',
    data: [
      ['名前', '苗字', '結合結果'],
      ['太郎', '田中', '=CONCATENATE(B2," ",A2)'],
      ['花子', '山田', '=CONCAT(B3,A3)'],
      ['A', 'B', '=CONCATENATE(A4,B4,"C")']
    ],
    expectedValues: {
      'C2': '田中 太郎',
      'C3': '山田花子',
      'C4': 'ABC'
    }
  },
  {
    name: 'LEN/LEFT/RIGHT/MID関数',
    category: '文字列',
    description: '文字列操作関数の基本的な動作をテスト',
    data: [
      ['文字列', '長さ', '左3文字', '右3文字', '中間'],
      ['Hello World', '=LEN(A2)', '=LEFT(A2,3)', '=RIGHT(A2,3)', '=MID(A2,3,3)'],
      ['日本語テスト', '=LEN(A3)', '=LEFT(A3,3)', '=RIGHT(A3,3)', null],
      ['1234567890', '=LEN(A4)', '=LEFT(A4,3)', '=RIGHT(A4,3)', null]
    ],
    expectedValues: {
      'B2': 11,
      'C2': 'Hel',
      'D2': 'rld',
      'E2': 'llo',
      'B3': 6,
      'C3': '日本語',
      'D3': 'スト',
      'B4': 10,
      'C4': '123',
      'D4': '890'
    }
  },
  // 日付関数のテスト
  {
    name: 'TODAY/NOW関数',
    category: '日付',
    description: '日付関数の基本的な動作をテスト',
    data: [
      ['関数名', '結果'],
      ['TODAY', '=TODAY()'],
      ['NOW', '=NOW()'],
      ['年', '=YEAR(TODAY())'],
      ['月', '=MONTH(TODAY())'],
      ['日', '=DAY(TODAY())']
    ]
  },
  // 検索関数のテスト
  {
    name: 'VLOOKUP関数',
    category: '検索',
    description: 'VLOOKUP関数の基本的な動作をテスト',
    data: [
      ['商品ID', '商品名', '価格', '', '検索ID', '結果'],
      ['A001', 'りんご', 100, '', 'A002', '=VLOOKUP(E2,A2:C4,2,FALSE)'],
      ['A002', 'みかん', 80, '', 'A003', '=VLOOKUP(E3,A2:C4,3,FALSE)'],
      ['A003', 'バナナ', 120, '', 'A001', '=VLOOKUP(E4,A2:C4,2,FALSE)']
    ],
    expectedValues: {
      'F2': 'みかん',
      'F3': 120,
      'F4': 'りんご'
    }
  },
  // 統計関数のテスト
  {
    name: 'STDEV/VAR関数',
    category: '統計',
    description: '標準偏差と分散の計算をテスト',
    data: [
      ['データ1', 'データ2', 'データ3', 'データ4', '標準偏差', '分散'],
      [10, 20, 30, 40, '=STDEV(A2:D2)', '=VAR(A2:D2)'],
      [5, 5, 5, 5, '=STDEV(A3:D3)', '=VAR(A3:D3)'],
      [1, 2, 3, 4, '=STDEV(A4:D4)', '=VAR(A4:D4)']
    ],
    expectedValues: {
      'E3': 0,
      'F3': 0
    }
  },
  // 財務関数のテスト
  {
    name: 'PMT関数',
    category: '財務',
    description: 'ローン返済額の計算をテスト',
    data: [
      ['元本', '年利率', '返済回数', '月額返済額'],
      [1000000, 0.03, 36, '=PMT(B2/12,C2,-A2)'],
      [500000, 0.05, 24, '=PMT(B3/12,C3,-A3)'],
      [2000000, 0.02, 60, '=PMT(B4/12,C4,-A4)']
    ]
  }
];

// カテゴリ別にテストケースを取得する関数
export function getTestsByCategory(category: string): FunctionTestCase[] {
  return excelFunctionTests.filter(test => test.category === category);
}

// すべてのカテゴリを取得する関数
export function getAllCategories(): string[] {
  return [...new Set(excelFunctionTests.map(test => test.category))];
}