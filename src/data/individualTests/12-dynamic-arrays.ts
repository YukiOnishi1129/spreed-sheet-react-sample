import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const dynamicArraysTests: IndividualFunctionTest[] = [
  {
    name: 'FILTER',
    category: '12. 動的配列',
    description: '条件でフィルタ',
    data: [
      ['名前', '年齢', 'フィルタ結果'],
      ['田中', 25, '=FILTER(A2:B4,B2:B4>25)'],
      ['佐藤', 30, ''],
      ['鈴木', 28, '']
    ]
  },
  {
    name: 'SORT',
    category: '12. 動的配列',
    description: '並べ替え',
    data: [
      ['名前', '点数', 'ソート結果'],
      ['田中', 85, '=SORT(A2:B4,2,-1)'],
      ['佐藤', 92, ''],
      ['鈴木', 78, '']
    ]
  },
  {
    name: 'UNIQUE',
    category: '12. 動的配列',
    description: '一意の値を抽出',
    data: [
      ['値', '一意の値'],
      ['A', '=UNIQUE(A2:A6)'],
      ['B', ''],
      ['A', ''],
      ['C', ''],
      ['B', '']
    ]
  },
  {
    name: 'TRANSPOSE',
    category: '12. 動的配列',
    description: '行列を入れ替え',
    data: [
      ['', 'A', 'B', 'C'],
      ['1', 1, 2, 3],
      ['2', 4, 5, 6],
      ['', '', '', ''],
      ['転置', '=TRANSPOSE(B2:D3)', '', '']
    ]
  },
  {
    name: 'SEQUENCE',
    category: '12. 動的配列',
    description: '連続値を生成',
    data: [
      ['行数', '列数', '開始', 'ステップ', '連続値'],
      [5, 1, 1, 2, '=SEQUENCE(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'RANDARRAY',
    category: '12. 動的配列',
    description: 'ランダム配列を生成',
    data: [
      ['行数', '列数', '最小', '最大', '整数', 'ランダム配列'],
      [3, 3, 1, 10, 'TRUE', '=RANDARRAY(A2,B2,C2,D2,E2)']
    ]
  },
  {
    name: 'LAMBDA',
    category: '12. 動的配列',
    description: 'カスタム関数を作成',
    data: [
      ['値', '二乗'],
      [5, '=LAMBDA(x,x*x)(A2)'],
      [7, '=LAMBDA(x,x*x)(A3)']
    ],
    expectedValues: { 'B2': 25, 'B3': 49 }
  },
  {
    name: 'LET',
    category: '12. 動的配列',
    description: '変数を定義して使用',
    data: [
      ['長さ', '幅', '面積'],
      [10, 5, '=LET(l,A2,w,B2,l*w)'],
      [7, 3, '=LET(l,A3,w,B3,l*w)']
    ],
    expectedValues: { 'C2': 50, 'C3': 21 }
  },
  {
    name: 'HSTACK',
    category: '12. 動的配列',
    description: '水平方向に結合',
    data: [
      ['配列1', '', '配列2', '', '結合結果'],
      [1, 2, 'A', 'B', '=HSTACK(A2:B3,C2:D3)'],
      [3, 4, 'C', 'D', '']
    ]
  },
  {
    name: 'VSTACK',
    category: '12. 動的配列',
    description: '垂直方向に結合',
    data: [
      ['配列1', '配列2', '結合結果'],
      [1, 2, '=VSTACK(A2:B3,A5:B6)'],
      [3, 4, ''],
      ['', '', ''],
      [5, 6, ''],
      [7, 8, '']
    ]
  },
  {
    name: 'BYROW',
    category: '12. 動的配列',
    description: '各行に関数を適用',
    data: [
      ['値1', '値2', '値3', '行の合計'],
      [1, 2, 3, '=BYROW(A2:C4,LAMBDA(row,SUM(row)))'],
      [4, 5, 6, ''],
      [7, 8, 9, '']
    ],
    expectedValues: { 'D2': 6, 'D3': 15, 'D4': 24 }
  },
  {
    name: 'BYCOL',
    category: '12. 動的配列',
    description: '各列に関数を適用',
    data: [
      ['値1', '値2', '値3'],
      [1, 4, 7],
      [2, 5, 8],
      [3, 6, 9],
      ['列の合計', '', ''],
      ['=BYCOL(A2:C4,LAMBDA(col,SUM(col)))', '', '']
    ],
    expectedValues: { 'A6': 6, 'B6': 15, 'C6': 24 }
  },
  {
    name: 'MAKEARRAY',
    category: '12. 動的配列',
    description: '配列を生成',
    data: [
      ['行数', '列数', '配列'],
      [3, 4, '=MAKEARRAY(A2,B2,LAMBDA(r,c,r*c))']
    ]
  },
  {
    name: 'MAP',
    category: '12. 動的配列',
    description: '配列の各要素に関数を適用',
    data: [
      ['値', '二乗'],
      [1, '=MAP(A2:A5,LAMBDA(x,x^2))'],
      [2, ''],
      [3, ''],
      [4, '']
    ],
    expectedValues: { 'B2': 1, 'B3': 4, 'B4': 9, 'B5': 16 }
  },
  {
    name: 'REDUCE',
    category: '12. 動的配列',
    description: '配列を集約',
    data: [
      ['値', '累積合計'],
      [1, ''],
      [2, ''],
      [3, ''],
      [4, '=REDUCE(0,A2:A5,LAMBDA(acc,val,acc+val))']
    ],
    expectedValues: { 'B5': 10 }
  },
  {
    name: 'SCAN',
    category: '12. 動的配列',
    description: '配列を累積的に処理',
    data: [
      ['値', '累積合計'],
      [1, '=SCAN(0,A2:A5,LAMBDA(acc,val,acc+val))'],
      [2, ''],
      [3, ''],
      [4, '']
    ],
    expectedValues: { 'B2': 1, 'B3': 3, 'B4': 6, 'B5': 10 }
  },
  {
    name: 'SORTBY',
    category: '12. 動的配列',
    description: '別の配列でソート',
    data: [
      ['商品', '売上', 'ソート結果'],
      ['りんご', 300, '=SORTBY(A2:A5,B2:B5)'],
      ['バナナ', 100, ''],
      ['オレンジ', 400, ''],
      ['ぶどう', 200, '']
    ]
  },
  {
    name: 'TAKE',
    category: '12. 動的配列',
    description: '配列の一部を取得',
    data: [
      ['データ', '上位3つ'],
      [100, '=TAKE(A2:A7,3)'],
      [200, ''],
      [300, ''],
      [400, ''],
      [500, ''],
      [600, '']
    ]
  },
  {
    name: 'DROP',
    category: '12. 動的配列',
    description: '配列の一部を削除',
    data: [
      ['データ', '最初の2つを削除'],
      [100, '=DROP(A2:A7,2)'],
      [200, ''],
      [300, ''],
      [400, ''],
      [500, ''],
      [600, '']
    ]
  },
  {
    name: 'EXPAND',
    category: '12. 動的配列',
    description: '配列を拡張',
    data: [
      ['元データ', '', '拡張結果'],
      [1, 2, '=EXPAND(A2:B3,4,3,"-")'],
      [3, 4, ''],
      ['', '', ''],
      ['', '', '']
    ]
  },
  {
    name: 'TOCOL',
    category: '12. 動的配列',
    description: '配列を1列に変換',
    data: [
      ['配列', '', '1列変換'],
      [1, 2, '=TOCOL(A2:B3)'],
      [3, 4, '']
    ]
  },
  {
    name: 'TOROW',
    category: '12. 動的配列',
    description: '配列を1行に変換',
    data: [
      ['配列', '', '', '1行変換'],
      [1, 2, '', '=TOROW(A2:B3)'],
      [3, 4, '', '']
    ]
  },
  {
    name: 'CHOOSEROWS',
    category: '12. 動的配列',
    description: '指定した行を選択',
    data: [
      ['データ', '列2', '選択結果'],
      ['A', 1, '=CHOOSEROWS(A2:B5,1,3)'],
      ['B', 2, ''],
      ['C', 3, ''],
      ['D', 4, '']
    ]
  },
  {
    name: 'CHOOSECOLS',
    category: '12. 動的配列',
    description: '指定した列を選択',
    data: [
      ['データ', '列2', '列3', '選択結果'],
      [1, 2, 3, '=CHOOSECOLS(A2:C4,1,3)'],
      [4, 5, 6, ''],
      [7, 8, 9, '']
    ]
  },
  {
    name: 'WRAPROWS',
    category: '12. 動的配列',
    description: '配列を行で折り返し',
    data: [
      ['データ', '', '', '折り返し結果'],
      [1, 2, 3, '=WRAPROWS(A2:C3,3)'],
      [4, 5, 6, '']
    ]
  },
  {
    name: 'WRAPCOLS',
    category: '12. 動的配列',
    description: '配列を列で折り返し',
    data: [
      ['データ', '', '折り返し結果'],
      [1, 2, '=WRAPCOLS(A2:B4,2)'],
      [3, 4, ''],
      [5, 6, '']
    ]
  }
];
