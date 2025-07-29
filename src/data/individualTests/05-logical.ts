import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const logicalTests: IndividualFunctionTest[] = [
  {
    name: 'IF',
    category: '05. 論理',
    description: '条件分岐',
    data: [
      ['値', '結果'],
      [80, '=IF(A2>=60,"合格","不合格")'],
      [50, '=IF(A3>=60,"合格","不合格")']
    ],
    expectedValues: { 'B2': '合格', 'B3': '不合格' }
  },
  {
    name: 'IFS',
    category: '05. 論理',
    description: '複数条件分岐',
    data: [
      ['点数', '評価'],
      [90, '=IFS(A2>=90,"A",A2>=80,"B",A2>=70,"C",TRUE,"D")'],
      [75, '=IFS(A3>=90,"A",A3>=80,"B",A3>=70,"C",TRUE,"D")']
    ],
    expectedValues: { 'B2': 'A', 'B3': 'C' }
  },
  {
    name: 'AND',
    category: '05. 論理',
    description: '論理積（すべて真）',
    data: [
      ['条件1', '条件2', '結果'],
      [true, true, '=AND(A2,B2)'],
      [true, false, '=AND(A3,B3)']
    ],
    expectedValues: { 'C2': true, 'C3': false }
  },
  {
    name: 'OR',
    category: '05. 論理',
    description: '論理和（いずれか真）',
    data: [
      ['条件1', '条件2', '結果'],
      [true, false, '=OR(A2,B2)'],
      [false, false, '=OR(A3,B3)']
    ],
    expectedValues: { 'C2': true, 'C3': false }
  },
  {
    name: 'NOT',
    category: '05. 論理',
    description: '論理否定',
    data: [
      ['条件', '結果'],
      [true, '=NOT(A2)'],
      [false, '=NOT(A3)']
    ],
    expectedValues: { 'B2': false, 'B3': true }
  },
  {
    name: 'XOR',
    category: '05. 論理',
    description: '排他的論理和',
    data: [
      ['条件1', '条件2', '結果'],
      [true, true, '=XOR(A2,B2)'],
      [true, false, '=XOR(A3,B3)']
    ],
    expectedValues: { 'C2': false, 'C3': true }
  },
  {
    name: 'TRUE',
    category: '05. 論理',
    description: '真を返す',
    data: [
      ['結果'],
      ['=TRUE()']
    ],
    expectedValues: { 'A2': true }
  },
  {
    name: 'FALSE',
    category: '05. 論理',
    description: '偽を返す',
    data: [
      ['結果'],
      ['=FALSE()']
    ],
    expectedValues: { 'A2': false }
  },
  {
    name: 'IFERROR',
    category: '05. 論理',
    description: 'エラー時の値を指定',
    data: [
      ['値1', '値2', '結果'],
      [10, 0, '=IFERROR(A2/B2,"エラー")'],
      [10, 2, '=IFERROR(A3/B3,"エラー")']
    ],
    expectedValues: { 'C2': 'エラー', 'C3': 5 }
  },
  {
    name: 'IFNA',
    category: '05. 論理',
    description: '#N/Aエラー時の値',
    data: [
      ['検索値', '結果', 'NA関数', 'キー', '値'],
      ['存在しない', '=IFNA(VLOOKUP(A2,D:E,2,FALSE),"見つかりません")', '=IFNA(NA(),"NAエラー")', 'Item1', 'Value1'],
      ['Item2', '=IFNA(VLOOKUP(A3,D:E,2,FALSE),"見つかりません")', '', 'Item2', 'Value2'],
      ['', '', '', 'Item3', 'Value3']
    ],
    expectedValues: { 'B2': '見つかりません', 'B3': 'Value2', 'C2': 'NAエラー' }
  },
  {
    name: 'SWITCH',
    category: '05. 論理',
    description: '値に応じて切り替え',
    data: [
      ['値', '結果'],
      [1, '=SWITCH(A2,1,"一",2,"二",3,"三","その他")'],
      [2, '=SWITCH(A3,1,"一",2,"二",3,"三","その他")'],
      [4, '=SWITCH(A4,1,"一",2,"二",3,"三","その他")']
    ],
    expectedValues: { 'B2': '一', 'B3': '二', 'B4': 'その他' }
  }
];
