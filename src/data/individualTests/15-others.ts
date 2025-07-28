import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const othersTests: IndividualFunctionTest[] = [
  {
    name: 'ISOMITTED',
    category: '15. その他',
    description: '引数が省略されているか判定',
    data: [
      ['値', '省略判定'],
      ['=LAMBDA(x,y,ISOMITTED(y))(A2)', '=A2'],
      ['=LAMBDA(x,y,ISOMITTED(y))(A3,B3)', '=A3']
    ],
    expectedValues: { 'A2': true, 'B2': true, 'A3': false, 'B3': false }
  },
  {
    name: 'STOCKHISTORY',
    category: '15. その他',
    description: '株価履歴を取得',
    data: [
      ['銘柄', '開始日', '終了日', '間隔', 'ヘッダー', 'プロパティ', '履歴'],
      ['MSFT', '2024/1/1', '2024/1/31', 0, 1, 0, '=STOCKHISTORY(A2,B2,C2,D2,E2,F2)']
    ],
    expectedValues: { 'G2': '#N/A' }
  },
  {
    name: 'GPT',
    category: '15. その他',
    description: 'GPTによるテキスト生成',
    data: [
      ['プロンプト', '生成結果'],
      ['Excelの便利な使い方を教えて', '=GPT(A2)']
    ],
    expectedValues: { 'B2': '#N/A' }
  }
];
