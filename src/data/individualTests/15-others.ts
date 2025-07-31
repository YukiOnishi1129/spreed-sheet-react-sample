import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const othersTests: IndividualFunctionTest[] = [
  {
    name: 'ISOMITTED',
    category: '15. その他',
    description: '引数が省略されているか判定',
    data: [
      ['値', '省略判定1', '省略判定2', '省略判定3'],
      ['test', '=ISOMITTED("")', '=ISOMITTED(A2)', '=ISOMITTED()'],
      ['', '=ISOMITTED(A3)', '', '']
    ],
    expectedValues: { 'B2': true, 'C2': false, 'D2': true, 'B3': true }
  },
  {
    name: 'STOCKHISTORY',
    category: '15. その他',
    description: '株価履歴を取得',
    data: [
      ['銘柄', '開始日', '終了日', '履歴'],
      ['AAPL', '2024-01-01', '2024-01-31', '=STOCKHISTORY(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 'Date' }
  },
  {
    name: 'GPT',
    category: '15. その他', 
    description: 'GPTによるテキスト生成',
    data: [
      ['プロンプト', '生成結果'],
      ['translate hello', '=GPT(A2)']
    ],
    expectedValues: { 'B2': 'Translation: hello' }
  }
];
