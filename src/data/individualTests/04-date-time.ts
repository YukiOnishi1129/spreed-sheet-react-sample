import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const dateTimeTests: IndividualFunctionTest[] = [
  {
    name: 'TODAY',
    category: '04. 日付',
    description: '今日の日付を返す',
    data: [
      ['今日の日付'],
      ['=TODAY()']
    ]
  },
  {
    name: 'NOW',
    category: '04. 日付',
    description: '現在の日時を返す',
    data: [
      ['現在の日時'],
      ['=NOW()']
    ]
  },
  {
    name: 'DATE',
    category: '04. 日付',
    description: '日付を作成',
    data: [
      ['年', '月', '日', '日付'],
      [2024, 12, 25, '=DATE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 45651 } // Excelのシリアル値
  },
  {
    name: 'TIME',
    category: '04. 日付',
    description: '時刻を作成',
    data: [
      ['時', '分', '秒', '時刻'],
      [13, 30, 45, '=TIME(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 0.563020833 }
  },
  {
    name: 'YEAR',
    category: '04. 日付',
    description: '年を抽出',
    data: [
      ['日付', '年'],
      ['2024/12/25', '=YEAR(A2)']
    ],
    expectedValues: { 'B2': 2024 }
  },
  {
    name: 'MONTH',
    category: '04. 日付',
    description: '月を抽出',
    data: [
      ['日付', '月'],
      ['2024/12/25', '=MONTH(A2)']
    ],
    expectedValues: { 'B2': 12 }
  },
  {
    name: 'DAY',
    category: '04. 日付',
    description: '日を抽出',
    data: [
      ['日付', '日'],
      ['2024/12/25', '=DAY(A2)']
    ],
    expectedValues: { 'B2': 25 }
  },
  {
    name: 'HOUR',
    category: '04. 日付',
    description: '時を抽出',
    data: [
      ['時刻', '時'],
      ['13:30:45', '=HOUR(A2)']
    ],
    expectedValues: { 'B2': 13 }
  },
  {
    name: 'MINUTE',
    category: '04. 日付',
    description: '分を抽出',
    data: [
      ['時刻', '分'],
      ['13:30:45', '=MINUTE(A2)']
    ],
    expectedValues: { 'B2': 30 }
  },
  {
    name: 'SECOND',
    category: '04. 日付',
    description: '秒を抽出',
    data: [
      ['時刻', '秒'],
      ['13:30:45', '=SECOND(A2)']
    ],
    expectedValues: { 'B2': 45 }
  },
  {
    name: 'WEEKDAY',
    category: '04. 日付',
    description: '曜日を数値で返す',
    data: [
      ['日付', '曜日番号'],
      ['2024/12/25', '=WEEKDAY(A2)'] // 水曜日 = 4
    ],
    expectedValues: { 'B2': 4 }
  },
  {
    name: 'DATEDIF',
    category: '04. 日付',
    description: '日付間の期間を計算',
    data: [
      ['開始日', '終了日', '期間（日）', '期間（月）', '期間（年）'],
      ['2024/1/1', '2024/12/31', '=DATEDIF(A2,B2,"D")', '=DATEDIF(A2,B2,"M")', '=DATEDIF(A2,B2,"Y")']
    ],
    expectedValues: { 'C2': 365, 'D2': 11, 'E2': 0 }
  },
  {
    name: 'DAYS',
    category: '04. 日付',
    description: '日数差を計算',
    data: [
      ['開始日', '終了日', '日数'],
      ['2024/1/1', '2024/12/31', '=DAYS(B2,A2)']
    ],
    expectedValues: { 'C2': 365 }
  },
  {
    name: 'EDATE',
    category: '04. 日付',
    description: '月数後の日付',
    data: [
      ['開始日', '月数', '結果'],
      ['2024/1/15', 3, '=EDATE(A2,B2)']
    ],
    expectedValues: { 'C2': 44946 }
  },
  {
    name: 'EOMONTH',
    category: '04. 日付',
    description: '月末日を返す',
    data: [
      ['開始日', '月数', '月末日'],
      ['2024/2/15', 0, '=EOMONTH(A2,B2)']
    ],
    expectedValues: { 'C2': 45352 }
  },
  {
    name: 'NETWORKDAYS',
    category: '04. 日付',
    description: '営業日数を計算',
    data: [
      ['開始日', '終了日', '営業日数'],
      ['2024/1/1', '2024/1/31', '=NETWORKDAYS(A2,B2)'],
      ['2024/1/15', '2024/1/19', '=NETWORKDAYS(A3,B3)']
    ],
    expectedValues: { 'C2': 23, 'C3': 5 }
  },
  {
    name: 'WORKDAY',
    category: '04. 日付',
    description: '営業日後の日付を計算',
    data: [
      ['開始日', '営業日数', '結果日'],
      ['2024/1/1', 10, '=WORKDAY(A2,B2)'],
      ['2024/1/15', 5, '=WORKDAY(A3,B3)']
    ],
    expectedValues: { 'C2': 45303, 'C3': 45317 }
  },
  {
    name: 'DATEVALUE',
    category: '04. 日付',
    description: '日付文字列を日付値に変換',
    data: [
      ['日付文字列', '日付値'],
      ['2024/12/25', '=DATEVALUE(A2)'],
      ['2024-01-15', '=DATEVALUE(A3)']
    ],
    expectedValues: { 'B2': 45651, 'B3': 45306 }
  },
  {
    name: 'TIMEVALUE',
    category: '04. 日付',
    description: '時刻文字列を時刻値に変換',
    data: [
      ['時刻文字列', '時刻値'],
      ['13:30:45', '=TIMEVALUE(A2)'],
      ['9:15 AM', '=TIMEVALUE(A3)']
    ],
    expectedValues: { 'B2': 0.5630208, 'B3': 0.3854167 }
  },
  {
    name: 'WEEKNUM',
    category: '04. 日付',
    description: '週番号を返す',
    data: [
      ['日付', '週番号'],
      ['2024/1/15', '=WEEKNUM(A2)'],
      ['2024/7/1', '=WEEKNUM(A3)']
    ],
    expectedValues: { 'B2': 3, 'B3': 27 }
  },
  {
    name: 'ISOWEEKNUM',
    category: '04. 日付',
    description: 'ISO週番号を返す',
    data: [
      ['日付', 'ISO週番号'],
      ['2024/1/1', '=ISOWEEKNUM(A2)'],
      ['2024/12/31', '=ISOWEEKNUM(A3)']
    ],
    expectedValues: { 'B2': 1, 'B3': 1 }
  },
  {
    name: 'YEARFRAC',
    category: '04. 日付',
    description: '年の端数を計算',
    data: [
      ['開始日', '終了日', '基準', '年の端数'],
      ['2024/1/1', '2024/7/1', 0, '=YEARFRAC(A2,B2,C2)'],
      ['2024/1/1', '2025/1/1', 1, '=YEARFRAC(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 0.5, 'D3': 1 }
  },
  {
    name: 'DAYS360',
    category: '04. 日付',
    description: '360日基準の日数',
    data: [
      ['開始日', '終了日', '方式', '日数'],
      ['2024/1/1', '2024/7/1', 'FALSE', '=DAYS360(A2,B2,C2)'],
      ['2024/1/31', '2024/3/1', 'TRUE', '=DAYS360(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 180 }
  },
  {
    name: 'NETWORKDAYS.INTL',
    category: '04. 日付',
    description: '営業日数（国際版）',
    data: [
      ['開始日', '終了日', '週末', '営業日数'],
      ['2024/1/1', '2024/1/31', 1, '=NETWORKDAYS.INTL(A2,B2,C2)'],
      ['2024/1/1', '2024/1/15', '"0000011"', '=NETWORKDAYS.INTL(A3,B3,C3)']
    ],
    },
  {
    name: 'WORKDAY.INTL',
    category: '04. 日付',
    description: '営業日後の日付（国際版）',
    data: [
      ['開始日', '営業日数', '週末', '結果日'],
      ['2024/1/1', 10, 1, '=WORKDAY.INTL(A2,B2,C2)'],
      ['2024/1/1', 5, '"0000011"', '=WORKDAY.INTL(A3,B3,C3)']
    ],
    }
];
