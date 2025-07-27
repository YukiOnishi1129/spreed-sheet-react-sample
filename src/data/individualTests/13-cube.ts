import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const cubeTests: IndividualFunctionTest[] = [
  {
    name: 'CUBEVALUE',
    category: '13. キューブ',
    description: 'キューブから集計値を返す',
    data: [
      ['接続', 'メンバー1', 'メンバー2', '値'],
      ['Sales', '[Products].[All]', '[Time].[2024]', '=CUBEVALUE(A2,B2,C2)']
    ]
  },
  {
    name: 'CUBEMEMBER',
    category: '13. キューブ',
    description: 'キューブのメンバーを返す',
    data: [
      ['接続', 'メンバー式', 'キャプション', 'メンバー'],
      ['Sales', '[Products].[Bikes]', 'Bikes', '=CUBEMEMBER(A2,B2,C2)']
    ]
  },
  {
    name: 'CUBESET',
    category: '13. キューブ',
    description: 'キューブからセットを定義',
    data: [
      ['接続', 'セット式', 'キャプション', 'セット'],
      ['Sales', '{[Products].[Bikes],[Products].[Cars]}', 'Vehicles', '=CUBESET(A2,B2,C2)']
    ]
  },
  {
    name: 'CUBESETCOUNT',
    category: '13. キューブ',
    description: 'セット内のアイテム数',
    data: [
      ['セット', 'カウント'],
      ['=CUBESET("Sales","{[Products].[Bikes],[Products].[Cars]}")','=CUBESETCOUNT(A2)']
    ]
  },
  {
    name: 'CUBERANKEDMEMBER',
    category: '13. キューブ',
    description: 'セット内のn番目のメンバー',
    data: [
      ['接続', 'セット式', 'ランク', 'メンバー'],
      ['Sales', '[Products].Children', 1, '=CUBERANKEDMEMBER(A2,B2,C2)']
    ]
  },
  {
    name: 'CUBEMEMBERPROPERTY',
    category: '13. キューブ',
    description: 'メンバーのプロパティ値',
    data: [
      ['接続', 'メンバー', 'プロパティ', '値'],
      ['Sales', '[Products].[Bikes]', 'Caption', '=CUBEMEMBERPROPERTY(A2,B2,C2)']
    ]
  },
  {
    name: 'CUBEKPIMEMBER',
    category: '13. キューブ',
    description: 'KPIメンバーを返す',
    data: [
      ['接続', 'KPI名', 'KPIプロパティ', 'メンバー'],
      ['Sales', 'SalesAmount', 'Goal', '=CUBEKPIMEMBER(A2,B2,C2)']
    ]
  }
];
