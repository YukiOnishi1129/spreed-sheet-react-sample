import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const matrixTests: IndividualFunctionTest[] = [
  {
    name: 'MMULT',
    category: '08. 行列',
    description: '行列の積',
    data: [
      ['行列A', '', '行列B', '', '積'],
      [1, 2, 5, 6, '=MMULT(A2:B3,C2:D3)'],
      [3, 4, 7, 8, '']
    ]
  },
  {
    name: 'MDETERM',
    category: '08. 行列',
    description: '行列式',
    data: [
      ['行列', '', '行列式'],
      [3, 2, '=MDETERM(A2:B3)'],
      [1, 4, '']
    ],
    expectedValues: { 'C2': 10 }
  },
  {
    name: 'MINVERSE',
    category: '08. 行列',
    description: '逆行列',
    data: [
      ['行列', '', '逆行列'],
      [3, 2, '=MINVERSE(A2:B3)'],
      [1, 4, '']
    ]
  },
  {
    name: 'MUNIT',
    category: '08. 行列',
    description: '単位行列',
    data: [
      ['サイズ', '単位行列'],
      [3, '=MUNIT(A2)']
    ]
  }
];
