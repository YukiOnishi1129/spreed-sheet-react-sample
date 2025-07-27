import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const engineeringTests: IndividualFunctionTest[] = [
  {
    name: 'CONVERT',
    category: '11. エンジニアリング',
    description: '単位変換',
    data: [
      ['値', '変換元', '変換先', '結果'],
      [1, 'lbm', 'kg', '=CONVERT(A2,B2,C2)'],
      [100, 'cm', 'm', '=CONVERT(A3,B3,C3)'],
      [32, 'F', 'C', '=CONVERT(A4,B4,C4)']
    ],
    expectedValues: { 'D2': 0.453592, 'D3': 1, 'D4': 0 }
  },
  {
    name: 'BIN2DEC',
    category: '11. エンジニアリング',
    description: '2進数→10進数',
    data: [
      ['2進数', '10進数'],
      ['1100100', '=BIN2DEC(A2)'],
      ['1111111111', '=BIN2DEC(A3)']
    ],
    expectedValues: { 'B2': 100, 'B3': -1 }
  },
  {
    name: 'DEC2BIN',
    category: '11. エンジニアリング',
    description: '10進数→2進数',
    data: [
      ['10進数', '桁数', '2進数'],
      [9, 4, '=DEC2BIN(A2,B2)'],
      [100, '', '=DEC2BIN(A3)']
    ],
    expectedValues: { 'C2': '1001', 'C3': '1100100' }
  },
  {
    name: 'HEX2DEC',
    category: '11. エンジニアリング',
    description: '16進数→10進数',
    data: [
      ['16進数', '10進数'],
      ['FF', '=HEX2DEC(A2)'],
      ['3E8', '=HEX2DEC(A3)']
    ],
    expectedValues: { 'B2': 255, 'B3': 1000 }
  },
  {
    name: 'COMPLEX',
    category: '11. エンジニアリング',
    description: '複素数を作成',
    data: [
      ['実部', '虚部', '複素数'],
      [3, 4, '=COMPLEX(A2,B2)'],
      [0, 1, '=COMPLEX(A3,B3)']
    ],
    expectedValues: { 'C2': '3+4i', 'C3': 'i' }
  },
  {
    name: 'BIN2HEX',
    category: '11. エンジニアリング',
    description: '2進数→16進数',
    data: [
      ['2進数', '桁数', '16進数'],
      ['1100100', 4, '=BIN2HEX(A2,B2)'],
      ['1111', '', '=BIN2HEX(A3)']
    ],
    expectedValues: { 'C2': '0064', 'C3': 'F' }
  },
  {
    name: 'BIN2OCT',
    category: '11. エンジニアリング',
    description: '2進数→8進数',
    data: [
      ['2進数', '桁数', '8進数'],
      ['1100100', 4, '=BIN2OCT(A2,B2)'],
      ['111', '', '=BIN2OCT(A3)']
    ],
    expectedValues: { 'C2': '0144', 'C3': '7' }
  },
  {
    name: 'DEC2HEX',
    category: '11. エンジニアリング',
    description: '10進数→16進数',
    data: [
      ['10進数', '桁数', '16進数'],
      [255, 4, '=DEC2HEX(A2,B2)'],
      [1000, '', '=DEC2HEX(A3)']
    ],
    expectedValues: { 'C2': '00FF', 'C3': '3E8' }
  },
  {
    name: 'DEC2OCT',
    category: '11. エンジニアリング',
    description: '10進数→8進数',
    data: [
      ['10進数', '桁数', '8進数'],
      [100, 4, '=DEC2OCT(A2,B2)'],
      [8, '', '=DEC2OCT(A3)']
    ],
    expectedValues: { 'C2': '0144', 'C3': '10' }
  },
  {
    name: 'HEX2BIN',
    category: '11. エンジニアリング',
    description: '16進数→2進数',
    data: [
      ['16進数', '桁数', '2進数'],
      ['F', 8, '=HEX2BIN(A2,B2)'],
      ['3E8', '', '=HEX2BIN(A3)']
    ],
    expectedValues: { 'C2': '00001111' }
  },
  {
    name: 'HEX2OCT',
    category: '11. エンジニアリング',
    description: '16進数→8進数',
    data: [
      ['16進数', '桁数', '8進数'],
      ['FF', 4, '=HEX2OCT(A2,B2)'],
      ['3E8', '', '=HEX2OCT(A3)']
    ],
    expectedValues: { 'C2': '0377', 'C3': '1750' }
  },
  {
    name: 'OCT2BIN',
    category: '11. エンジニアリング',
    description: '8進数→2進数',
    data: [
      ['8進数', '桁数', '2進数'],
      ['7', 3, '=OCT2BIN(A2,B2)'],
      ['144', '', '=OCT2BIN(A3)']
    ],
    expectedValues: { 'C2': '111', 'C3': '1100100' }
  },
  {
    name: 'OCT2DEC',
    category: '11. エンジニアリング',
    description: '8進数→10進数',
    data: [
      ['8進数', '10進数'],
      ['144', '=OCT2DEC(A2)'],
      ['377', '=OCT2DEC(A3)']
    ],
    expectedValues: { 'B2': 100, 'B3': 255 }
  },
  {
    name: 'OCT2HEX',
    category: '11. エンジニアリング',
    description: '8進数→16進数',
    data: [
      ['8進数', '桁数', '16進数'],
      ['144', 4, '=OCT2HEX(A2,B2)'],
      ['377', '', '=OCT2HEX(A3)']
    ],
    expectedValues: { 'C2': '0064', 'C3': 'FF' }
  },
  {
    name: 'IMABS',
    category: '11. エンジニアリング',
    description: '複素数の絶対値',
    data: [
      ['複素数', '絶対値'],
      ['3+4i', '=IMABS(A2)'],
      ['5-12i', '=IMABS(A3)']
    ],
    expectedValues: { 'B2': 5, 'B3': 13 }
  },
  {
    name: 'IMAGINARY',
    category: '11. エンジニアリング',
    description: '複素数の虚部',
    data: [
      ['複素数', '虚部'],
      ['3+4i', '=IMAGINARY(A2)'],
      ['5-12i', '=IMAGINARY(A3)']
    ],
    expectedValues: { 'B2': 4, 'B3': -12 }
  },
  {
    name: 'IMREAL',
    category: '11. エンジニアリング',
    description: '複素数の実部',
    data: [
      ['複素数', '実部'],
      ['3+4i', '=IMREAL(A2)'],
      ['5-12i', '=IMREAL(A3)']
    ],
    expectedValues: { 'B2': 3, 'B3': 5 }
  },
  {
    name: 'IMCONJUGATE',
    category: '11. エンジニアリング',
    description: '複素共役',
    data: [
      ['複素数', '共役'],
      ['3+4i', '=IMCONJUGATE(A2)'],
      ['5-12i', '=IMCONJUGATE(A3)']
    ],
    expectedValues: { 'B2': '3-4i', 'B3': '5+12i' }
  },
  {
    name: 'IMARGUMENT',
    category: '11. エンジニアリング',
    description: '複素数の偏角',
    data: [
      ['複素数', '偏角'],
      ['1+i', '=IMARGUMENT(A2)'],
      ['-1+i', '=IMARGUMENT(A3)']
    ],
    expectedValues: { 'B2': 0.785398, 'B3': 2.356194 }
  },
  {
    name: 'IMSUM',
    category: '11. エンジニアリング',
    description: '複素数の和',
    data: [
      ['複素数1', '複素数2', '和'],
      ['3+4i', '1+2i', '=IMSUM(A2,B2)'],
      ['5-3i', '2+7i', '=IMSUM(A3,B3)']
    ],
    expectedValues: { 'C2': '4+6i', 'C3': '7+4i' }
  },
  {
    name: 'IMSUB',
    category: '11. エンジニアリング',
    description: '複素数の差',
    data: [
      ['複素数1', '複素数2', '差'],
      ['3+4i', '1+2i', '=IMSUB(A2,B2)'],
      ['5-3i', '2+7i', '=IMSUB(A3,B3)']
    ],
    expectedValues: { 'C2': '2+2i', 'C3': '3-10i' }
  },
  {
    name: 'IMPRODUCT',
    category: '11. エンジニアリング',
    description: '複素数の積',
    data: [
      ['複素数1', '複素数2', '積'],
      ['3+4i', '1+2i', '=IMPRODUCT(A2,B2)'],
      ['2+3i', '4-5i', '=IMPRODUCT(A3,B3)']
    ],
    expectedValues: { 'C2': '-5+10i', 'C3': '23+2i' }
  },
  {
    name: 'IMDIV',
    category: '11. エンジニアリング',
    description: '複素数の商',
    data: [
      ['複素数1', '複素数2', '商'],
      ['1+2i', '3+4i', '=IMDIV(A2,B2)'],
      ['5', '1+i', '=IMDIV(A3,B3)']
    ]
  },
  {
    name: 'IMPOWER',
    category: '11. エンジニアリング',
    description: '複素数のべき乗',
    data: [
      ['複素数', 'べき指数', 'べき乗'],
      ['2+2i', 2, '=IMPOWER(A2,B2)'],
      ['1+i', 3, '=IMPOWER(A3,B3)']
    ],
    expectedValues: { 'C2': '0+8i' }
  },
  {
    name: 'IMLOG10',
    category: '11. エンジニアリング',
    description: '複素数の常用対数',
    data: [
      ['複素数', '常用対数'],
      ['10+10i', '=IMLOG10(A2)'],
      ['100', '=IMLOG10(A3)']
    ]
  },
  {
    name: 'IMLOG2',
    category: '11. エンジニアリング',
    description: '複素数の2を底とする対数',
    data: [
      ['複素数', '対数'],
      ['8+8i', '=IMLOG2(A2)'],
      ['16', '=IMLOG2(A3)']
    ]
  },
  {
    name: 'ERF.PRECISE',
    category: '11. エンジニアリング',
    description: '誤差関数（精密）',
    data: [
      ['値', '誤差関数'],
      [1, '=ERF.PRECISE(A2)'],
      [0.5, '=ERF.PRECISE(A3)']
    ],
    expectedValues: { 'B2': 0.84270079, 'B3': 0.52049988 }
  },
  {
    name: 'ERFC.PRECISE',
    category: '11. エンジニアリング',
    description: '相補誤差関数（精密）',
    data: [
      ['値', '相補誤差関数'],
      [1, '=ERFC.PRECISE(A2)'],
      [0.5, '=ERFC.PRECISE(A3)']
    ],
    expectedValues: { 'B2': 0.15729921, 'B3': 0.47950012 }
  },
  {
    name: 'PHONETIC',
    category: '11. エンジニアリング',
    description: 'ふりがなを抽出',
    data: [
      ['テキスト', 'ふりがな'],
      ['漢字', '=PHONETIC(A2)'],
      ['日本', '=PHONETIC(A3)']
    ]
  },
  {
    name: 'IMSQRT',
    category: '11. エンジニアリング',
    description: '複素数の平方根',
    data: [
      ['複素数', '平方根'],
      ['3+4i', '=IMSQRT(A2)'],
      ['8+6i', '=IMSQRT(A3)']
    ]
  },
  {
    name: 'IMEXP',
    category: '11. エンジニアリング',
    description: '複素数の指数',
    data: [
      ['複素数', '指数'],
      ['1+i', '=IMEXP(A2)'],
      ['2+3i', '=IMEXP(A3)']
    ]
  },
  {
    name: 'IMLN',
    category: '11. エンジニアリング',
    description: '複素数の自然対数',
    data: [
      ['複素数', '自然対数'],
      ['1+i', '=IMLN(A2)'],
      ['2+3i', '=IMLN(A3)']
    ]
  },
  {
    name: 'IMSIN',
    category: '11. エンジニアリング',
    description: '複素数の正弦',
    data: [
      ['複素数', '正弦'],
      ['1+i', '=IMSIN(A2)'],
      ['π/2', '=IMSIN(A3)']
    ]
  },
  {
    name: 'IMCOS',
    category: '11. エンジニアリング',
    description: '複素数の余弦',
    data: [
      ['複素数', '余弦'],
      ['1+i', '=IMCOS(A2)'],
      ['π', '=IMCOS(A3)']
    ]
  },
  {
    name: 'IMTAN',
    category: '11. エンジニアリング',
    description: '複素数の正接',
    data: [
      ['複素数', '正接'],
      ['1+i', '=IMTAN(A2)'],
      ['π/4', '=IMTAN(A3)']
    ]
  },
  {
    name: 'BESSELJ',
    category: '11. エンジニアリング',
    description: 'ベッセル関数Jn(x)',
    data: [
      ['X', 'N', 'ベッセルJ'],
      [1.9, 2, '=BESSELJ(A2,B2)']
    ]
  },
  {
    name: 'BESSELY',
    category: '11. エンジニアリング',
    description: 'ベッセル関数Yn(x)',
    data: [
      ['X', 'N', 'ベッセルY'],
      [2.5, 1, '=BESSELY(A2,B2)']
    ]
  },
  {
    name: 'BESSELI',
    category: '11. エンジニアリング',
    description: '修正ベッセル関数In(x)',
    data: [
      ['X', 'N', '修正ベッセルI'],
      [1.5, 1, '=BESSELI(A2,B2)']
    ]
  },
  {
    name: 'BESSELK',
    category: '11. エンジニアリング',
    description: '修正ベッセル関数Kn(x)',
    data: [
      ['X', 'N', '修正ベッセルK'],
      [1.5, 1, '=BESSELK(A2,B2)']
    ]
  },
  {
    name: 'BITAND',
    category: '11. エンジニアリング',
    description: 'ビット単位AND',
    data: [
      ['値1', '値2', 'AND結果'],
      [5, 3, '=BITAND(A2,B2)'],
      [23, 10, '=BITAND(A3,B3)']
    ],
    expectedValues: { 'C2': 1, 'C3': 2 }
  },
  {
    name: 'BITOR',
    category: '11. エンジニアリング',
    description: 'ビット単位OR',
    data: [
      ['値1', '値2', 'OR結果'],
      [5, 3, '=BITOR(A2,B2)'],
      [23, 10, '=BITOR(A3,B3)']
    ],
    expectedValues: { 'C2': 7, 'C3': 31 }
  },
  {
    name: 'BITXOR',
    category: '11. エンジニアリング',
    description: 'ビット単位XOR',
    data: [
      ['値1', '値2', 'XOR結果'],
      [5, 3, '=BITXOR(A2,B2)'],
      [23, 10, '=BITXOR(A3,B3)']
    ],
    expectedValues: { 'C2': 6, 'C3': 29 }
  },
  {
    name: 'BITLSHIFT',
    category: '11. エンジニアリング',
    description: 'ビット左シフト',
    data: [
      ['値', 'シフト量', '結果'],
      [4, 2, '=BITLSHIFT(A2,B2)'],
      [3, 3, '=BITLSHIFT(A3,B3)']
    ],
    expectedValues: { 'C2': 16, 'C3': 24 }
  },
  {
    name: 'BITRSHIFT',
    category: '11. エンジニアリング',
    description: 'ビット右シフト',
    data: [
      ['値', 'シフト量', '結果'],
      [13, 2, '=BITRSHIFT(A2,B2)'],
      [64, 3, '=BITRSHIFT(A3,B3)']
    ],
    expectedValues: { 'C2': 3, 'C3': 8 }
  },
  {
    name: 'DELTA',
    category: '11. エンジニアリング',
    description: 'クロネッカーのデルタ関数',
    data: [
      ['値1', '値2', 'デルタ'],
      [5, 5, '=DELTA(A2,B2)'],
      [5, 4, '=DELTA(A3,B3)']
    ],
    expectedValues: { 'C2': 1, 'C3': 0 }
  },
  {
    name: 'GESTEP',
    category: '11. エンジニアリング',
    description: 'ステップ関数',
    data: [
      ['値', 'ステップ', '結果'],
      [5, 4, '=GESTEP(A2,B2)'],
      [3, 5, '=GESTEP(A3,B3)']
    ],
    expectedValues: { 'C2': 1, 'C3': 0 }
  }
];
