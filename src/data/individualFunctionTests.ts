export interface IndividualFunctionTest {
  name: string;
  category: string;
  description: string;
  data: (string | number | boolean | null)[][];
  expectedValues?: { [key: string]: string | number | boolean };
}

// 数学・三角関数の個別テスト
export const mathFunctionTests: IndividualFunctionTest[] = [
  // 基本的な数学関数
  {
    name: 'SUM',
    category: '数学',
    description: '数値の合計を計算',
    data: [
      ['値1', '値2', '値3', '値4', '結果'],
      [10, 20, 30, 40, '=SUM(A2:D2)']
    ],
    expectedValues: { 'E2': 100 }
  },
  {
    name: 'SUMIF',
    category: '数学',
    description: '条件に一致するセルの合計',
    data: [
      ['値', '条件', '結果'],
      [10, '>0', ''],
      [20, '>0', ''],
      [-5, '<0', ''],
      [30, '>0', '=SUMIF(B2:B5,">0",A2:A5)']
    ],
    expectedValues: { 'C5': 60 }
  },
  {
    name: 'SUMIFS',
    category: '数学',
    description: '複数条件に一致するセルの合計',
    data: [
      ['値', '条件1', '条件2', '結果'],
      [10, 'A', 1, ''],
      [20, 'B', 2, ''],
      [30, 'A', 2, ''],
      [40, 'A', 1, '=SUMIFS(A2:A5,B2:B5,"A",C2:C5,1)']
    ],
    expectedValues: { 'D5': 50 }
  },
  {
    name: 'PRODUCT',
    category: '数学',
    description: '数値の積を計算',
    data: [
      ['値1', '値2', '値3', '結果'],
      [2, 3, 4, '=PRODUCT(A2:C2)']
    ],
    expectedValues: { 'D2': 24 }
  },
  {
    name: 'SQRT',
    category: '数学',
    description: '平方根を計算',
    data: [
      ['値', '結果'],
      [16, '=SQRT(A2)'],
      [25, '=SQRT(A3)'],
      [100, '=SQRT(A4)']
    ],
    expectedValues: { 'B2': 4, 'B3': 5, 'B4': 10 }
  },
  {
    name: 'POWER',
    category: '数学',
    description: 'べき乗を計算',
    data: [
      ['底', '指数', '結果'],
      [2, 3, '=POWER(A2,B2)'],
      [5, 2, '=POWER(A3,B3)']
    ],
    expectedValues: { 'C2': 8, 'C3': 25 }
  },
  {
    name: 'ABS',
    category: '数学',
    description: '絶対値を返す',
    data: [
      ['値', '結果'],
      [-10, '=ABS(A2)'],
      [10, '=ABS(A3)'],
      [0, '=ABS(A4)']
    ],
    expectedValues: { 'B2': 10, 'B3': 10, 'B4': 0 }
  },
  {
    name: 'ROUND',
    category: '数学',
    description: '指定桁数で四捨五入',
    data: [
      ['値', '桁数', '結果'],
      [3.14159, 2, '=ROUND(A2,B2)'],
      [123.456, -1, '=ROUND(A3,B3)'],
      [2.5, 0, '=ROUND(A4,B4)']
    ],
    expectedValues: { 'C2': 3.14, 'C3': 120, 'C4': 3 }
  },
  {
    name: 'ROUNDUP',
    category: '数学',
    description: '指定桁数で切り上げ',
    data: [
      ['値', '桁数', '結果'],
      [3.14159, 2, '=ROUNDUP(A2,B2)'],
      [123.456, -1, '=ROUNDUP(A3,B3)'],
      [2.1, 0, '=ROUNDUP(A4,B4)']
    ],
    expectedValues: { 'C2': 3.15, 'C3': 130, 'C4': 3 }
  },
  {
    name: 'ROUNDDOWN',
    category: '数学',
    description: '指定桁数で切り下げ',
    data: [
      ['値', '桁数', '結果'],
      [3.14159, 2, '=ROUNDDOWN(A2,B2)'],
      [123.456, -1, '=ROUNDDOWN(A3,B3)'],
      [2.9, 0, '=ROUNDDOWN(A4,B4)']
    ],
    expectedValues: { 'C2': 3.14, 'C3': 120, 'C4': 2 }
  },
  {
    name: 'CEILING',
    category: '数学',
    description: '基準値の倍数に切り上げ',
    data: [
      ['値', '基準値', '結果'],
      [2.3, 1, '=CEILING(A2,B2)'],
      [4.7, 2, '=CEILING(A3,B3)'],
      [11, 5, '=CEILING(A4,B4)']
    ],
    expectedValues: { 'C2': 3, 'C3': 6, 'C4': 15 }
  },
  {
    name: 'FLOOR',
    category: '数学',
    description: '基準値の倍数に切り下げ',
    data: [
      ['値', '基準値', '結果'],
      [2.3, 1, '=FLOOR(A2,B2)'],
      [4.7, 2, '=FLOOR(A3,B3)'],
      [11, 5, '=FLOOR(A4,B4)']
    ],
    expectedValues: { 'C2': 2, 'C3': 4, 'C4': 10 }
  },
  {
    name: 'MOD',
    category: '数学',
    description: '剰余を返す',
    data: [
      ['被除数', '除数', '結果'],
      [10, 3, '=MOD(A2,B2)'],
      [15, 4, '=MOD(A3,B3)'],
      [20, 5, '=MOD(A4,B4)']
    ],
    expectedValues: { 'C2': 1, 'C3': 3, 'C4': 0 }
  },
  {
    name: 'INT',
    category: '数学',
    description: '整数部分を返す',
    data: [
      ['値', '結果'],
      [3.14, '=INT(A2)'],
      [-3.14, '=INT(A3)'],
      [5.999, '=INT(A4)']
    ],
    expectedValues: { 'B2': 3, 'B3': -4, 'B4': 5 }
  },
  {
    name: 'TRUNC',
    category: '数学',
    description: '小数部分を切り捨て',
    data: [
      ['値', '桁数', '結果'],
      [3.14159, 2, '=TRUNC(A2,B2)'],
      [-3.14159, 2, '=TRUNC(A3,B3)'],
      [123.456, 0, '=TRUNC(A4,B4)']
    ],
    expectedValues: { 'C2': 3.14, 'C3': -3.14, 'C4': 123 }
  },
  {
    name: 'SIGN',
    category: '数学',
    description: '数値の符号を返す',
    data: [
      ['値', '結果'],
      [10, '=SIGN(A2)'],
      [-10, '=SIGN(A3)'],
      [0, '=SIGN(A4)']
    ],
    expectedValues: { 'B2': 1, 'B3': -1, 'B4': 0 }
  },
  {
    name: 'RAND',
    category: '数学',
    description: '0以上1未満の乱数',
    data: [
      ['乱数1', '乱数2', '乱数3'],
      ['=RAND()', '=RAND()', '=RAND()']
    ]
  },
  {
    name: 'RANDBETWEEN',
    category: '数学',
    description: '指定範囲の整数乱数',
    data: [
      ['最小', '最大', '結果'],
      [1, 10, '=RANDBETWEEN(A2,B2)'],
      [1, 100, '=RANDBETWEEN(A3,B3)']
    ]
  },
  {
    name: 'EXP',
    category: '数学',
    description: 'eのべき乗を計算',
    data: [
      ['値', '結果'],
      [0, '=EXP(A2)'],
      [1, '=EXP(A3)'],
      [2, '=EXP(A4)']
    ],
    expectedValues: { 'B2': 1, 'B3': 2.718281828, 'B4': 7.389056099 }
  },
  {
    name: 'LN',
    category: '数学',
    description: '自然対数を計算',
    data: [
      ['値', '結果'],
      [1, '=LN(A2)'],
      [2.718281828, '=LN(A3)'],
      [10, '=LN(A4)']
    ],
    expectedValues: { 'B2': 0, 'B3': 1, 'B4': 2.302585093 }
  },
  {
    name: 'LOG',
    category: '数学',
    description: '対数を計算',
    data: [
      ['値', '底', '結果'],
      [100, 10, '=LOG(A2,B2)'],
      [8, 2, '=LOG(A3,B3)'],
      [27, 3, '=LOG(A4,B4)']
    ],
    expectedValues: { 'C2': 2, 'C3': 3, 'C4': 3 }
  },
  {
    name: 'LOG10',
    category: '数学',
    description: '常用対数を計算',
    data: [
      ['値', '結果'],
      [10, '=LOG10(A2)'],
      [100, '=LOG10(A3)'],
      [1000, '=LOG10(A4)']
    ],
    expectedValues: { 'B2': 1, 'B3': 2, 'B4': 3 }
  },
  {
    name: 'FACT',
    category: '数学',
    description: '階乗を計算',
    data: [
      ['値', '結果'],
      [0, '=FACT(A2)'],
      [3, '=FACT(A3)'],
      [5, '=FACT(A4)']
    ],
    expectedValues: { 'B2': 1, 'B3': 6, 'B4': 120 }
  },
  {
    name: 'COMBIN',
    category: '数学',
    description: '組み合わせ数を計算',
    data: [
      ['総数', '選択数', '結果'],
      [5, 2, '=COMBIN(A2,B2)'],
      [10, 3, '=COMBIN(A3,B3)'],
      [6, 4, '=COMBIN(A4,B4)']
    ],
    expectedValues: { 'C2': 10, 'C3': 120, 'C4': 15 }
  },
  {
    name: 'PERMUT',
    category: '数学',
    description: '順列数を計算',
    data: [
      ['総数', '選択数', '結果'],
      [5, 2, '=PERMUT(A2,B2)'],
      [10, 3, '=PERMUT(A3,B3)'],
      [6, 4, '=PERMUT(A4,B4)']
    ],
    expectedValues: { 'C2': 20, 'C3': 720, 'C4': 360 }
  },
  {
    name: 'GCD',
    category: '数学',
    description: '最大公約数を計算',
    data: [
      ['値1', '値2', '値3', '結果'],
      [12, 18, 24, '=GCD(A2:C2)'],
      [10, 15, 20, '=GCD(A3:C3)']
    ],
    expectedValues: { 'D2': 6, 'D3': 5 }
  },
  {
    name: 'LCM',
    category: '数学',
    description: '最小公倍数を計算',
    data: [
      ['値1', '値2', '値3', '結果'],
      [4, 6, 8, '=LCM(A2:C2)'],
      [3, 5, 7, '=LCM(A3:C3)']
    ],
    expectedValues: { 'D2': 24, 'D3': 105 }
  },
  {
    name: 'QUOTIENT',
    category: '数学',
    description: '商の整数部分を返す',
    data: [
      ['被除数', '除数', '結果'],
      [10, 3, '=QUOTIENT(A2,B2)'],
      [20, 6, '=QUOTIENT(A3,B3)'],
      [-10, 3, '=QUOTIENT(A4,B4)']
    ],
    expectedValues: { 'C2': 3, 'C3': 3, 'C4': -3 }
  },
  {
    name: 'MROUND',
    category: '数学',
    description: '倍数に丸める',
    data: [
      ['値', '倍数', '結果'],
      [10, 3, '=MROUND(A2,B2)'],
      [14, 5, '=MROUND(A3,B3)'],
      [17, 4, '=MROUND(A4,B4)']
    ],
    expectedValues: { 'C2': 9, 'C3': 15, 'C4': 16 }
  },
  {
    name: 'SUMSQ',
    category: '数学',
    description: '平方和を計算',
    data: [
      ['値1', '値2', '値3', '結果'],
      [3, 4, 0, '=SUMSQ(A2:C2)'],
      [1, 2, 3, '=SUMSQ(A3:C3)']
    ],
    expectedValues: { 'D2': 25, 'D3': 14 }
  },
  {
    name: 'SUMPRODUCT',
    category: '数学',
    description: '配列の積の和を計算',
    data: [
      ['配列1', '', '', '配列2', '', '', '結果'],
      [2, 3, 4, 5, 6, 7, '=SUMPRODUCT(A2:C2,D2:F2)']
    ],
    expectedValues: { 'G2': 56 }
  },
  // 三角関数
  {
    name: 'SIN',
    category: '三角',
    description: '正弦を計算',
    data: [
      ['角度(ラジアン)', '結果'],
      [0, '=SIN(A2)'],
      ['=PI()/2', '=SIN(A3)'],
      ['=PI()', '=SIN(A4)']
    ],
    expectedValues: { 'B2': 0, 'B3': 1, 'B4': 0 }
  },
  {
    name: 'COS',
    category: '三角',
    description: '余弦を計算',
    data: [
      ['角度(ラジアン)', '結果'],
      [0, '=COS(A2)'],
      ['=PI()/2', '=COS(A3)'],
      ['=PI()', '=COS(A4)']
    ],
    expectedValues: { 'B2': 1, 'B3': 0, 'B4': -1 }
  },
  {
    name: 'TAN',
    category: '三角',
    description: '正接を計算',
    data: [
      ['角度(ラジアン)', '結果'],
      [0, '=TAN(A2)'],
      ['=PI()/4', '=TAN(A3)']
    ],
    expectedValues: { 'B2': 0, 'B3': 1 }
  },
  {
    name: 'ASIN',
    category: '三角',
    description: '逆正弦を計算',
    data: [
      ['値', '結果'],
      [0, '=ASIN(A2)'],
      [0.5, '=ASIN(A3)'],
      [1, '=ASIN(A4)']
    ],
    expectedValues: { 'B2': 0, 'B3': 0.523598776, 'B4': 1.570796327 }
  },
  {
    name: 'ACOS',
    category: '三角',
    description: '逆余弦を計算',
    data: [
      ['値', '結果'],
      [1, '=ACOS(A2)'],
      [0.5, '=ACOS(A3)'],
      [0, '=ACOS(A4)']
    ],
    expectedValues: { 'B2': 0, 'B3': 1.047197551, 'B4': 1.570796327 }
  },
  {
    name: 'ATAN',
    category: '三角',
    description: '逆正接を計算',
    data: [
      ['値', '結果'],
      [0, '=ATAN(A2)'],
      [1, '=ATAN(A3)'],
      [-1, '=ATAN(A4)']
    ],
    expectedValues: { 'B2': 0, 'B3': 0.785398163, 'B4': -0.785398163 }
  },
  {
    name: 'ATAN2',
    category: '三角',
    description: 'x,y座標から角度を計算',
    data: [
      ['x座標', 'y座標', '結果'],
      [1, 1, '=ATAN2(A2,B2)'],
      [1, 0, '=ATAN2(A3,B3)'],
      [0, 1, '=ATAN2(A4,B4)']
    ],
    expectedValues: { 'C2': 0.785398163, 'C3': 0, 'C4': 1.570796327 }
  },
  {
    name: 'DEGREES',
    category: '三角',
    description: 'ラジアンを度に変換',
    data: [
      ['ラジアン', '度'],
      ['=PI()', '=DEGREES(A2)'],
      ['=PI()/2', '=DEGREES(A3)'],
      ['=PI()/4', '=DEGREES(A4)']
    ],
    expectedValues: { 'B2': 180, 'B3': 90, 'B4': 45 }
  },
  {
    name: 'RADIANS',
    category: '三角',
    description: '度をラジアンに変換',
    data: [
      ['度', 'ラジアン'],
      [180, '=RADIANS(A2)'],
      [90, '=RADIANS(A3)'],
      [45, '=RADIANS(A4)']
    ],
    expectedValues: { 'B2': 3.141592654, 'B3': 1.570796327, 'B4': 0.785398163 }
  },
  {
    name: 'PI',
    category: '三角',
    description: '円周率を返す',
    data: [
      ['円周率'],
      ['=PI()']
    ],
    expectedValues: { 'A2': 3.141592654 }
  },
  // 追加の数学関数
  {
    name: 'LOG10',
    category: '数学',
    description: '常用対数を計算',
    data: [
      ['値', '常用対数'],
      [100, '=LOG10(A2)'],
      [1000, '=LOG10(A3)']
    ],
    expectedValues: { 'B2': 2, 'B3': 3 }
  },
  {
    name: 'CEILING.MATH',
    category: '数学',
    description: '数学的な切り上げ',
    data: [
      ['値', '基準値', '結果'],
      [4.3, 1, '=CEILING.MATH(A2,B2)'],
      [-4.3, 1, '=CEILING.MATH(A3,B3)']
    ],
    expectedValues: { 'C2': 5, 'C3': -4 }
  },
  {
    name: 'FLOOR.MATH',
    category: '数学',
    description: '数学的な切り下げ',
    data: [
      ['値', '基準値', '結果'],
      [4.7, 1, '=FLOOR.MATH(A2,B2)'],
      [-4.7, 1, '=FLOOR.MATH(A3,B3)']
    ],
    expectedValues: { 'C2': 4, 'C3': -5 }
  },
  {
    name: 'EVEN',
    category: '数学',
    description: '最も近い偶数に切り上げ',
    data: [
      ['値', '偶数'],
      [3, '=EVEN(A2)'],
      [4, '=EVEN(A3)'],
      [5.1, '=EVEN(A4)']
    ],
    expectedValues: { 'B2': 4, 'B3': 4, 'B4': 6 }
  },
  {
    name: 'ODD',
    category: '数学',
    description: '最も近い奇数に切り上げ',
    data: [
      ['値', '奇数'],
      [2, '=ODD(A2)'],
      [3, '=ODD(A3)'],
      [4.1, '=ODD(A4)']
    ],
    expectedValues: { 'B2': 3, 'B3': 3, 'B4': 5 }
  },
  {
    name: 'SINH',
    category: '三角',
    description: '双曲線正弦を計算',
    data: [
      ['値', '双曲線正弦'],
      [0, '=SINH(A2)'],
      [1, '=SINH(A3)']
    ],
    expectedValues: { 'B2': 0, 'B3': 1.175201 }
  },
  {
    name: 'COSH',
    category: '三角',
    description: '双曲線余弦を計算',
    data: [
      ['値', '双曲線余弦'],
      [0, '=COSH(A2)'],
      [1, '=COSH(A3)']
    ],
    expectedValues: { 'B2': 1, 'B3': 1.543081 }
  },
  {
    name: 'TANH',
    category: '三角',
    description: '双曲線正接を計算',
    data: [
      ['値', '双曲線正接'],
      [0, '=TANH(A2)'],
      [0.5, '=TANH(A3)']
    ],
    expectedValues: { 'B2': 0, 'B3': 0.462117 }
  },
  {
    name: 'ATAN2',
    category: '三角',
    description: 'x,y座標から角度を計算',
    data: [
      ['X座標', 'Y座標', '角度(ラジアン)'],
      [1, 1, '=ATAN2(A2,B2)'],
      [0, 1, '=ATAN2(A3,B3)']
    ],
    expectedValues: { 'C2': 0.785398, 'C3': 1.570796 }
  }
];

// 統計関数の個別テスト
export const statisticalFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'AVERAGE',
    category: '統計',
    description: '平均値を計算',
    data: [
      ['値1', '値2', '値3', '値4', '平均'],
      [10, 20, 30, 40, '=AVERAGE(A2:D2)']
    ],
    expectedValues: { 'E2': 25 }
  },
  {
    name: 'AVERAGEIF',
    category: '統計',
    description: '条件付き平均値',
    data: [
      ['値', '条件', '平均'],
      [10, 'A', ''],
      [20, 'B', ''],
      [30, 'A', ''],
      [40, 'A', '=AVERAGEIF(B2:B5,"A",A2:A5)']
    ],
    expectedValues: { 'C5': 26.666667 }
  },
  {
    name: 'AVERAGEIFS',
    category: '統計',
    description: '複数条件付き平均値',
    data: [
      ['値', '条件1', '条件2', '平均'],
      [10, 'A', 1, ''],
      [20, 'B', 2, ''],
      [30, 'A', 1, ''],
      [40, 'A', 2, '=AVERAGEIFS(A2:A5,B2:B5,"A",C2:C5,1)']
    ],
    expectedValues: { 'D5': 20 }
  },
  {
    name: 'COUNT',
    category: '統計',
    description: '数値の個数を数える',
    data: [
      ['値1', '値2', '値3', '値4', '個数'],
      [10, 20, 'テキスト', 30, '=COUNT(A2:D2)']
    ],
    expectedValues: { 'E2': 3 }
  },
  {
    name: 'COUNTA',
    category: '統計',
    description: '空白以外のセル数',
    data: [
      ['値1', '値2', '値3', '値4', '個数'],
      [10, 20, 'テキスト', '', '=COUNTA(A2:D2)']
    ],
    expectedValues: { 'E2': 3 }
  },
  {
    name: 'COUNTBLANK',
    category: '統計',
    description: '空白セルの個数',
    data: [
      ['値1', '値2', '値3', '値4', '空白数'],
      [10, '', 'テキスト', '', '=COUNTBLANK(A2:D2)']
    ],
    expectedValues: { 'E2': 2 }
  },
  {
    name: 'COUNTIF',
    category: '統計',
    description: '条件に一致するセル数',
    data: [
      ['値', '個数'],
      [10, ''],
      [20, ''],
      [10, ''],
      [30, '=COUNTIF(A2:A5,10)']
    ],
    expectedValues: { 'B5': 2 }
  },
  {
    name: 'COUNTIFS',
    category: '統計',
    description: '複数条件に一致するセル数',
    data: [
      ['値1', '値2', '個数'],
      [10, 'A', ''],
      [20, 'B', ''],
      [10, 'A', ''],
      [30, 'A', '=COUNTIFS(A2:A5,10,B2:B5,"A")']
    ],
    expectedValues: { 'C5': 2 }
  },
  {
    name: 'MAX',
    category: '統計',
    description: '最大値を返す',
    data: [
      ['値1', '値2', '値3', '値4', '最大値'],
      [10, 30, 20, 40, '=MAX(A2:D2)']
    ],
    expectedValues: { 'E2': 40 }
  },
  {
    name: 'MIN',
    category: '統計',
    description: '最小値を返す',
    data: [
      ['値1', '値2', '値3', '値4', '最小値'],
      [10, 30, 20, 40, '=MIN(A2:D2)']
    ],
    expectedValues: { 'E2': 10 }
  },
  {
    name: 'MAXIFS',
    category: '統計',
    description: '条件付き最大値',
    data: [
      ['値', '条件', '最大値'],
      [10, 'A', ''],
      [20, 'B', ''],
      [30, 'A', ''],
      [40, 'A', '=MAXIFS(A2:A5,B2:B5,"A")']
    ],
    expectedValues: { 'C5': 40 }
  },
  {
    name: 'MINIFS',
    category: '統計',
    description: '条件付き最小値',
    data: [
      ['値', '条件', '最小値'],
      [10, 'A', ''],
      [20, 'B', ''],
      [30, 'A', ''],
      [40, 'A', '=MINIFS(A2:A5,B2:B5,"A")']
    ],
    expectedValues: { 'C5': 10 }
  },
  {
    name: 'MEDIAN',
    category: '統計',
    description: '中央値を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '中央値'],
      [1, 3, 3, 6, 7, '=MEDIAN(A2:E2)']
    ],
    expectedValues: { 'F2': 3 }
  },
  {
    name: 'MODE',
    category: '統計',
    description: '最頻値を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '最頻値'],
      [1, 2, 3, 3, 4, '=MODE(A2:E2)']
    ],
    expectedValues: { 'F2': 3 }
  },
  {
    name: 'STDEV',
    category: '統計',
    description: '標準偏差（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '標準偏差'],
      [10, 20, 30, 40, '=STDEV(A2:D2)']
    ],
    expectedValues: { 'E2': 12.90994449 }
  },
  {
    name: 'STDEV.S',
    category: '統計',
    description: '標準偏差（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '標準偏差'],
      [10, 20, 30, 40, '=STDEV.S(A2:D2)']
    ],
    expectedValues: { 'E2': 12.90994449 }
  },
  {
    name: 'STDEV.P',
    category: '統計',
    description: '標準偏差（母集団）',
    data: [
      ['値1', '値2', '値3', '値4', '標準偏差'],
      [10, 20, 30, 40, '=STDEV.P(A2:D2)']
    ],
    expectedValues: { 'E2': 11.18033989 }
  },
  {
    name: 'VAR',
    category: '統計',
    description: '分散（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '分散'],
      [10, 20, 30, 40, '=VAR(A2:D2)']
    ],
    expectedValues: { 'E2': 166.6666667 }
  },
  {
    name: 'VAR.S',
    category: '統計',
    description: '分散（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '分散'],
      [10, 20, 30, 40, '=VAR.S(A2:D2)']
    ],
    expectedValues: { 'E2': 166.6666667 }
  },
  {
    name: 'VAR.P',
    category: '統計',
    description: '分散（母集団）',
    data: [
      ['値1', '値2', '値3', '値4', '分散'],
      [10, 20, 30, 40, '=VAR.P(A2:D2)']
    ],
    expectedValues: { 'E2': 125 }
  },
  {
    name: 'CORREL',
    category: '統計',
    description: '相関係数を計算',
    data: [
      ['X', 'Y', '', '相関係数'],
      [1, 2, '', ''],
      [2, 4, '', ''],
      [3, 6, '', ''],
      [4, 8, '', '=CORREL(A2:A5,B2:B5)']
    ],
    expectedValues: { 'D5': 1 }
  },
  {
    name: 'LARGE',
    category: '統計',
    description: 'k番目に大きい値',
    data: [
      ['値', '', '', '', 'k', '結果'],
      [10, 30, 20, 40, 2, '=LARGE(A2:D2,E2)']
    ],
    expectedValues: { 'F2': 30 }
  },
  {
    name: 'SMALL',
    category: '統計',
    description: 'k番目に小さい値',
    data: [
      ['値', '', '', '', 'k', '結果'],
      [10, 30, 20, 40, 2, '=SMALL(A2:D2,E2)']
    ],
    expectedValues: { 'F2': 20 }
  },
  {
    name: 'RANK',
    category: '統計',
    description: '順位を返す',
    data: [
      ['値', '順位'],
      [10, '=RANK(A2,$A$2:$A$5)'],
      [30, '=RANK(A3,$A$2:$A$5)'],
      [20, '=RANK(A4,$A$2:$A$5)'],
      [40, '=RANK(A5,$A$2:$A$5)']
    ],
    expectedValues: { 'B2': 4, 'B3': 2, 'B4': 3, 'B5': 1 }
  },
  {
    name: 'PERCENTILE',
    category: '統計',
    description: 'パーセンタイル値',
    data: [
      ['値', '', '', '', 'k', '結果'],
      [1, 2, 3, 4, 0.5, '=PERCENTILE(A2:D2,E2)']
    ],
    expectedValues: { 'F2': 2.5 }
  },
  {
    name: 'QUARTILE',
    category: '統計',
    description: '四分位数',
    data: [
      ['値', '', '', '', '四分位', '結果'],
      [1, 2, 3, 4, 1, '=QUARTILE(A2:D2,E2)']
    ],
    expectedValues: { 'F2': 1.75 }
  },
  // 追加の統計関数
  {
    name: 'SKEW',
    category: '統計',
    description: '歪度を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '歪度'],
      [3, 4, 5, 2, 1, '=SKEW(A2:E2)']
    ]
  },
  {
    name: 'KURT',
    category: '統計',
    description: '尖度を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '尖度'],
      [3, 4, 5, 2, 1, '=KURT(A2:E2)']
    ]
  },
  {
    name: 'STANDARDIZE',
    category: '統計',
    description: '標準化する',
    data: [
      ['値', '平均', '標準偏差', '標準化値'],
      [42, 40, 1.5, '=STANDARDIZE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 1.333333 }
  },
  {
    name: 'DEVSQ',
    category: '統計',
    description: '偏差平方和を計算',
    data: [
      ['値1', '値2', '値3', '値4', '偏差平方和'],
      [4, 5, 8, 7, '=DEVSQ(A2:D2)']
    ],
    expectedValues: { 'E2': 10 }
  },
  {
    name: 'GEOMEAN',
    category: '統計',
    description: '幾何平均を計算',
    data: [
      ['値1', '値2', '値3', '値4', '幾何平均'],
      [4, 5, 8, 7, '=GEOMEAN(A2:D2)']
    ]
  },
  {
    name: 'HARMEAN',
    category: '統計',
    description: '調和平均を計算',
    data: [
      ['値1', '値2', '値3', '値4', '調和平均'],
      [4, 5, 8, 7, '=HARMEAN(A2:D2)']
    ]
  },
  {
    name: 'TRIMMEAN',
    category: '統計',
    description: 'トリム平均を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '値6', '除外率', 'トリム平均'],
      [1, 2, 3, 4, 5, 100, 0.2, '=TRIMMEAN(A2:F2,G2)']
    ]
  }
];

// 文字列関数の個別テスト
export const textFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'CONCATENATE',
    category: '文字列',
    description: '文字列を結合',
    data: [
      ['文字列1', '文字列2', '文字列3', '結果'],
      ['Hello', ' ', 'World', '=CONCATENATE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 'Hello World' }
  },
  {
    name: 'CONCAT',
    category: '文字列',
    description: '文字列を結合（新版）',
    data: [
      ['文字列1', '文字列2', '結果'],
      ['ABC', 'DEF', '=CONCAT(A2,B2)']
    ],
    expectedValues: { 'C2': 'ABCDEF' }
  },
  {
    name: 'TEXTJOIN',
    category: '文字列',
    description: '区切り文字で結合',
    data: [
      ['値1', '値2', '値3', '結果'],
      ['A', 'B', 'C', '=TEXTJOIN("-",TRUE,A2:C2)']
    ],
    expectedValues: { 'D2': 'A-B-C' }
  },
  {
    name: 'LEFT',
    category: '文字列',
    description: '左から文字を抽出',
    data: [
      ['文字列', '文字数', '結果'],
      ['Hello World', 5, '=LEFT(A2,B2)']
    ],
    expectedValues: { 'C2': 'Hello' }
  },
  {
    name: 'RIGHT',
    category: '文字列',
    description: '右から文字を抽出',
    data: [
      ['文字列', '文字数', '結果'],
      ['Hello World', 5, '=RIGHT(A2,B2)']
    ],
    expectedValues: { 'C2': 'World' }
  },
  {
    name: 'MID',
    category: '文字列',
    description: '中間の文字を抽出',
    data: [
      ['文字列', '開始位置', '文字数', '結果'],
      ['Hello World', 3, 3, '=MID(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 'llo' }
  },
  {
    name: 'LEN',
    category: '文字列',
    description: '文字数を返す',
    data: [
      ['文字列', '文字数'],
      ['Hello', '=LEN(A2)'],
      ['こんにちは', '=LEN(A3)']
    ],
    expectedValues: { 'B2': 5, 'B3': 5 }
  },
  {
    name: 'FIND',
    category: '文字列',
    description: '文字位置を検索（大小区別）',
    data: [
      ['検索文字列', '対象文字列', '位置'],
      ['o', 'Hello World', '=FIND(A2,B2)']
    ],
    expectedValues: { 'C2': 5 }
  },
  {
    name: 'SEARCH',
    category: '文字列',
    description: '文字位置を検索（大小区別なし）',
    data: [
      ['検索文字列', '対象文字列', '位置'],
      ['world', 'Hello World', '=SEARCH(A2,B2)']
    ],
    expectedValues: { 'C2': 7 }
  },
  {
    name: 'REPLACE',
    category: '文字列',
    description: '文字を置換（位置指定）',
    data: [
      ['文字列', '開始位置', '文字数', '新文字列', '結果'],
      ['Hello World', 7, 5, 'Excel', '=REPLACE(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 'Hello Excel' }
  },
  {
    name: 'SUBSTITUTE',
    category: '文字列',
    description: '文字を置換（文字指定）',
    data: [
      ['文字列', '検索文字列', '置換文字列', '結果'],
      ['Hello World', 'World', 'Excel', '=SUBSTITUTE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 'Hello Excel' }
  },
  {
    name: 'UPPER',
    category: '文字列',
    description: '大文字に変換',
    data: [
      ['文字列', '結果'],
      ['Hello World', '=UPPER(A2)']
    ],
    expectedValues: { 'B2': 'HELLO WORLD' }
  },
  {
    name: 'LOWER',
    category: '文字列',
    description: '小文字に変換',
    data: [
      ['文字列', '結果'],
      ['Hello World', '=LOWER(A2)']
    ],
    expectedValues: { 'B2': 'hello world' }
  },
  {
    name: 'PROPER',
    category: '文字列',
    description: '先頭を大文字に変換',
    data: [
      ['文字列', '結果'],
      ['hello world', '=PROPER(A2)']
    ],
    expectedValues: { 'B2': 'Hello World' }
  },
  {
    name: 'TRIM',
    category: '文字列',
    description: '余分なスペースを削除',
    data: [
      ['文字列', '結果'],
      ['  Hello   World  ', '=TRIM(A2)']
    ],
    expectedValues: { 'B2': 'Hello World' }
  },
  {
    name: 'TEXT',
    category: '文字列',
    description: '数値を書式付き文字列に変換',
    data: [
      ['値', '表示形式', '結果'],
      [1234.56, '#,##0.00', '=TEXT(A2,B2)']
    ],
    expectedValues: { 'C2': '1,234.56' }
  },
  {
    name: 'VALUE',
    category: '文字列',
    description: '文字列を数値に変換',
    data: [
      ['文字列', '結果'],
      ['123.45', '=VALUE(A2)']
    ],
    expectedValues: { 'B2': 123.45 }
  },
  {
    name: 'REPT',
    category: '文字列',
    description: '文字列を繰り返す',
    data: [
      ['文字列', '繰り返し回数', '結果'],
      ['★', 5, '=REPT(A2,B2)']
    ],
    expectedValues: { 'C2': '★★★★★' }
  },
  {
    name: 'CHAR',
    category: '文字列',
    description: '文字コードから文字を返す',
    data: [
      ['文字コード', '文字'],
      [65, '=CHAR(A2)'],
      [97, '=CHAR(A3)']
    ],
    expectedValues: { 'B2': 'A', 'B3': 'a' }
  },
  {
    name: 'CODE',
    category: '文字列',
    description: '文字から文字コードを返す',
    data: [
      ['文字', '文字コード'],
      ['A', '=CODE(A2)'],
      ['a', '=CODE(A3)']
    ],
    expectedValues: { 'B2': 65, 'B3': 97 }
  },
  {
    name: 'EXACT',
    category: '文字列',
    description: '文字列が同一か判定',
    data: [
      ['文字列1', '文字列2', '結果'],
      ['Hello', 'Hello', '=EXACT(A2,B2)'],
      ['Hello', 'hello', '=EXACT(A3,B3)']
    ],
    expectedValues: { 'C2': true, 'C3': false }
  }
];

// 日付・時刻関数の個別テスト
export const dateFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'TODAY',
    category: '日付',
    description: '今日の日付を返す',
    data: [
      ['今日の日付'],
      ['=TODAY()']
    ]
  },
  {
    name: 'NOW',
    category: '日付',
    description: '現在の日時を返す',
    data: [
      ['現在の日時'],
      ['=NOW()']
    ]
  },
  {
    name: 'DATE',
    category: '日付',
    description: '日付を作成',
    data: [
      ['年', '月', '日', '日付'],
      [2024, 12, 25, '=DATE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 45651 } // Excelのシリアル値
  },
  {
    name: 'TIME',
    category: '日付',
    description: '時刻を作成',
    data: [
      ['時', '分', '秒', '時刻'],
      [13, 30, 45, '=TIME(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 0.563020833 }
  },
  {
    name: 'YEAR',
    category: '日付',
    description: '年を抽出',
    data: [
      ['日付', '年'],
      ['2024/12/25', '=YEAR(A2)']
    ],
    expectedValues: { 'B2': 2024 }
  },
  {
    name: 'MONTH',
    category: '日付',
    description: '月を抽出',
    data: [
      ['日付', '月'],
      ['2024/12/25', '=MONTH(A2)']
    ],
    expectedValues: { 'B2': 12 }
  },
  {
    name: 'DAY',
    category: '日付',
    description: '日を抽出',
    data: [
      ['日付', '日'],
      ['2024/12/25', '=DAY(A2)']
    ],
    expectedValues: { 'B2': 25 }
  },
  {
    name: 'HOUR',
    category: '日付',
    description: '時を抽出',
    data: [
      ['時刻', '時'],
      ['13:30:45', '=HOUR(A2)']
    ],
    expectedValues: { 'B2': 13 }
  },
  {
    name: 'MINUTE',
    category: '日付',
    description: '分を抽出',
    data: [
      ['時刻', '分'],
      ['13:30:45', '=MINUTE(A2)']
    ],
    expectedValues: { 'B2': 30 }
  },
  {
    name: 'SECOND',
    category: '日付',
    description: '秒を抽出',
    data: [
      ['時刻', '秒'],
      ['13:30:45', '=SECOND(A2)']
    ],
    expectedValues: { 'B2': 45 }
  },
  {
    name: 'WEEKDAY',
    category: '日付',
    description: '曜日を数値で返す',
    data: [
      ['日付', '曜日番号'],
      ['2024/12/25', '=WEEKDAY(A2)'] // 水曜日 = 4
    ],
    expectedValues: { 'B2': 4 }
  },
  {
    name: 'DATEDIF',
    category: '日付',
    description: '日付間の期間を計算',
    data: [
      ['開始日', '終了日', '単位', '期間'],
      ['2024/1/1', '2024/12/31', '"D"', '=DATEDIF(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 365 }
  },
  {
    name: 'DAYS',
    category: '日付',
    description: '日数差を計算',
    data: [
      ['開始日', '終了日', '日数'],
      ['2024/1/1', '2024/12/31', '=DAYS(B2,A2)']
    ],
    expectedValues: { 'C2': 365 }
  },
  {
    name: 'EDATE',
    category: '日付',
    description: '月数後の日付',
    data: [
      ['開始日', '月数', '結果'],
      ['2024/1/15', 3, '=EDATE(A2,B2)']
    ]
  },
  {
    name: 'EOMONTH',
    category: '日付',
    description: '月末日を返す',
    data: [
      ['開始日', '月数', '月末日'],
      ['2024/2/15', 0, '=EOMONTH(A2,B2)']
    ]
  }
];

// 論理関数の個別テスト
export const logicalFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'IF',
    category: '論理',
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
    category: '論理',
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
    category: '論理',
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
    category: '論理',
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
    category: '論理',
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
    category: '論理',
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
    category: '論理',
    description: '真を返す',
    data: [
      ['結果'],
      ['=TRUE()']
    ],
    expectedValues: { 'A2': true }
  },
  {
    name: 'FALSE',
    category: '論理',
    description: '偽を返す',
    data: [
      ['結果'],
      ['=FALSE()']
    ],
    expectedValues: { 'A2': false }
  },
  {
    name: 'IFERROR',
    category: '論理',
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
    category: '論理',
    description: '#N/Aエラー時の値',
    data: [
      ['検索値', '結果'],
      ['存在しない', '=IFNA(VLOOKUP(A2,D:E,2,FALSE),"見つかりません")']
    ],
    expectedValues: { 'B2': '見つかりません' }
  },
  {
    name: 'SWITCH',
    category: '論理',
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

// 検索・参照関数の個別テスト
export const lookupFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'VLOOKUP',
    category: '検索',
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
    category: '検索',
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
    category: '検索',
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
    category: '検索',
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
    category: '検索',
    description: '値の位置を検索',
    data: [
      ['値', '', '', '検索値', '位置'],
      [10, 20, 30, 20, '=MATCH(D2,A2:C2,0)']
    ],
    expectedValues: { 'E2': 2 }
  },
  {
    name: 'CHOOSE',
    category: '検索',
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
    category: '検索',
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
    category: '検索',
    description: '文字列を参照に変換',
    data: [
      ['値1', '値2', '', '参照文字列', '結果'],
      [100, 200, '', '"A2"', '=INDIRECT(D2)']
    ],
    expectedValues: { 'E2': 100 }
  },
  {
    name: 'ROW',
    category: '検索',
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
    category: '検索',
    description: '列番号を返す',
    data: [
      ['列1', '列2', '列3'],
      ['=COLUMN()', '=COLUMN()', '=COLUMN()']
    ],
    expectedValues: { 'A2': 1, 'B2': 2, 'C2': 3 }
  }
];

// 財務関数の個別テスト
export const financialFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'PMT',
    category: '財務',
    description: 'ローン定期支払額',
    data: [
      ['元本', '年利率', '期間(月)', '月額支払額'],
      [100000, 0.06, 12, '=PMT(B2/12,C2,-A2)']
    ]
  },
  {
    name: 'FV',
    category: '財務',
    description: '将来価値',
    data: [
      ['定期支払額', '利率', '期間', '将来価値'],
      [1000, 0.05, 10, '=FV(B2/12,C2,-A2)']
    ]
  },
  {
    name: 'PV',
    category: '財務',
    description: '現在価値',
    data: [
      ['定期支払額', '利率', '期間', '現在価値'],
      [1000, 0.05, 10, '=PV(B2/12,C2,-A2)']
    ]
  },
  {
    name: 'RATE',
    category: '財務',
    description: '利率',
    data: [
      ['期間', '支払額', '現在価値', '利率'],
      [12, -1000, 10000, '=RATE(A2,B2,C2)*12']
    ]
  },
  {
    name: 'NPER',
    category: '財務',
    description: '支払回数',
    data: [
      ['利率', '支払額', '現在価値', '期間'],
      [0.05, -1000, 10000, '=NPER(A2/12,B2,C2)']
    ]
  },
  {
    name: 'NPV',
    category: '財務',
    description: '正味現在価値',
    data: [
      ['割引率', 'CF1', 'CF2', 'CF3', 'NPV'],
      [0.1, -10000, 3000, 4200, '=NPV(A2,B2:D2)']
    ]
  },
  {
    name: 'IRR',
    category: '財務',
    description: '内部収益率',
    data: [
      ['初期投資', 'CF1', 'CF2', 'CF3', 'IRR'],
      [-10000, 3000, 4200, 6800, '=IRR(A2:D2)']
    ]
  }
];

// 情報関数の個別テスト
export const informationFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'ISBLANK',
    category: '情報',
    description: '空白セルか判定',
    data: [
      ['値', '空白判定'],
      ['テキスト', '=ISBLANK(A2)'],
      ['', '=ISBLANK(A3)'],
      [0, '=ISBLANK(A4)']
    ],
    expectedValues: { 'B2': false, 'B3': true, 'B4': false }
  },
  {
    name: 'ISERROR',
    category: '情報',
    description: 'エラー値か判定',
    data: [
      ['値', 'エラー判定'],
      ['=1/0', '=ISERROR(A2)'],
      [100, '=ISERROR(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'ISNA',
    category: '情報',
    description: '#N/Aエラーか判定',
    data: [
      ['値', '#N/A判定'],
      ['=NA()', '=ISNA(A2)'],
      ['#DIV/0!', '=ISNA(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'ISTEXT',
    category: '情報',
    description: '文字列か判定',
    data: [
      ['値', '文字列判定'],
      ['テキスト', '=ISTEXT(A2)'],
      [123, '=ISTEXT(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'ISNUMBER',
    category: '情報',
    description: '数値か判定',
    data: [
      ['値', '数値判定'],
      [123, '=ISNUMBER(A2)'],
      ['123', '=ISNUMBER(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'ISLOGICAL',
    category: '情報',
    description: '論理値か判定',
    data: [
      ['値', '論理値判定'],
      ['=TRUE()', '=ISLOGICAL(A2)'],
      [1, '=ISLOGICAL(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'TYPE',
    category: '情報',
    description: 'データ型を返す',
    data: [
      ['値', 'データ型'],
      [123, '=TYPE(A2)'],
      ['テキスト', '=TYPE(A3)'],
      ['=TRUE()', '=TYPE(A4)']
    ],
    expectedValues: { 'B2': 1, 'B3': 2, 'B4': 4 }
  },
  {
    name: 'N',
    category: '情報',
    description: '数値に変換',
    data: [
      ['値', '数値変換'],
      [7, '=N(A2)'],
      ['7', '=N(A3)'],
      ['=TRUE()', '=N(A4)']
    ],
    expectedValues: { 'B2': 7, 'B3': 0, 'B4': 1 }
  }
];

// データベース関数の個別テスト
export const databaseFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'DSUM',
    category: 'データベース',
    description: '条件付き合計',
    data: [
      ['名前', '部署', '売上'],
      ['田中', '営業', 100],
      ['佐藤', '営業', 150],
      ['鈴木', '開発', 120],
      ['', '', ''],
      ['条件', '', ''],
      ['部署', '', ''],
      ['営業', '', '=DSUM(A1:C4,C1,A6:A7)']
    ],
    expectedValues: { 'C7': 250 }
  },
  {
    name: 'DAVERAGE',
    category: 'データベース',
    description: '条件付き平均',
    data: [
      ['名前', '部署', '売上'],
      ['田中', '営業', 100],
      ['佐藤', '営業', 150],
      ['鈴木', '開発', 120],
      ['', '', ''],
      ['条件', '', ''],
      ['部署', '', ''],
      ['営業', '', '=DAVERAGE(A1:C4,C1,A6:A7)']
    ],
    expectedValues: { 'C7': 125 }
  },
  {
    name: 'DCOUNT',
    category: 'データベース',
    description: '条件付きカウント',
    data: [
      ['名前', '部署', '売上'],
      ['田中', '営業', 100],
      ['佐藤', '営業', 150],
      ['鈴木', '開発', 120],
      ['', '', ''],
      ['条件', '', ''],
      ['部署', '', ''],
      ['営業', '', '=DCOUNT(A1:C4,C1,A6:A7)']
    ],
    expectedValues: { 'C7': 2 }
  }
];

// エンジニアリング関数の個別テスト
export const engineeringFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'CONVERT',
    category: 'エンジニアリング',
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
    category: 'エンジニアリング',
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
    category: 'エンジニアリング',
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
    category: 'エンジニアリング',
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
    category: 'エンジニアリング',
    description: '複素数を作成',
    data: [
      ['実部', '虚部', '複素数'],
      [3, 4, '=COMPLEX(A2,B2)'],
      [0, 1, '=COMPLEX(A3,B3)']
    ],
    expectedValues: { 'C2': '3+4i', 'C3': 'i' }
  }
];

// 動的配列関数の個別テスト
export const dynamicArrayFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'FILTER',
    category: '動的配列',
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
    category: '動的配列',
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
    category: '動的配列',
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
    category: '動的配列',
    description: '行列を入れ替え',
    data: [
      ['', 'A', 'B', 'C'],
      ['1', 1, 2, 3],
      ['2', 4, 5, 6],
      ['', '', '', ''],
      ['転置', '=TRANSPOSE(B2:D3)', '', '']
    ]
  }
];

// Web関数の個別テスト
export const webFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'ENCODEURL',
    category: 'Web',
    description: 'URLエンコード',
    data: [
      ['テキスト', 'エンコード結果'],
      ['Hello World', '=ENCODEURL(A2)'],
      ['こんにちは', '=ENCODEURL(A3)']
    ],
    expectedValues: { 'B2': 'Hello%20World' }
  }
];

// すべての個別テストを結合
export const allIndividualFunctionTests: IndividualFunctionTest[] = [
  ...mathFunctionTests,
  ...statisticalFunctionTests,
  ...textFunctionTests,
  ...dateFunctionTests,
  ...logicalFunctionTests,
  ...lookupFunctionTests,
  ...financialFunctionTests,
  ...informationFunctionTests,
  ...databaseFunctionTests,
  ...engineeringFunctionTests,
  ...dynamicArrayFunctionTests,
  ...webFunctionTests
];

// カテゴリ別にテストを取得
export function getIndividualTestsByCategory(category: string): IndividualFunctionTest[] {
  return allIndividualFunctionTests.filter(test => test.category === category);
}

// 関数名でテストを取得
export function getIndividualTestByName(name: string): IndividualFunctionTest | undefined {
  return allIndividualFunctionTests.find(test => test.name === name);
}

// すべてのカテゴリを取得
export function getAllIndividualCategories(): string[] {
  return [...new Set(allIndividualFunctionTests.map(test => test.category))];
}