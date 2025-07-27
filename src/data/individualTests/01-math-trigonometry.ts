import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const mathTrigonometryTests: IndividualFunctionTest[] = [
  {
    name: 'SUM',
    category: '01. 数学',
    description: '数値の合計を計算',
    data: [
      ['値1', '値2', '値3', '値4', '結果'],
      [10, 20, 30, 40, '=SUM(A2:D2)']
    ],
    expectedValues: { 'E2': 100 }
  },
  {
    name: 'SUMIF',
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
    description: '数値の積を計算',
    data: [
      ['値1', '値2', '値3', '結果'],
      [2, 3, 4, '=PRODUCT(A2:C2)']
    ],
    expectedValues: { 'D2': 24 }
  },
  {
    name: 'SQRT',
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
    description: '0以上1未満の乱数',
    data: [
      ['乱数1', '乱数2', '乱数3'],
      ['=RAND()', '=RAND()', '=RAND()']
    ]
  },
  {
    name: 'RANDBETWEEN',
    category: '01. 数学',
    description: '指定範囲の整数乱数',
    data: [
      ['最小', '最大', '結果'],
      [1, 10, '=RANDBETWEEN(A2,B2)'],
      [1, 100, '=RANDBETWEEN(A3,B3)']
    ]
  },
  {
    name: 'EXP',
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
    description: '配列の積の和を計算',
    data: [
      ['配列1', '', '', '配列2', '', '', '結果'],
      [2, 3, 4, 5, 6, 7, '=SUMPRODUCT(A2:C2,D2:F2)']
    ],
    expectedValues: { 'G2': 56 }
  },
  {
    name: 'SIN',
    category: '01. 数学・三角',
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
    category: '01. 数学・三角',
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
    category: '01. 数学・三角',
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
    category: '01. 数学・三角',
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
    category: '01. 数学・三角',
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
    category: '01. 数学・三角',
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
    category: '01. 数学・三角',
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
    category: '01. 数学・三角',
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
    category: '01. 数学・三角',
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
    category: '01. 数学・三角',
    description: '円周率を返す',
    data: [
      ['円周率'],
      ['=PI()']
    ],
    expectedValues: { 'A2': 3.141592654 }
  },
  {
    name: 'SUMX2MY2',
    category: '01. 数学',
    description: 'x^2-y^2の和を計算',
    data: [
      ['X値', 'Y値'],
      [3, 2],
      [4, 3],
      [5, 4],
      ['結果', '=SUMX2MY2(A2:A4,B2:B4)']
    ],
    expectedValues: { 'B5': 21 }
  },
  {
    name: 'SUMX2PY2',
    category: '01. 数学',
    description: 'x^2+y^2の和を計算',
    data: [
      ['X値', 'Y値'],
      [3, 4],
      [0, 1],
      [1, 0],
      ['結果', '=SUMX2PY2(A2:A4,B2:B4)']
    ],
    expectedValues: { 'B5': 27 }
  },
  {
    name: 'SUMXMY2',
    category: '01. 数学',
    description: '(x-y)^2の和を計算',
    data: [
      ['X値', 'Y値'],
      [5, 3],
      [7, 4],
      [2, 1],
      ['結果', '=SUMXMY2(A2:A4,B2:B4)']
    ],
    expectedValues: { 'B5': 14 }
  },
  {
    name: 'CEILING.MATH',
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学',
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
    category: '01. 数学・三角',
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
    category: '01. 数学・三角',
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
    category: '01. 数学・三角',
    description: '双曲線正接を計算',
    data: [
      ['値', '双曲線正接'],
      [0, '=TANH(A2)'],
      [0.5, '=TANH(A3)']
    ],
    expectedValues: { 'B2': 0, 'B3': 0.462117 }
  },
  {
    name: 'ASINH',
    category: '01. 数学・三角',
    description: '逆双曲線正弦',
    data: [
      ['値', '逆双曲線正弦'],
      [0, '=ASINH(A2)'],
      [1, '=ASINH(A3)']
    ],
    expectedValues: { 'B2': 0, 'B3': 0.881374 }
  },
  {
    name: 'ACOSH',
    category: '01. 数学・三角',
    description: '逆双曲線余弦',
    data: [
      ['値', '逆双曲線余弦'],
      [1, '=ACOSH(A2)'],
      [2, '=ACOSH(A3)']
    ],
    expectedValues: { 'B2': 0, 'B3': 1.316958 }
  },
  {
    name: 'ATANH',
    category: '01. 数学・三角',
    description: '逆双曲線正接',
    data: [
      ['値', '逆双曲線正接'],
      [0, '=ATANH(A2)'],
      [0.5, '=ATANH(A3)']
    ],
    expectedValues: { 'B2': 0, 'B3': 0.549306 }
  },
  {
    name: 'SEC',
    category: '01. 数学・三角',
    description: '正割',
    data: [
      ['角度(ラジアン)', '正割'],
      [0, '=SEC(A2)'],
      [1, '=SEC(A3)']
    ],
    expectedValues: { 'B2': 1, 'B3': 1.850816 }
  },
  {
    name: 'CSC',
    category: '01. 数学・三角',
    description: '余割',
    data: [
      ['角度(ラジアン)', '余割'],
      [1, '=CSC(A2)'],
      [2, '=CSC(A3)']
    ],
    expectedValues: { 'B2': 1.188395, 'B3': 1.099750 }
  },
  {
    name: 'COT',
    category: '01. 数学・三角',
    description: '余接',
    data: [
      ['角度(ラジアン)', '余接'],
      [1, '=COT(A2)'],
      [2, '=COT(A3)']
    ],
    expectedValues: { 'B2': 0.642093, 'B3': -0.457658 }
  },
  {
    name: 'ACOT',
    category: '01. 数学・三角',
    description: '逆余接',
    data: [
      ['値', '逆余接'],
      [1, '=ACOT(A2)'],
      [0, '=ACOT(A3)']
    ],
    expectedValues: { 'B2': 0.785398, 'B3': 1.570796 }
  },
  {
    name: 'ARABIC',
    category: '01. 数学',
    description: 'ローマ数字をアラビア数字に',
    data: [
      ['ローマ数字', 'アラビア数字'],
      ['XVI', '=ARABIC(A2)'],
      ['MCMXII', '=ARABIC(A3)']
    ],
    expectedValues: { 'B2': 16, 'B3': 1912 }
  },
  {
    name: 'ROMAN',
    category: '01. 数学',
    description: 'アラビア数字をローマ数字に',
    data: [
      ['アラビア数字', 'ローマ数字'],
      [16, '=ROMAN(A2)'],
      [1912, '=ROMAN(A3)']
    ],
    expectedValues: { 'B2': 'XVI', 'B3': 'MCMXII' }
  },
  {
    name: 'BASE',
    category: '01. 数学',
    description: '基数変換',
    data: [
      ['数値', '基数', '最小桁数', '結果'],
      [15, 2, 8, '=BASE(A2,B2,C2)'],
      [255, 16, 4, '=BASE(A3,B3,C3)']
    ],
    expectedValues: { 'D2': '00001111', 'D3': '00FF' }
  },
  {
    name: 'DECIMAL',
    category: '01. 数学',
    description: '基数から10進数に変換',
    data: [
      ['テキスト', '基数', '10進数'],
      ['FF', 16, '=DECIMAL(A2,B2)'],
      ['1111', 2, '=DECIMAL(A3,B3)']
    ],
    expectedValues: { 'C2': 255, 'C3': 15 }
  },
  {
    name: 'SQRTPI',
    category: '01. 数学',
    description: '数値×πの平方根',
    data: [
      ['数値', '√(数値×π)'],
      [1, '=SQRTPI(A2)'],
      [2, '=SQRTPI(A3)']
    ],
    expectedValues: { 'B2': 1.772454, 'B3': 2.506628 }
  },
  {
    name: 'FACTDOUBLE',
    category: '01. 数学',
    description: '二重階乗',
    data: [
      ['数値', '二重階乗'],
      [5, '=FACTDOUBLE(A2)'],
      [6, '=FACTDOUBLE(A3)']
    ],
    expectedValues: { 'B2': 15, 'B3': 48 }
  },
  {
    name: 'MULTINOMIAL',
    category: '01. 数学',
    description: '多項係数',
    data: [
      ['値1', '値2', '値3', '多項係数'],
      [2, 3, 4, '=MULTINOMIAL(A2:C2)']
    ],
    expectedValues: { 'D2': 1260 }
  },
  {
    name: 'COMBINA',
    category: '01. 数学',
    description: '重複組み合わせ',
    data: [
      ['要素数', '選択数', '重複組み合わせ'],
      [4, 3, '=COMBINA(A2,B2)'],
      [10, 3, '=COMBINA(A3,B3)']
    ],
    expectedValues: { 'C2': 20, 'C3': 220 }
  },
  {
    name: 'PERMUTATIONA',
    category: '01. 数学',
    description: '重複順列',
    data: [
      ['要素数', '選択数', '重複順列'],
      [4, 2, '=PERMUTATIONA(A2,B2)'],
      [3, 3, '=PERMUTATIONA(A3,B3)']
    ],
    expectedValues: { 'C2': 16, 'C3': 27 }
  },
  {
    name: 'SERIESSUM',
    category: '01. 数学',
    description: 'べき級数の和',
    data: [
      ['x', 'n', 'm', '係数', '', '', 'べき級数'],
      [2, 1, 1, 1, 2, 3, '=SERIESSUM(A2,B2,C2,D2:F2)']
    ]
  },
  {
    name: 'CEILING.PRECISE',
    category: '01. 数学',
    description: '正確な切り上げ',
    data: [
      ['数値', '基準値', '切り上げ'],
      [4.3, 1, '=CEILING.PRECISE(A2,B2)'],
      [-4.3, 1, '=CEILING.PRECISE(A3,B3)']
    ],
    expectedValues: { 'C2': 5, 'C3': -4 }
  },
  {
    name: 'FLOOR.PRECISE',
    category: '01. 数学',
    description: '正確な切り捨て',
    data: [
      ['数値', '基準値', '切り捨て'],
      [4.3, 1, '=FLOOR.PRECISE(A2)'],
      [-4.3, 1, '=FLOOR.PRECISE(A3,B3)']
    ],
    expectedValues: { 'C2': 4, 'C3': -5 }
  },
  {
    name: 'ISO.CEILING',
    category: '01. 数学',
    description: 'ISO規格の切り上げ',
    data: [
      ['数値', '基準値', '切り上げ'],
      [4.3, 1, '=ISO.CEILING(A2,B2)'],
      [-4.3, 1, '=ISO.CEILING(A3,B3)']
    ],
    expectedValues: { 'C2': 5, 'C3': -4 }
  },
  {
    name: 'AGGREGATE',
    category: '01. 数学',
    description: '集約関数',
    data: [
      ['機能番号', 'オプション', '配列', '結果'],
      [1, 0, 'C4:C7', '=AGGREGATE(A2,B2,C4:C7)'],
      [4, 0, 'C4:C7', '=AGGREGATE(A3,B3,C4:C7)'],
      ['', '', 10, ''],
      ['', '', 20, ''],
      ['', '', 30, ''],
      ['', '', 40, '']
    ],
    expectedValues: { 'D2': 25, 'D3': 40 }
  },
  {
    name: 'ACOTH',
    category: '01. 数学・三角',
    description: '双曲線逆余接',
    data: [
      ['値', '双曲線逆余接'],
      [2, '=ACOTH(A2)'],
      [5, '=ACOTH(A3)'],
      [10, '=ACOTH(A4)']
    ],
    expectedValues: { 'B2': 0.549306, 'B3': 0.202733, 'B4': 0.100335 }
  },
  {
    name: 'COTH',
    category: '01. 数学・三角',
    description: '双曲線余接',
    data: [
      ['値', '双曲線余接'],
      [1, '=COTH(A2)'],
      [2, '=COTH(A3)'],
      [3, '=COTH(A4)']
    ],
    expectedValues: { 'B2': 1.313035, 'B3': 1.037315, 'B4': 1.004964 }
  },
  {
    name: 'CSCH',
    category: '01. 数学・三角',
    description: '双曲線余割',
    data: [
      ['値', '双曲線余割'],
      [1, '=CSCH(A2)'],
      [2, '=CSCH(A3)'],
      [3, '=CSCH(A4)']
    ],
    expectedValues: { 'B2': 0.850918, 'B3': 0.275721, 'B4': 0.099821 }
  },
  {
    name: 'SECH',
    category: '01. 数学・三角',
    description: '双曲線正割',
    data: [
      ['値', '双曲線正割'],
      [0, '=SECH(A2)'],
      [1, '=SECH(A3)'],
      [2, '=SECH(A4)']
    ],
    expectedValues: { 'B2': 1, 'B3': 0.648054, 'B4': 0.265802 }
  }
];
