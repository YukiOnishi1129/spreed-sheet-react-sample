export interface IndividualFunctionTest {
  name: string;
  category: string;
  description: string;
  data: (string | number | boolean | null)[][];
  expectedValues?: { [key: string]: string | number | boolean };
}

// 基本演算子の個別テスト
export const basicOperatorTests: IndividualFunctionTest[] = [
  {
    name: 'ADD_OPERATOR',
    category: '00. 基本演算子',
    description: '加算演算子（+）',
    data: [
      ['値1', '値2', '結果'],
      [10, 20, '=A2+B2'],
      [5.5, 4.5, '=A3+B3'],
      [-10, 5, '=A4+B4']
    ],
    expectedValues: { 'C2': 30, 'C3': 10, 'C4': -5 }
  },
  {
    name: 'SUBTRACT_OPERATOR',
    category: '00. 基本演算子',
    description: '減算演算子（-）',
    data: [
      ['値1', '値2', '結果'],
      [30, 10, '=A2-B2'],
      [5.5, 4.5, '=A3-B3'],
      [-10, 5, '=A4-B4']
    ],
    expectedValues: { 'C2': 20, 'C3': 1, 'C4': -15 }
  },
  {
    name: 'MULTIPLY_OPERATOR',
    category: '00. 基本演算子',
    description: '乗算演算子（*）',
    data: [
      ['値1', '値2', '結果'],
      [10, 3, '=A2*B2'],
      [2.5, 4, '=A3*B3'],
      [-5, 3, '=A4*B4']
    ],
    expectedValues: { 'C2': 30, 'C3': 10, 'C4': -15 }
  },
  {
    name: 'DIVIDE_OPERATOR',
    category: '00. 基本演算子',
    description: '除算演算子（/）',
    data: [
      ['値1', '値2', '結果'],
      [20, 4, '=A2/B2'],
      [15, 3, '=A3/B3'],
      [-10, 2, '=A4/B4']
    ],
    expectedValues: { 'C2': 5, 'C3': 5, 'C4': -5 }
  },
  {
    name: 'GREATER_THAN',
    category: '00. 基本演算子',
    description: '大なり演算子（>）',
    data: [
      ['値1', '値2', '結果'],
      [10, 5, '=A2>B2'],
      [5, 10, '=A3>B3'],
      [5, 5, '=A4>B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': false }
  },
  {
    name: 'LESS_THAN',
    category: '00. 基本演算子',
    description: '小なり演算子（<）',
    data: [
      ['値1', '値2', '結果'],
      [5, 10, '=A2<B2'],
      [10, 5, '=A3<B3'],
      [5, 5, '=A4<B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': false }
  },
  {
    name: 'GREATER_THAN_OR_EQUAL',
    category: '00. 基本演算子',
    description: '以上演算子（>=）',
    data: [
      ['値1', '値2', '結果'],
      [10, 5, '=A2>=B2'],
      [5, 10, '=A3>=B3'],
      [5, 5, '=A4>=B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': true }
  },
  {
    name: 'LESS_THAN_OR_EQUAL',
    category: '00. 基本演算子',
    description: '以下演算子（<=）',
    data: [
      ['値1', '値2', '結果'],
      [5, 10, '=A2<=B2'],
      [10, 5, '=A3<=B3'],
      [5, 5, '=A4<=B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': true }
  },
  {
    name: 'EQUAL',
    category: '00. 基本演算子',
    description: '等価演算子（=）',
    data: [
      ['値1', '値2', '結果'],
      [10, 10, '=A2=B2'],
      [5, 10, '=A3=B3'],
      ['ABC', 'ABC', '=A4=B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': true }
  },
  {
    name: 'NOT_EQUAL',
    category: '00. 基本演算子',
    description: '不等価演算子（<>）',
    data: [
      ['値1', '値2', '結果'],
      [10, 5, '=A2<>B2'],
      [10, 10, '=A3<>B3'],
      ['ABC', 'DEF', '=A4<>B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': true }
  }
];

// 数学・三角関数の個別テスト
export const mathFunctionTests: IndividualFunctionTest[] = [
  // 基本的な数学関数
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
  // 三角関数
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
  // 追加の数学関数
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
  }
];

// 統計関数の個別テスト
export const statisticalFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'AVERAGE',
    category: '02. 統計',
    description: '平均値を計算',
    data: [
      ['値1', '値2', '値3', '値4', '平均'],
      [10, 20, 30, 40, '=AVERAGE(A2:D2)']
    ],
    expectedValues: { 'E2': 25 }
  },
  {
    name: 'AVERAGEIF',
    category: '02. 統計',
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
    category: '02. 統計',
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
    category: '02. 統計',
    description: '数値の個数を数える',
    data: [
      ['値1', '値2', '値3', '値4', '個数'],
      [10, 20, 'テキスト', 30, '=COUNT(A2:D2)']
    ],
    expectedValues: { 'E2': 3 }
  },
  {
    name: 'COUNTA',
    category: '02. 統計',
    description: '空白以外のセル数',
    data: [
      ['値1', '値2', '値3', '値4', '個数'],
      [10, 20, 'テキスト', '', '=COUNTA(A2:D2)']
    ],
    expectedValues: { 'E2': 3 }
  },
  {
    name: 'COUNTBLANK',
    category: '02. 統計',
    description: '空白セルの個数',
    data: [
      ['値1', '値2', '値3', '値4', '空白数'],
      [10, '', 'テキスト', '', '=COUNTBLANK(A2:D2)']
    ],
    expectedValues: { 'E2': 2 }
  },
  {
    name: 'COUNTIF',
    category: '02. 統計',
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
    category: '02. 統計',
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
    category: '02. 統計',
    description: '最大値を返す',
    data: [
      ['値1', '値2', '値3', '値4', '最大値'],
      [10, 30, 20, 40, '=MAX(A2:D2)']
    ],
    expectedValues: { 'E2': 40 }
  },
  {
    name: 'MIN',
    category: '02. 統計',
    description: '最小値を返す',
    data: [
      ['値1', '値2', '値3', '値4', '最小値'],
      [10, 30, 20, 40, '=MIN(A2:D2)']
    ],
    expectedValues: { 'E2': 10 }
  },
  {
    name: 'MAXIFS',
    category: '02. 統計',
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
    category: '02. 統計',
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
    category: '02. 統計',
    description: '中央値を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '中央値'],
      [1, 3, 3, 6, 7, '=MEDIAN(A2:E2)']
    ],
    expectedValues: { 'F2': 3 }
  },
  {
    name: 'MODE',
    category: '02. 統計',
    description: '最頻値を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '最頻値'],
      [1, 2, 3, 3, 4, '=MODE(A2:E2)']
    ],
    expectedValues: { 'F2': 3 }
  },
  {
    name: 'STDEV',
    category: '02. 統計',
    description: '標準偏差（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '標準偏差'],
      [10, 20, 30, 40, '=STDEV(A2:D2)']
    ],
    expectedValues: { 'E2': 12.90994449 }
  },
  {
    name: 'STDEV.S',
    category: '02. 統計',
    description: '標準偏差（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '標準偏差'],
      [10, 20, 30, 40, '=STDEV.S(A2:D2)']
    ],
    expectedValues: { 'E2': 12.90994449 }
  },
  {
    name: 'STDEV.P',
    category: '02. 統計',
    description: '標準偏差（母集団）',
    data: [
      ['値1', '値2', '値3', '値4', '標準偏差'],
      [10, 20, 30, 40, '=STDEV.P(A2:D2)']
    ],
    expectedValues: { 'E2': 11.18033989 }
  },
  {
    name: 'VAR',
    category: '02. 統計',
    description: '分散（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '分散'],
      [10, 20, 30, 40, '=VAR(A2:D2)']
    ],
    expectedValues: { 'E2': 166.6666667 }
  },
  {
    name: 'VAR.S',
    category: '02. 統計',
    description: '分散（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '分散'],
      [10, 20, 30, 40, '=VAR.S(A2:D2)']
    ],
    expectedValues: { 'E2': 166.6666667 }
  },
  {
    name: 'VAR.P',
    category: '02. 統計',
    description: '分散（母集団）',
    data: [
      ['値1', '値2', '値3', '値4', '分散'],
      [10, 20, 30, 40, '=VAR.P(A2:D2)']
    ],
    expectedValues: { 'E2': 125 }
  },
  {
    name: 'CORREL',
    category: '02. 統計',
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
    category: '02. 統計',
    description: 'k番目に大きい値',
    data: [
      ['値', '', '', '', 'k', '結果'],
      [10, 30, 20, 40, 2, '=LARGE(A2:D2,E2)']
    ],
    expectedValues: { 'F2': 30 }
  },
  {
    name: 'SMALL',
    category: '02. 統計',
    description: 'k番目に小さい値',
    data: [
      ['値', '', '', '', 'k', '結果'],
      [10, 30, 20, 40, 2, '=SMALL(A2:D2,E2)']
    ],
    expectedValues: { 'F2': 20 }
  },
  {
    name: 'RANK',
    category: '02. 統計',
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
    category: '02. 統計',
    description: 'パーセンタイル値',
    data: [
      ['値', '', '', '', 'k', '結果'],
      [1, 2, 3, 4, 0.5, '=PERCENTILE(A2:D2,E2)']
    ],
    expectedValues: { 'F2': 2.5 }
  },
  {
    name: 'QUARTILE',
    category: '02. 統計',
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
    category: '02. 統計',
    description: '歪度を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '歪度'],
      [3, 4, 5, 2, 1, '=SKEW(A2:E2)']
    ]
  },
  {
    name: 'SKEW.P',
    category: '02. 統計',
    description: '歪度（母集団）',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '歪度'],
      [3, 4, 5, 2, 1, '=SKEW.P(A2:E2)']
    ]
  },
  {
    name: 'MODE.SNGL',
    category: '02. 統計',
    description: '最頻値（単一）',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '最頻値'],
      [1, 2, 2, 3, 3, '=MODE.SNGL(A2:E2)']
    ]
  },
  {
    name: 'QUARTILE.INC',
    category: '02. 統計',
    description: '四分位数（境界値含む）',
    data: [
      ['値1', '値2', '値3', '値4', '第1四分位'],
      [1, 2, 3, 4, '=QUARTILE.INC(A2:D2,1)']
    ]
  },
  {
    name: 'PERCENTILE.INC',
    category: '02. 統計',
    description: 'パーセンタイル（境界値含む）',
    data: [
      ['値1', '値2', '値3', '値4', '50%タイル'],
      [1, 2, 3, 4, '=PERCENTILE.INC(A2:D2,0.5)']
    ]
  },
  {
    name: 'RANK.EQ',
    category: '02. 統計',
    description: '順位（同順位は最小）',
    data: [
      ['値', '順位'],
      [85, '=RANK.EQ(A2,$A$2:$A$5)'],
      [92, '=RANK.EQ(A3,$A$2:$A$5)'],
      [78, '=RANK.EQ(A4,$A$2:$A$5)'],
      [92, '=RANK.EQ(A5,$A$2:$A$5)']
    ]
  },
  {
    name: 'KURT',
    category: '02. 統計',
    description: '尖度を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '尖度'],
      [3, 4, 5, 2, 1, '=KURT(A2:E2)']
    ]
  },
  {
    name: 'STANDARDIZE',
    category: '02. 統計',
    description: '標準化する',
    data: [
      ['値', '平均', '標準偏差', '標準化値'],
      [42, 40, 1.5, '=STANDARDIZE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 1.333333 }
  },
  {
    name: 'DEVSQ',
    category: '02. 統計',
    description: '偏差平方和を計算',
    data: [
      ['値1', '値2', '値3', '値4', '偏差平方和'],
      [4, 5, 8, 7, '=DEVSQ(A2:D2)']
    ],
    expectedValues: { 'E2': 10 }
  },
  {
    name: 'GEOMEAN',
    category: '02. 統計',
    description: '幾何平均を計算',
    data: [
      ['値1', '値2', '値3', '値4', '幾何平均'],
      [4, 5, 8, 7, '=GEOMEAN(A2:D2)']
    ]
  },
  {
    name: 'HARMEAN',
    category: '02. 統計',
    description: '調和平均を計算',
    data: [
      ['値1', '値2', '値3', '値4', '調和平均'],
      [4, 5, 8, 7, '=HARMEAN(A2:D2)']
    ]
  },
  {
    name: 'TRIMMEAN',
    category: '02. 統計',
    description: 'トリム平均を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '値6', '除外率', 'トリム平均'],
      [1, 2, 3, 4, 5, 100, 0.2, '=TRIMMEAN(A2:F2,G2)']
    ]
  },
  // 統計分布関数（高優先度）
  {
    name: 'NORM.DIST',
    category: '02. 統計',
    description: '正規分布',
    data: [
      ['値', '平均', '標準偏差', '累積', '確率密度'],
      [42, 40, 1.5, 'FALSE', '=NORM.DIST(A2,B2,C2,D2)'],
      [42, 40, 1.5, 'TRUE', '=NORM.DIST(A3,B3,C3,D3)']
    ]
  },
  {
    name: 'NORM.INV',
    category: '02. 統計',
    description: '正規分布の逆関数',
    data: [
      ['確率', '平均', '標準偏差', '値'],
      [0.908789, 40, 1.5, '=NORM.INV(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 42 }
  },
  {
    name: 'NORM.S.DIST',
    category: '02. 統計',
    description: '標準正規分布',
    data: [
      ['z値', '累積', '確率'],
      [1.333333, 'TRUE', '=NORM.S.DIST(A2,B2)'],
      [1.333333, 'FALSE', '=NORM.S.DIST(A3,B3)']
    ]
  },
  {
    name: 'NORM.S.INV',
    category: '02. 統計',
    description: '標準正規分布の逆関数',
    data: [
      ['確率', 'z値'],
      [0.908789, '=NORM.S.INV(A2)']
    ],
    expectedValues: { 'B2': 1.333333 }
  },
  {
    name: 'T.DIST',
    category: '02. 統計',
    description: 't分布（左側）',
    data: [
      ['値', '自由度', '累積', '確率'],
      [1.96, 60, 'TRUE', '=T.DIST(A2,B2,C2)']
    ]
  },
  {
    name: 'T.DIST.2T',
    category: '02. 統計',
    description: 't分布（両側）',
    data: [
      ['値', '自由度', '確率'],
      [1.96, 60, '=T.DIST.2T(A2,B2)']
    ]
  },
  {
    name: 'T.INV',
    category: '02. 統計',
    description: 't分布の逆関数（左側）',
    data: [
      ['確率', '自由度', '値'],
      [0.975, 60, '=T.INV(A2,B2)']
    ],
    expectedValues: { 'C2': 2.000298 }
  },
  {
    name: 'T.INV.2T',
    category: '02. 統計',
    description: 't分布の逆関数（両側）',
    data: [
      ['確率', '自由度', '値'],
      [0.05, 60, '=T.INV.2T(A2,B2)']
    ],
    expectedValues: { 'C2': 2.000298 }
  },
  {
    name: 'CHISQ.DIST',
    category: '02. 統計',
    description: 'カイ二乗分布',
    data: [
      ['値', '自由度', '累積', '確率'],
      [18.307, 10, 'TRUE', '=CHISQ.DIST(A2,B2,C2)']
    ]
  },
  {
    name: 'CHISQ.INV',
    category: '02. 統計',
    description: 'カイ二乗分布の逆関数',
    data: [
      ['確率', '自由度', '値'],
      [0.95, 10, '=CHISQ.INV(A2,B2)']
    ],
    expectedValues: { 'C2': 18.307 }
  },
  {
    name: 'F.DIST',
    category: '02. 統計',
    description: 'F分布',
    data: [
      ['値', '自由度1', '自由度2', '累積', '確率'],
      [2.5, 5, 10, 'TRUE', '=F.DIST(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 0.8823 }
  },
  {
    name: 'F.INV',
    category: '02. 統計',
    description: 'F分布の逆関数',
    data: [
      ['確率', '自由度1', '自由度2', '値'],
      [0.95, 5, 10, '=F.INV(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 3.3258 }
  },
  {
    name: 'EXPON.DIST',
    category: '02. 統計',
    description: '指数分布',
    data: [
      ['値', 'λ', '累積', '確率'],
      [0.2, 10, 'TRUE', '=EXPON.DIST(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 0.8647 }
  },
  {
    name: 'COVARIANCE.P',
    category: '02. 統計',
    description: '母共分散',
    data: [
      ['データ1', 'データ2', '', '共分散'],
      [1, 10, '', '=COVARIANCE.P(A2:A5,B2:B5)'],
      [2, 15, '', ''],
      [3, 20, '', ''],
      [4, 25, '', '']
    ],
    expectedValues: { 'D2': 6.25 }
  },
  {
    name: 'COVARIANCE.S',
    category: '02. 統計',
    description: '標本共分散',
    data: [
      ['データ1', 'データ2', '', '共分散'],
      [1, 10, '', '=COVARIANCE.S(A2:A5,B2:B5)'],
      [2, 15, '', ''],
      [3, 20, '', ''],
      [4, 25, '', '']
    ]
  },
  {
    name: 'MODE.MULT',
    category: '02. 統計',
    description: '最頻値（複数）',
    data: [
      ['データ', '最頻値'],
      [1, '=MODE.MULT(A2:A8)'],
      [2, ''],
      [2, ''],
      [3, ''],
      [3, ''],
      [3, ''],
      [4, '']
    ]
  },
  {
    name: 'RANK.AVG',
    category: '02. 統計',
    description: '順位（平均）',
    data: [
      ['データ', '順位'],
      [89, '=RANK.AVG(A2,$A$2:$A$6)'],
      [92, '=RANK.AVG(A3,$A$2:$A$6)'],
      [85, '=RANK.AVG(A4,$A$2:$A$6)'],
      [92, '=RANK.AVG(A5,$A$2:$A$6)'],
      [88, '=RANK.AVG(A6,$A$2:$A$6)']
    ],
    expectedValues: { 'B2': 3, 'B3': 1.5, 'B4': 5, 'B5': 1.5, 'B6': 4 }
  },
  {
    name: 'LOGNORM.DIST',
    category: '02. 統計',
    description: '対数正規分布',
    data: [
      ['値', '平均', '標準偏差', '累積', '確率'],
      [4, 3.5, 1.2, 'TRUE', '=LOGNORM.DIST(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'LOGNORM.INV',
    category: '02. 統計',
    description: '対数正規分布の逆関数',
    data: [
      ['確率', '平均', '標準偏差', '値'],
      [0.5, 3.5, 1.2, '=LOGNORM.INV(A2,B2,C2)']
    ]
  },
  {
    name: 'T.DIST.RT',
    category: '02. 統計',
    description: 'T分布（右側）',
    data: [
      ['値', '自由度', '確率'],
      [2, 10, '=T.DIST.RT(A2,B2)']
    ]
  },
  {
    name: 'CHISQ.DIST.RT',
    category: '02. 統計',
    description: 'カイ二乗分布（右側）',
    data: [
      ['値', '自由度', '確率'],
      [10, 5, '=CHISQ.DIST.RT(A2,B2)']
    ]
  },
  {
    name: 'CHISQ.INV.RT',
    category: '02. 統計',
    description: 'カイ二乗分布の逆関数（右側）',
    data: [
      ['確率', '自由度', '値'],
      [0.05, 5, '=CHISQ.INV.RT(A2,B2)']
    ]
  },
  {
    name: 'F.DIST.RT',
    category: '02. 統計',
    description: 'F分布（右側）',
    data: [
      ['値', '自由度1', '自由度2', '確率'],
      [2, 5, 10, '=F.DIST.RT(A2,B2,C2)']
    ]
  },
  {
    name: 'F.INV.RT',
    category: '02. 統計',
    description: 'F分布の逆関数（右側）',
    data: [
      ['確率', '自由度1', '自由度2', '値'],
      [0.05, 5, 10, '=F.INV.RT(A2,B2,C2)']
    ]
  },
  {
    name: 'BETA.DIST',
    category: '02. 統計',
    description: 'ベータ分布',
    data: [
      ['値', 'α', 'β', '下限', '上限', '累積', '確率'],
      [0.5, 2, 5, 0, 1, 'TRUE', '=BETA.DIST(A2,B2,C2,D2,E2,F2)']
    ]
  },
  {
    name: 'BETA.INV',
    category: '02. 統計',
    description: 'ベータ分布の逆関数',
    data: [
      ['確率', 'α', 'β', '下限', '上限', '値'],
      [0.5, 2, 5, 0, 1, '=BETA.INV(A2,B2,C2,D2,E2)']
    ]
  },
  {
    name: 'GAMMA.DIST',
    category: '02. 統計',
    description: 'ガンマ分布',
    data: [
      ['値', 'α', 'β', '累積', '確率'],
      [10, 9, 2, 'TRUE', '=GAMMA.DIST(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'GAMMA.INV',
    category: '02. 統計',
    description: 'ガンマ分布の逆関数',
    data: [
      ['確率', 'α', 'β', '値'],
      [0.5, 9, 2, '=GAMMA.INV(A2,B2,C2)']
    ]
  },
  {
    name: 'WEIBULL.DIST',
    category: '02. 統計',
    description: 'ワイブル分布',
    data: [
      ['値', 'α', 'β', '累積', '確率'],
      [100, 20, 100, 'TRUE', '=WEIBULL.DIST(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'PERCENTILE.EXC',
    category: '02. 統計',
    description: 'パーセンタイル（除外）',
    data: [
      ['データ', '', 'パーセンタイル'],
      [1, '', ''],
      [2, '', ''],
      [3, '', ''],
      [4, '', ''],
      [5, 0.25, '=PERCENTILE.EXC(A2:A6,B6)']
    ],
    expectedValues: { 'C6': 1.75 }
  },
  {
    name: 'PERCENTRANK.INC',
    category: '02. 統計',
    description: 'パーセントランク（含む）',
    data: [
      ['データ', '値', 'ランク'],
      [1, 3, '=PERCENTRANK.INC(A2:A6,B2)'],
      [2, '', ''],
      [3, '', ''],
      [4, '', ''],
      [5, '', '']
    ],
    expectedValues: { 'C2': 0.5 }
  },
  {
    name: 'PERCENTRANK.EXC',
    category: '02. 統計',
    description: 'パーセントランク（除外）',
    data: [
      ['データ', '値', 'ランク'],
      [1, 3, '=PERCENTRANK.EXC(A2:A6,B2)'],
      [2, '', ''],
      [3, '', ''],
      [4, '', ''],
      [5, '', '']
    ]
  },
  {
    name: 'QUARTILE.EXC',
    category: '02. 統計',
    description: '四分位数（除外）',
    data: [
      ['データ', '四分位', '値'],
      [1, 1, '=QUARTILE.EXC(A2:A9,B2)'],
      [2, 2, '=QUARTILE.EXC(A2:A9,B3)'],
      [3, 3, '=QUARTILE.EXC(A2:A9,B4)'],
      [4, '', ''],
      [5, '', ''],
      [6, '', ''],
      [7, '', ''],
      [8, '', '']
    ]
  },
  {
    name: 'STDEVA',
    category: '02. 統計',
    description: '標準偏差（テキスト含む）',
    data: [
      ['データ', '標準偏差'],
      [10, '=STDEVA(A2:A6)'],
      [20, ''],
      ['テキスト', ''],
      [30, ''],
      [40, '']
    ]
  },
  {
    name: 'STDEVPA',
    category: '02. 統計',
    description: '母標準偏差（テキスト含む）',
    data: [
      ['データ', '母標準偏差'],
      [10, '=STDEVPA(A2:A6)'],
      [20, ''],
      ['テキスト', ''],
      [30, ''],
      [40, '']
    ]
  },
  {
    name: 'VARA',
    category: '02. 統計',
    description: '分散（テキスト含む）',
    data: [
      ['データ', '分散'],
      [10, '=VARA(A2:A6)'],
      [20, ''],
      ['テキスト', ''],
      [30, ''],
      [40, '']
    ]
  },
  {
    name: 'VARPA',
    category: '02. 統計',
    description: '母分散（テキスト含む）',
    data: [
      ['データ', '母分散'],
      [10, '=VARPA(A2:A6)'],
      [20, ''],
      ['テキスト', ''],
      [30, ''],
      [40, '']
    ]
  }
];

// 文字列関数の個別テスト
export const textFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'CONCATENATE',
    category: '03. 文字列',
    description: '文字列を結合',
    data: [
      ['文字列1', '文字列2', '文字列3', '結果'],
      ['Hello', ' ', 'World', '=CONCATENATE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 'Hello World' }
  },
  {
    name: 'CONCAT',
    category: '03. 文字列',
    description: '文字列を結合（新版）',
    data: [
      ['文字列1', '文字列2', '結果'],
      ['ABC', 'DEF', '=CONCAT(A2,B2)']
    ],
    expectedValues: { 'C2': 'ABCDEF' }
  },
  {
    name: 'TEXTJOIN',
    category: '03. 文字列',
    description: '区切り文字で結合',
    data: [
      ['値1', '値2', '値3', '結果'],
      ['A', 'B', 'C', '=TEXTJOIN("-",TRUE,A2:C2)']
    ],
    expectedValues: { 'D2': 'A-B-C' }
  },
  {
    name: 'LEFT',
    category: '03. 文字列',
    description: '左から文字を抽出',
    data: [
      ['文字列', '文字数', '結果'],
      ['Hello World', 5, '=LEFT(A2,B2)']
    ],
    expectedValues: { 'C2': 'Hello' }
  },
  {
    name: 'RIGHT',
    category: '03. 文字列',
    description: '右から文字を抽出',
    data: [
      ['文字列', '文字数', '結果'],
      ['Hello World', 5, '=RIGHT(A2,B2)']
    ],
    expectedValues: { 'C2': 'World' }
  },
  {
    name: 'MID',
    category: '03. 文字列',
    description: '中間の文字を抽出',
    data: [
      ['文字列', '開始位置', '文字数', '結果'],
      ['Hello World', 3, 3, '=MID(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 'llo' }
  },
  {
    name: 'LEN',
    category: '03. 文字列',
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
    category: '03. 文字列',
    description: '文字位置を検索（大小区別）',
    data: [
      ['検索文字列', '対象文字列', '位置'],
      ['o', 'Hello World', '=FIND(A2,B2)']
    ],
    expectedValues: { 'C2': 5 }
  },
  {
    name: 'SEARCH',
    category: '03. 文字列',
    description: '文字位置を検索（大小区別なし）',
    data: [
      ['検索文字列', '対象文字列', '位置'],
      ['world', 'Hello World', '=SEARCH(A2,B2)']
    ],
    expectedValues: { 'C2': 7 }
  },
  {
    name: 'REPLACE',
    category: '03. 文字列',
    description: '文字を置換（位置指定）',
    data: [
      ['文字列', '開始位置', '文字数', '新文字列', '結果'],
      ['Hello World', 7, 5, 'Excel', '=REPLACE(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 'Hello Excel' }
  },
  {
    name: 'SUBSTITUTE',
    category: '03. 文字列',
    description: '文字を置換（文字指定）',
    data: [
      ['文字列', '検索文字列', '置換文字列', '結果'],
      ['Hello World', 'World', 'Excel', '=SUBSTITUTE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 'Hello Excel' }
  },
  {
    name: 'UPPER',
    category: '03. 文字列',
    description: '大文字に変換',
    data: [
      ['文字列', '結果'],
      ['Hello World', '=UPPER(A2)']
    ],
    expectedValues: { 'B2': 'HELLO WORLD' }
  },
  {
    name: 'LOWER',
    category: '03. 文字列',
    description: '小文字に変換',
    data: [
      ['文字列', '結果'],
      ['Hello World', '=LOWER(A2)']
    ],
    expectedValues: { 'B2': 'hello world' }
  },
  {
    name: 'PROPER',
    category: '03. 文字列',
    description: '先頭を大文字に変換',
    data: [
      ['文字列', '結果'],
      ['hello world', '=PROPER(A2)']
    ],
    expectedValues: { 'B2': 'Hello World' }
  },
  {
    name: 'TRIM',
    category: '03. 文字列',
    description: '余分なスペースを削除',
    data: [
      ['文字列', '結果'],
      ['  Hello   World  ', '=TRIM(A2)']
    ],
    expectedValues: { 'B2': 'Hello World' }
  },
  {
    name: 'TEXT',
    category: '03. 文字列',
    description: '数値を書式付き文字列に変換',
    data: [
      ['値', '表示形式', '結果'],
      [1234.56, '#,##0.00', '=TEXT(A2,B2)']
    ],
    expectedValues: { 'C2': '1,234.56' }
  },
  {
    name: 'VALUE',
    category: '03. 文字列',
    description: '文字列を数値に変換',
    data: [
      ['文字列', '結果'],
      ['123.45', '=VALUE(A2)']
    ],
    expectedValues: { 'B2': 123.45 }
  },
  {
    name: 'REPT',
    category: '03. 文字列',
    description: '文字列を繰り返す',
    data: [
      ['文字列', '繰り返し回数', '結果'],
      ['★', 5, '=REPT(A2,B2)']
    ],
    expectedValues: { 'C2': '★★★★★' }
  },
  {
    name: 'CHAR',
    category: '03. 文字列',
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
    category: '03. 文字列',
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
    category: '03. 文字列',
    description: '文字列が同一か判定',
    data: [
      ['文字列1', '文字列2', '結果'],
      ['Hello', 'Hello', '=EXACT(A2,B2)'],
      ['Hello', 'hello', '=EXACT(A3,B3)']
    ],
    expectedValues: { 'C2': true, 'C3': false }
  },
  {
    name: 'CLEAN',
    category: '03. 文字列',
    description: '印刷できない文字を削除',
    data: [
      ['テキスト', 'クリーン後'],
      ['Hello\x00World', '=CLEAN(A2)'],
      ['Test\x07String', '=CLEAN(A3)']
    ],
    expectedValues: { 'B2': 'HelloWorld', 'B3': 'TestString' }
  },
  {
    name: 'DOLLAR',
    category: '03. 文字列',
    description: '通貨形式に変換',
    data: [
      ['数値', '小数桁', '結果'],
      [1234.567, 2, '=DOLLAR(A2,B2)'],
      [1234.567, 0, '=DOLLAR(A3,B3)']
    ],
    expectedValues: { 'C2': '$1,234.57', 'C3': '$1,235' }
  },
  {
    name: 'FIXED',
    category: '03. 文字列',
    description: '小数点以下の桁数を固定',
    data: [
      ['数値', '小数桁', 'カンマなし', '結果'],
      [1234.567, 2, 'FALSE', '=FIXED(A2,B2,C2)'],
      [1234.567, 1, 'TRUE', '=FIXED(A3,B3,C3)']
    ],
    expectedValues: { 'D2': '1,234.57', 'D3': '1234.6' }
  },
  {
    name: 'NUMBERVALUE',
    category: '03. 文字列',
    description: 'テキストを数値に変換',
    data: [
      ['テキスト', '小数点', '桁区切り', '結果'],
      ['1,234.56', '.', ',', '=NUMBERVALUE(A2,B2,C2)'],
      ['1.234,56', ',', '.', '=NUMBERVALUE(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 1234.56, 'D3': 1234.56 }
  },
  {
    name: 'TEXTBEFORE',
    category: '03. 文字列',
    description: '指定文字の前のテキストを抽出',
    data: [
      ['テキスト', '区切り文字', '結果'],
      ['hello@example.com', '@', '=TEXTBEFORE(A2,B2)'],
      ['2024-01-15', '-', '=TEXTBEFORE(A3,B3)'],
      ['first,second,third', ',', '=TEXTBEFORE(A4,B4)']
    ],
    expectedValues: { 'C2': 'hello', 'C3': '2024', 'C4': 'first' }
  },
  {
    name: 'TEXTAFTER',
    category: '03. 文字列',
    description: '指定文字の後のテキストを抽出',
    data: [
      ['テキスト', '区切り文字', '結果'],
      ['hello@example.com', '@', '=TEXTAFTER(A2,B2)'],
      ['2024-01-15', '-', '=TEXTAFTER(A3,B3)'],
      ['first,second,third', ',', '=TEXTAFTER(A4,B4)']
    ],
    expectedValues: { 'C2': 'example.com', 'C3': '01-15', 'C4': 'second,third' }
  },
  {
    name: 'TEXTSPLIT',
    category: '03. 文字列',
    description: 'テキストを分割',
    data: [
      ['テキスト', '区切り文字', '分割結果'],
      ['apple,banana,orange', ',', '=TEXTSPLIT(A2,B2)'],
      ['2024-01-15', '-', '=TEXTSPLIT(A3,B3)'],
      ['one two three', ' ', '=TEXTSPLIT(A4,B4)']
    ]
  },
  {
    name: 'UNICHAR',
    category: '03. 文字列',
    description: 'Unicode番号から文字を返す',
    data: [
      ['Unicode番号', '文字'],
      [65, '=UNICHAR(A2)'],
      [8364, '=UNICHAR(A3)'],
      [12354, '=UNICHAR(A4)']
    ],
    expectedValues: { 'B2': 'A', 'B3': '€', 'B4': 'あ' }
  },
  {
    name: 'UNICODE',
    category: '03. 文字列',
    description: '文字からUnicode番号を返す',
    data: [
      ['文字', 'Unicode番号'],
      ['A', '=UNICODE(A2)'],
      ['€', '=UNICODE(A3)'],
      ['あ', '=UNICODE(A4)']
    ],
    expectedValues: { 'B2': 65, 'B3': 8364, 'B4': 12354 }
  },
  {
    name: 'T',
    category: '03. 文字列',
    description: 'テキストのみを返す',
    data: [
      ['値', 'テキスト'],
      ['Hello', '=T(A2)'],
      [123, '=T(A3)'],
      ['=TRUE()', '=T(A4)']
    ],
    expectedValues: { 'B2': 'Hello', 'B3': '', 'B4': '' }
  },
  {
    name: 'ASC',
    category: '03. 文字列',
    description: '全角を半角に変換',
    data: [
      ['全角文字', '半角文字'],
      ['ＨＥＬＬＯ', '=ASC(A2)'],
      ['１２３４５', '=ASC(A3)'],
      ['アイウエオ', '=ASC(A4)']
    ],
    expectedValues: { 'B2': 'HELLO', 'B3': '12345', 'B4': 'ｱｲｳｴｵ' }
  },
  {
    name: 'JIS',
    category: '03. 文字列',
    description: '半角を全角に変換',
    data: [
      ['半角文字', '全角文字'],
      ['HELLO', '=JIS(A2)'],
      ['12345', '=JIS(A3)'],
      ['ｱｲｳｴｵ', '=JIS(A4)']
    ],
    expectedValues: { 'B2': 'ＨＥＬＬＯ', 'B3': '１２３４５', 'B4': 'アイウエオ' }
  },
  {
    name: 'DBCS',
    category: '03. 文字列',
    description: '半角を全角に変換（DBCS）',
    data: [
      ['半角文字', '全角文字'],
      ['abc123', '=DBCS(A2)'],
      ['ｶﾀｶﾅ', '=DBCS(A3)']
    ],
    expectedValues: { 'B2': 'ａｂｃ１２３', 'B3': 'カタカナ' }
  },
  {
    name: 'LENB',
    category: '03. 文字列',
    description: 'バイト数を返す',
    data: [
      ['文字列', 'バイト数'],
      ['Hello', '=LENB(A2)'],
      ['こんにちは', '=LENB(A3)']
    ],
    expectedValues: { 'B2': 5, 'B3': 10 }
  },
  {
    name: 'FINDB',
    category: '03. 文字列',
    description: 'バイト位置を検索',
    data: [
      ['検索文字', '対象文字列', 'バイト位置'],
      ['o', 'Hello', '=FINDB(A2,B2)'],
      ['に', 'こんにちは', '=FINDB(A3,B3)']
    ],
    expectedValues: { 'C2': 5 }
  },
  {
    name: 'SEARCHB',
    category: '03. 文字列',
    description: 'バイト位置を検索（大小区別なし）',
    data: [
      ['検索文字', '対象文字列', 'バイト位置'],
      ['LO', 'Hello', '=SEARCHB(A2,B2)'],
      ['に', 'こんにちは', '=SEARCHB(A3,B3)']
    ]
  },
  {
    name: 'REPLACEB',
    category: '03. 文字列',
    description: 'バイト位置で置換',
    data: [
      ['文字列', '開始位置', 'バイト数', '新文字列', '結果'],
      ['Hello', 3, 2, 'XX', '=REPLACEB(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 'HeXXo' }
  }
];

// 日付・時刻関数の個別テスト
export const dateFunctionTests: IndividualFunctionTest[] = [
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
      ['開始日', '終了日', '単位', '期間'],
      ['2024/1/1', '2024/12/31', '"D"', '=DATEDIF(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 365 }
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
    ]
  },
  {
    name: 'EOMONTH',
    category: '04. 日付',
    description: '月末日を返す',
    data: [
      ['開始日', '月数', '月末日'],
      ['2024/2/15', 0, '=EOMONTH(A2,B2)']
    ]
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
    ]
  },
  {
    name: 'DATEVALUE',
    category: '04. 日付',
    description: '日付文字列を日付値に変換',
    data: [
      ['日付文字列', '日付値'],
      ['2024/12/25', '=DATEVALUE(A2)'],
      ['2024-01-15', '=DATEVALUE(A3)']
    ]
  },
  {
    name: 'TIMEVALUE',
    category: '04. 日付',
    description: '時刻文字列を時刻値に変換',
    data: [
      ['時刻文字列', '時刻値'],
      ['13:30:45', '=TIMEVALUE(A2)'],
      ['9:15 AM', '=TIMEVALUE(A3)']
    ]
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
    ]
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
    ]
  },
  {
    name: 'WORKDAY.INTL',
    category: '04. 日付',
    description: '営業日後の日付（国際版）',
    data: [
      ['開始日', '営業日数', '週末', '結果日'],
      ['2024/1/1', 10, 1, '=WORKDAY.INTL(A2,B2,C2)'],
      ['2024/1/1', 5, '"0000011"', '=WORKDAY.INTL(A3,B3,C3)']
    ]
  }
];

// 論理関数の個別テスト
export const logicalFunctionTests: IndividualFunctionTest[] = [
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
      ['検索値', '結果'],
      ['存在しない', '=IFNA(VLOOKUP(A2,D:E,2,FALSE),"見つかりません")']
    ],
    expectedValues: { 'B2': '見つかりません' }
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

// 検索・参照関数の個別テスト
export const lookupFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'VLOOKUP',
    category: '06. 検索',
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
    category: '06. 検索',
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
    category: '06. 検索',
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
    category: '06. 検索',
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
    category: '06. 検索',
    description: '値の位置を検索',
    data: [
      ['値', '', '', '検索値', '位置'],
      [10, 20, 30, 20, '=MATCH(D2,A2:C2,0)']
    ],
    expectedValues: { 'E2': 2 }
  },
  {
    name: 'CHOOSE',
    category: '06. 検索',
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
    category: '06. 検索',
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
    category: '06. 検索',
    description: '文字列を参照に変換',
    data: [
      ['値1', '値2', '', '参照文字列', '結果'],
      [100, 200, '', '"A2"', '=INDIRECT(D2)']
    ],
    expectedValues: { 'E2': 100 }
  },
  {
    name: 'ROW',
    category: '06. 検索',
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
    category: '06. 検索',
    description: '列番号を返す',
    data: [
      ['列1', '列2', '列3'],
      ['=COLUMN()', '=COLUMN()', '=COLUMN()']
    ],
    expectedValues: { 'A2': 1, 'B2': 2, 'C2': 3 }
  },
  {
    name: 'XMATCH',
    category: '06. 検索',
    description: '配列内の項目の位置を返す',
    data: [
      ['データ', '検索値', '位置'],
      ['りんご', 'オレンジ', '=XMATCH(B2,A2:A5)'],
      ['バナナ', '', ''],
      ['オレンジ', '', ''],
      ['ぶどう', '', '']
    ],
    expectedValues: { 'C2': 3 }
  },
  {
    name: 'ROWS',
    category: '06. 検索',
    description: '範囲の行数を返す',
    data: [
      ['データ', '', '行数'],
      [1, 2, '=ROWS(A2:B5)'],
      [3, 4, ''],
      [5, 6, ''],
      [7, 8, '']
    ],
    expectedValues: { 'C2': 4 }
  },
  {
    name: 'COLUMNS',
    category: '06. 検索',
    description: '範囲の列数を返す',
    data: [
      ['データ', '', '', '', '列数'],
      [1, 2, 3, 4, '=COLUMNS(A2:D2)']
    ],
    expectedValues: { 'E2': 4 }
  },
  {
    name: 'ADDRESS',
    category: '06. 検索',
    description: 'セルアドレスを文字列で返す',
    data: [
      ['行番号', '列番号', '絶対参照', 'アドレス'],
      [2, 3, 1, '=ADDRESS(A2,B2,C2)'],
      [5, 1, 4, '=ADDRESS(A3,B3,C3)']
    ],
    expectedValues: { 'D2': '$C$2', 'D3': 'A5' }
  },
  {
    name: 'HYPERLINK',
    category: '06. 検索',
    description: 'ハイパーリンクを作成',
    data: [
      ['URL', '表示テキスト', 'リンク'],
      ['https://example.com', 'Example', '=HYPERLINK(A2,B2)']
    ]
  },
  {
    name: 'LOOKUP',
    category: '06. 検索',
    description: 'ベクトル検索',
    data: [
      ['検索値', '検索範囲', '', '結果範囲', '', '結果'],
      [4.5, 1, 3, 'A', 'B', '=LOOKUP(A2,B2:C2,D2:E2)'],
      ['', 4.5, 6, 'C', 'D', '']
    ],
    expectedValues: { 'F2': 'B' }
  },
  {
    name: 'AREAS',
    category: '06. 検索',
    description: '参照の領域数',
    data: [
      ['数式', '領域数'],
      ['=A1:B2', '=AREAS(A2)'],
      ['=(A1:B2,D1:E2)', '=AREAS((A1:B2,D1:E2))']
    ],
    expectedValues: { 'B2': 1 }
  },
  {
    name: 'FORMULATEXT',
    category: '06. 検索',
    description: '数式を文字列として返す',
    data: [
      ['値', '数式テキスト'],
      ['=SUM(1,2,3)', '=FORMULATEXT(A2)'],
      [100, '=FORMULATEXT(A3)']
    ],
    expectedValues: { 'B2': '=SUM(1,2,3)' }
  },
  {
    name: 'GETPIVOTDATA',
    category: '06. 検索',
    description: 'ピボットテーブルからデータ取得',
    data: [
      ['製品', '地域', '売上'],
      ['A', '東', 100],
      ['A', '西', 150],
      ['B', '東', 200],
      ['', '', ''],
      ['取得値', '=GETPIVOTDATA("売上",A1,"製品","A","地域","東")', '']
    ],
    expectedValues: { 'B6': 100 }
  }
];

// 財務関数の個別テスト
export const financialFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'PMT',
    category: '07. 財務',
    description: 'ローン定期支払額',
    data: [
      ['元本', '年利率', '期間(月)', '月額支払額'],
      [100000, 0.06, 12, '=PMT(B2/12,C2,-A2)']
    ]
  },
  {
    name: 'FV',
    category: '07. 財務',
    description: '将来価値',
    data: [
      ['定期支払額', '利率', '期間', '将来価値'],
      [1000, 0.05, 10, '=FV(B2/12,C2,-A2)']
    ]
  },
  {
    name: 'PV',
    category: '07. 財務',
    description: '現在価値',
    data: [
      ['定期支払額', '利率', '期間', '現在価値'],
      [1000, 0.05, 10, '=PV(B2/12,C2,-A2)']
    ]
  },
  {
    name: 'RATE',
    category: '07. 財務',
    description: '利率',
    data: [
      ['期間', '支払額', '現在価値', '利率'],
      [12, -1000, 10000, '=RATE(A2,B2,C2)*12']
    ]
  },
  {
    name: 'NPER',
    category: '07. 財務',
    description: '支払回数',
    data: [
      ['利率', '支払額', '現在価値', '期間'],
      [0.05, -1000, 10000, '=NPER(A2/12,B2,C2)']
    ]
  },
  {
    name: 'NPV',
    category: '07. 財務',
    description: '正味現在価値',
    data: [
      ['割引率', 'CF1', 'CF2', 'CF3', 'NPV'],
      [0.1, -10000, 3000, 4200, '=NPV(A2,B2:D2)']
    ]
  },
  {
    name: 'IRR',
    category: '07. 財務',
    description: '内部収益率',
    data: [
      ['初期投資', 'CF1', 'CF2', 'CF3', 'IRR'],
      [-10000, 3000, 4200, 6800, '=IRR(A2:D2)']
    ]
  },
  {
    name: 'XNPV',
    category: '07. 財務',
    description: '不定期のキャッシュフローの正味現在価値',
    data: [
      ['割引率', 0.1, '', ''],
      ['日付', 'キャッシュフロー', '', 'XNPV'],
      ['2024/1/1', -10000, '', '=XNPV(B1,B3:B6,A3:A6)'],
      ['2024/4/1', 3000, '', ''],
      ['2024/7/1', 4200, '', ''],
      ['2024/10/1', 6800, '', '']
    ]
  },
  {
    name: 'XIRR',
    category: '07. 財務',
    description: '不定期のキャッシュフローの内部収益率',
    data: [
      ['日付', 'キャッシュフロー', 'XIRR'],
      ['2024/1/1', -10000, '=XIRR(B2:B5,A2:A5)'],
      ['2024/4/1', 3000, ''],
      ['2024/7/1', 4200, ''],
      ['2024/10/1', 6800, '']
    ]
  },
  {
    name: 'IPMT',
    category: '07. 財務',
    description: '利息部分の支払額',
    data: [
      ['利率', '期', '期間', '現在価値', '利息支払額'],
      [0.06, 1, 12, -100000, '=IPMT(A2/12,B2,C2,D2)']
    ]
  },
  {
    name: 'PPMT',
    category: '07. 財務',
    description: '元本部分の支払額',
    data: [
      ['利率', '期', '期間', '現在価値', '元本支払額'],
      [0.06, 1, 12, -100000, '=PPMT(A2/12,B2,C2,D2)']
    ]
  },
  {
    name: 'MIRR',
    category: '07. 財務',
    description: '修正内部収益率',
    data: [
      ['CF0', 'CF1', 'CF2', 'CF3', '投資利率', '再投資利率', 'MIRR'],
      [-10000, 3000, 4200, 6800, 0.1, 0.12, '=MIRR(A2:D2,E2,F2)']
    ]
  },
  {
    name: 'SLN',
    category: '07. 財務',
    description: '定額法による減価償却',
    data: [
      ['取得価額', '残存価額', '耐用年数', '減価償却費'],
      [100000, 10000, 5, '=SLN(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 18000 }
  },
  {
    name: 'SYD',
    category: '07. 財務',
    description: '級数法による減価償却',
    data: [
      ['取得価額', '残存価額', '耐用年数', '期', '減価償却費'],
      [100000, 10000, 5, 1, '=SYD(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'DDB',
    category: '07. 財務',
    description: '倍率法による減価償却',
    data: [
      ['取得価額', '残存価額', '耐用年数', '期', '倍率', '減価償却費'],
      [100000, 10000, 5, 1, 2, '=DDB(A2,B2,C2,D2,E2)']
    ]
  },
  {
    name: 'DB',
    category: '07. 財務',
    description: '定率法による減価償却',
    data: [
      ['取得価額', '残存価額', '耐用年数', '期', '月', '減価償却費'],
      [100000, 10000, 5, 1, 12, '=DB(A2,B2,C2,D2,E2)']
    ]
  },
  {
    name: 'VDB',
    category: '07. 財務',
    description: '倍率法減価償却（期間指定）',
    data: [
      ['取得価額', '残存価額', '耐用年数', '開始期', '終了期', '倍率', '切替なし', '償却費'],
      [100000, 10000, 5, 0, 1, 2, 'FALSE', '=VDB(A2,B2,C2,D2,E2,F2,G2)']
    ]
  },
  // 債券関連関数
  {
    name: 'ACCRINT',
    category: '07. 財務',
    description: '定期利息証券の未収利息額',
    data: [
      ['発行日', '初回利払日', '決済日', '利率', '額面', '頻度', '基準', '未収利息'],
      ['2024/1/1', '2024/7/1', '2024/4/1', 0.05, 1000, 2, 0, '=ACCRINT(A2,B2,C2,D2,E2,F2,G2)']
    ]
  },
  {
    name: 'ACCRINTM',
    category: '07. 財務',
    description: '満期一括払い証券の未収利息額',
    data: [
      ['発行日', '満期日', '利率', '額面', '基準', '未収利息'],
      ['2024/1/1', '2024/12/31', 0.05, 1000, 0, '=ACCRINTM(A2,B2,C2,D2,E2)']
    ]
  },
  {
    name: 'DISC',
    category: '07. 財務',
    description: '証券の割引率',
    data: [
      ['決済日', '満期日', '価格', '償還価格', '基準', '割引率'],
      ['2024/1/1', '2024/12/31', 95, 100, 0, '=DISC(A2,B2,C2,D2,E2)']
    ]
  },
  {
    name: 'DURATION',
    category: '07. 財務',
    description: '定期利息証券の年間マコーレー・デュレーション',
    data: [
      ['決済日', '満期日', '利率', '利回り', '頻度', '基準', 'デュレーション'],
      ['2024/1/1', '2029/1/1', 0.05, 0.06, 2, 0, '=DURATION(A2,B2,C2,D2,E2,F2)']
    ]
  },
  {
    name: 'MDURATION',
    category: '07. 財務',
    description: '修正マコーレー・デュレーション',
    data: [
      ['決済日', '満期日', '利率', '利回り', '頻度', '基準', '修正デュレーション'],
      ['2024/1/1', '2029/1/1', 0.05, 0.06, 2, 0, '=MDURATION(A2,B2,C2,D2,E2,F2)']
    ]
  },
  {
    name: 'PRICE',
    category: '07. 財務',
    description: '定期利息証券の価格',
    data: [
      ['決済日', '満期日', '利率', '利回り', '償還価格', '頻度', '基準', '価格'],
      ['2024/1/1', '2029/1/1', 0.05, 0.06, 100, 2, 0, '=PRICE(A2,B2,C2,D2,E2,F2,G2)']
    ]
  },
  {
    name: 'YIELD',
    category: '07. 財務',
    description: '定期利息証券の利回り',
    data: [
      ['決済日', '満期日', '利率', '価格', '償還価格', '頻度', '基準', '利回り'],
      ['2024/1/1', '2029/1/1', 0.05, 95, 100, 2, 0, '=YIELD(A2,B2,C2,D2,E2,F2,G2)']
    ]
  },
  // クーポン関連関数
  {
    name: 'COUPDAYBS',
    category: '07. 財務',
    description: '利払期の開始日から決済日までの日数',
    data: [
      ['決済日', '満期日', '頻度', '基準', '日数'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPDAYBS(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'COUPDAYS',
    category: '07. 財務',
    description: '決済日を含む利払期の日数',
    data: [
      ['決済日', '満期日', '頻度', '基準', '日数'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPDAYS(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'COUPDAYSNC',
    category: '07. 財務',
    description: '決済日から次回利払日までの日数',
    data: [
      ['決済日', '満期日', '頻度', '基準', '日数'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPDAYSNC(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'COUPNCD',
    category: '07. 財務',
    description: '決済日後の次回利払日',
    data: [
      ['決済日', '満期日', '頻度', '基準', '次回利払日'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPNCD(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'COUPNUM',
    category: '07. 財務',
    description: '決済日から満期日までの利払回数',
    data: [
      ['決済日', '満期日', '頻度', '基準', '利払回数'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPNUM(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'COUPPCD',
    category: '07. 財務',
    description: '決済日直前の利払日',
    data: [
      ['決済日', '満期日', '頻度', '基準', '前回利払日'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPPCD(A2,B2,C2,D2)']
    ]
  },
  // 減価償却関連関数
  {
    name: 'AMORDEGRC',
    category: '07. 財務',
    description: '減価償却費（フランス会計システム）',
    data: [
      ['取得価額', '購入日', '最初の期末', '残存価額', '期', '率', '基準', '減価償却費'],
      [10000, '2024/1/1', '2024/12/31', 1000, 1, 0.15, 1, '=AMORDEGRC(A2,B2,C2,D2,E2,F2,G2)']
    ]
  },
  {
    name: 'AMORLINC',
    category: '07. 財務',
    description: '減価償却費（フランス会計システム、線形）',
    data: [
      ['取得価額', '購入日', '最初の期末', '残存価額', '期', '率', '基準', '減価償却費'],
      [10000, '2024/1/1', '2024/12/31', 1000, 1, 0.15, 1, '=AMORLINC(A2,B2,C2,D2,E2,F2,G2)']
    ]
  },
  {
    name: 'CUMIPMT',
    category: '07. 財務',
    description: '累計利息支払額',
    data: [
      ['利率', '期間', '現在価値', '開始期', '終了期', 'タイプ', '累計利息'],
      [0.06, 36, 100000, 1, 12, 0, '=CUMIPMT(A2/12,B2,C2,D2,E2,F2)']
    ]
  },
  {
    name: 'CUMPRINC',
    category: '07. 財務',
    description: '累計元本支払額',
    data: [
      ['利率', '期間', '現在価値', '開始期', '終了期', 'タイプ', '累計元本'],
      [0.06, 36, 100000, 1, 12, 0, '=CUMPRINC(A2/12,B2,C2,D2,E2,F2)']
    ]
  },
  // 特殊債券関連関数
  {
    name: 'ODDFPRICE',
    category: '07. 財務',
    description: '期間が半端な最初の期の証券価格',
    data: [
      ['決済日', '満期日', '発行日', '初回利払日', '利率', '利回り', '償還価格', '頻度', '基準', '価格'],
      ['2024/2/15', '2029/1/1', '2024/1/1', '2024/7/1', 0.05, 0.06, 100, 2, 0, '=ODDFPRICE(A2,B2,C2,D2,E2,F2,G2,H2,I2)']
    ]
  },
  {
    name: 'ODDFYIELD',
    category: '07. 財務',
    description: '期間が半端な最初の期の証券利回り',
    data: [
      ['決済日', '満期日', '発行日', '初回利払日', '利率', '価格', '償還価格', '頻度', '基準', '利回り'],
      ['2024/2/15', '2029/1/1', '2024/1/1', '2024/7/1', 0.05, 95, 100, 2, 0, '=ODDFYIELD(A2,B2,C2,D2,E2,F2,G2,H2,I2)']
    ]
  },
  {
    name: 'ODDLPRICE',
    category: '07. 財務',
    description: '期間が半端な最後の期の証券価格',
    data: [
      ['決済日', '満期日', '最終利払日', '利率', '利回り', '償還価格', '頻度', '基準', '価格'],
      ['2024/1/1', '2024/10/15', '2024/7/1', 0.05, 0.06, 100, 2, 0, '=ODDLPRICE(A2,B2,C2,D2,E2,F2,G2,H2)']
    ]
  },
  {
    name: 'ODDLYIELD',
    category: '07. 財務',
    description: '期間が半端な最後の期の証券利回り',
    data: [
      ['決済日', '満期日', '最終利払日', '利率', '価格', '償還価格', '頻度', '基準', '利回り'],
      ['2024/1/1', '2024/10/15', '2024/7/1', 0.05, 95, 100, 2, 0, '=ODDLYIELD(A2,B2,C2,D2,E2,F2,G2,H2)']
    ]
  },
  // 財務省証券関連関数
  {
    name: 'TBILLEQ',
    category: '07. 財務',
    description: '財務省短期証券の債券換算利回り',
    data: [
      ['決済日', '満期日', '割引率', '債券換算利回り'],
      ['2024/1/1', '2024/6/30', 0.045, '=TBILLEQ(A2,B2,C2)']
    ]
  },
  {
    name: 'TBILLPRICE',
    category: '07. 財務',
    description: '財務省短期証券の価格',
    data: [
      ['決済日', '満期日', '割引率', '価格'],
      ['2024/1/1', '2024/6/30', 0.045, '=TBILLPRICE(A2,B2,C2)']
    ]
  },
  {
    name: 'TBILLYIELD',
    category: '07. 財務',
    description: '財務省短期証券の利回り',
    data: [
      ['決済日', '満期日', '価格', '利回り'],
      ['2024/1/1', '2024/6/30', 97.75, '=TBILLYIELD(A2,B2,C2)']
    ]
  },
  // 通貨変換関連関数
  {
    name: 'DOLLARDE',
    category: '07. 財務',
    description: '分数表記のドル価格を小数表記に変換',
    data: [
      ['分数価格', '分母', '小数価格'],
      [1.02, 16, '=DOLLARDE(A2,B2)'],
      [1.1, 32, '=DOLLARDE(A3,B3)']
    ]
  },
  {
    name: 'DOLLARFR',
    category: '07. 財務',
    description: '小数表記のドル価格を分数表記に変換',
    data: [
      ['小数価格', '分母', '分数価格'],
      [1.125, 16, '=DOLLARFR(A2,B2)'],
      [1.3125, 32, '=DOLLARFR(A3,B3)']
    ]
  },
  // 利率変換関連関数
  {
    name: 'EFFECT',
    category: '07. 財務',
    description: '実効年利率',
    data: [
      ['名目利率', '年間複利回数', '実効利率'],
      [0.06, 4, '=EFFECT(A2,B2)'],
      [0.06, 12, '=EFFECT(A3,B3)']
    ]
  },
  {
    name: 'NOMINAL',
    category: '07. 財務',
    description: '名目年利率',
    data: [
      ['実効利率', '年間複利回数', '名目利率'],
      [0.0614, 4, '=NOMINAL(A2,B2)'],
      [0.0618, 12, '=NOMINAL(A3,B3)']
    ]
  },
  // その他の証券関連関数
  {
    name: 'PRICEDISC',
    category: '07. 財務',
    description: '割引証券の価格',
    data: [
      ['決済日', '満期日', '割引率', '償還価格', '基準', '価格'],
      ['2024/1/1', '2024/12/31', 0.05, 100, 0, '=PRICEDISC(A2,B2,C2,D2,E2)']
    ]
  },
  {
    name: 'RECEIVED',
    category: '07. 財務',
    description: '満期保有証券の受取額',
    data: [
      ['決済日', '満期日', '投資額', '割引率', '基準', '受取額'],
      ['2024/1/1', '2024/12/31', 95, 0.05, 0, '=RECEIVED(A2,B2,C2,D2,E2)']
    ]
  },
  {
    name: 'INTRATE',
    category: '07. 財務',
    description: '満期保有証券の利率',
    data: [
      ['決済日', '満期日', '投資額', '償還価格', '基準', '利率'],
      ['2024/1/1', '2024/12/31', 95, 100, 0, '=INTRATE(A2,B2,C2,D2,E2)']
    ]
  },
  {
    name: 'PRICEMAT',
    category: '07. 財務',
    description: '満期利息付証券の価格',
    data: [
      ['決済日', '満期日', '発行日', '利率', '利回り', '基準', '価格'],
      ['2024/1/1', '2024/12/31', '2023/1/1', 0.05, 0.06, 0, '=PRICEMAT(A2,B2,C2,D2,E2,F2)']
    ]
  },
  {
    name: 'YIELDDISC',
    category: '07. 財務',
    description: '割引証券の年利回り',
    data: [
      ['決済日', '満期日', '価格', '償還価格', '基準', '利回り'],
      ['2024/1/1', '2024/12/31', 95, 100, 0, '=YIELDDISC(A2,B2,C2,D2,E2)']
    ]
  },
  {
    name: 'YIELDMAT',
    category: '07. 財務',
    description: '満期利息付証券の年利回り',
    data: [
      ['決済日', '満期日', '発行日', '利率', '価格', '基準', '利回り'],
      ['2024/1/1', '2024/12/31', '2023/1/1', 0.05, 95, 0, '=YIELDMAT(A2,B2,C2,D2,E2,F2)']
    ]
  }
];

// 行列関数の個別テスト
export const matrixFunctionTests: IndividualFunctionTest[] = [
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

// 情報関数の個別テスト
export const informationFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'ISBLANK',
    category: '09. 情報',
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
    category: '09. 情報',
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
    category: '09. 情報',
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
    category: '09. 情報',
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
    category: '09. 情報',
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
    category: '09. 情報',
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
    category: '09. 情報',
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
    category: '09. 情報',
    description: '数値に変換',
    data: [
      ['値', '数値変換'],
      [7, '=N(A2)'],
      ['7', '=N(A3)'],
      ['=TRUE()', '=N(A4)']
    ],
    expectedValues: { 'B2': 7, 'B3': 0, 'B4': 1 }
  },
  {
    name: 'ERROR.TYPE',
    category: '09. 情報',
    description: 'エラーの種類',
    data: [
      ['エラー', 'エラー番号'],
      ['=1/0', '=ERROR.TYPE(A2)'],
      ['=NA()', '=ERROR.TYPE(A3)']
    ],
    expectedValues: { 'B2': 2, 'B3': 7 }
  },
  {
    name: 'SHEET',
    category: '09. 情報',
    description: 'シート番号を返す',
    data: [
      ['シート番号'],
      ['=SHEET()']
    ]
  },
  {
    name: 'SHEETS',
    category: '09. 情報',
    description: 'シート数を返す',
    data: [
      ['シート数'],
      ['=SHEETS()']
    ]
  },
  {
    name: 'ISFORMULA',
    category: '09. 情報',
    description: '数式か判定',
    data: [
      ['値', '数式判定'],
      ['=A1+1', '=ISFORMULA(A2)'],
      [100, '=ISFORMULA(A3)']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'ISREF',
    category: '09. 情報',
    description: '参照か判定',
    data: [
      ['値', '参照判定'],
      ['A1', '=ISREF(A1)'],
      ['テキスト', '=ISREF("テキスト")']
    ],
    expectedValues: { 'B2': true, 'B3': false }
  },
  {
    name: 'CELL',
    category: '09. 情報',
    description: 'セル情報を取得',
    data: [
      ['情報タイプ', '参照', '結果'],
      ['type', 'A2', '=CELL(A2,B2)'],
      ['address', 'B3', '=CELL(A3,B3)']
    ]
  },
  {
    name: 'INFO',
    category: '09. 情報',
    description: 'システム情報',
    data: [
      ['情報タイプ', '結果'],
      ['numfile', '=INFO(A2)'],
      ['osversion', '=INFO(A3)']
    ]
  },
  {
    name: 'ISBETWEEN',
    category: '09. 情報',
    description: '値が範囲内か判定',
    data: [
      ['値', '下限', '上限', '範囲内判定'],
      [5, 1, 10, '=ISBETWEEN(A2,B2,C2)'],
      [15, 1, 10, '=ISBETWEEN(A3,B3,C3)'],
      [10, 1, 10, '=ISBETWEEN(A4,B4,C4)']
    ],
    expectedValues: { 'D2': true, 'D3': false, 'D4': true }
  }
];

// データベース関数の個別テスト
export const databaseFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'DSUM',
    category: '10. データベース',
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
    category: '10. データベース',
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
    category: '10. データベース',
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
  },
  {
    name: 'DCOUNTA',
    category: '10. データベース',
    description: '条件付き非空白カウント',
    data: [
      ['名前', '部署', '売上'],
      ['田中', '営業', 100],
      ['佐藤', '営業', ''],
      ['鈴木', '開発', 120],
      ['', '', ''],
      ['条件', '', ''],
      ['部署', '', ''],
      ['営業', '', '=DCOUNTA(A1:C4,A1,A6:A7)']
    ],
    expectedValues: { 'C7': 2 }
  },
  {
    name: 'DMAX',
    category: '10. データベース',
    description: '条件付き最大値',
    data: [
      ['名前', '部署', '売上'],
      ['田中', '営業', 100],
      ['佐藤', '営業', 150],
      ['鈴木', '開発', 120],
      ['', '', ''],
      ['条件', '', ''],
      ['部署', '', ''],
      ['営業', '', '=DMAX(A1:C4,C1,A6:A7)']
    ],
    expectedValues: { 'C7': 150 }
  },
  {
    name: 'DMIN',
    category: '10. データベース',
    description: '条件付き最小値',
    data: [
      ['名前', '部署', '売上'],
      ['田中', '営業', 100],
      ['佐藤', '営業', 150],
      ['鈴木', '開発', 120],
      ['', '', ''],
      ['条件', '', ''],
      ['部署', '', ''],
      ['営業', '', '=DMIN(A1:C4,C1,A6:A7)']
    ],
    expectedValues: { 'C7': 100 }
  },
  {
    name: 'DPRODUCT',
    category: '10. データベース',
    description: '条件付き積',
    data: [
      ['商品', 'カテゴリ', '数量'],
      ['A', '文具', 2],
      ['B', '文具', 3],
      ['C', '家電', 4],
      ['', '', ''],
      ['条件', '', ''],
      ['カテゴリ', '', ''],
      ['文具', '', '=DPRODUCT(A1:C4,C1,A6:A7)']
    ],
    expectedValues: { 'C7': 6 }
  },
  {
    name: 'DSTDEV',
    category: '10. データベース',
    description: '条件付き標準偏差',
    data: [
      ['名前', '部署', '売上'],
      ['田中', '営業', 100],
      ['佐藤', '営業', 150],
      ['鈴木', '営業', 200],
      ['', '', ''],
      ['条件', '', ''],
      ['部署', '', ''],
      ['営業', '', '=DSTDEV(A1:C4,C1,A6:A7)']
    ]
  },
  {
    name: 'DSTDEVP',
    category: '10. データベース',
    description: '条件付き母標準偏差',
    data: [
      ['名前', '部署', '売上'],
      ['田中', '営業', 100],
      ['佐藤', '営業', 150],
      ['鈴木', '営業', 200],
      ['', '', ''],
      ['条件', '', ''],
      ['部署', '', ''],
      ['営業', '', '=DSTDEVP(A1:C4,C1,A6:A7)']
    ]
  },
  {
    name: 'DVAR',
    category: '10. データベース',
    description: '条件付き分散',
    data: [
      ['名前', '部署', '売上'],
      ['田中', '営業', 100],
      ['佐藤', '営業', 150],
      ['鈴木', '営業', 200],
      ['', '', ''],
      ['条件', '', ''],
      ['部署', '', ''],
      ['営業', '', '=DVAR(A1:C4,C1,A6:A7)']
    ]
  },
  {
    name: 'DVARP',
    category: '10. データベース',
    description: '条件付き母分散',
    data: [
      ['名前', '部署', '売上'],
      ['田中', '営業', 100],
      ['佐藤', '営業', 150],
      ['鈴木', '営業', 200],
      ['', '', ''],
      ['条件', '', ''],
      ['部署', '', ''],
      ['営業', '', '=DVARP(A1:C4,C1,A6:A7)']
    ]
  },
  {
    name: 'DGET',
    category: '10. データベース',
    description: '条件に一致する単一値',
    data: [
      ['ID', '名前', '部署'],
      [101, '田中', '営業'],
      [102, '佐藤', '技術'],
      [103, '鈴木', '人事'],
      ['', '', ''],
      ['条件', '', ''],
      ['ID', '', ''],
      [102, '', '=DGET(A1:C4,B1,A6:A7)']
    ],
    expectedValues: { 'C7': '佐藤' }
  }
];

// エンジニアリング関数の個別テスト
export const engineeringFunctionTests: IndividualFunctionTest[] = [
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
  // ベッセル関数
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
  // ビット演算関数
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
  // エラー関数
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

// 動的配列関数の個別テスト
export const dynamicArrayFunctionTests: IndividualFunctionTest[] = [
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

// キューブ関数の個別テスト
export const cubeFunctionTests: IndividualFunctionTest[] = [
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

// Web関数の個別テスト
export const webFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'ENCODEURL',
    category: '14. Web・インポート',
    description: 'URLエンコード',
    data: [
      ['テキスト', 'エンコード結果'],
      ['Hello World', '=ENCODEURL(A2)'],
      ['こんにちは', '=ENCODEURL(A3)']
    ],
    expectedValues: { 'B2': 'Hello%20World' }
  },
  {
    name: 'SPLIT',
    category: '14. Web・インポート',
    description: 'テキストを分割',
    data: [
      ['テキスト', '区切り文字', '分割結果'],
      ['apple,banana,orange', ',', '=SPLIT(A2,B2)'],
      ['one-two-three', '-', '=SPLIT(A3,B3)']
    ]
  },
  {
    name: 'JOIN',
    category: '14. Web・インポート',
    description: 'テキストを結合',
    data: [
      ['区切り文字', '値1', '値2', '値3', '結合結果'],
      [',', 'apple', 'banana', 'orange', '=JOIN(A2,B2:D2)'],
      ['-', 'one', 'two', 'three', '=JOIN(A3,B3:D3)']
    ],
    expectedValues: { 'E2': 'apple,banana,orange', 'E3': 'one-two-three' }
  },
  {
    name: 'QUERY',
    category: '14. Web・インポート',
    description: 'データのクエリ',
    data: [
      ['名前', '年齢', '部署', '', 'クエリ結果'],
      ['田中', 25, '営業', '', '=QUERY(A2:C4,"SELECT A, B WHERE B > 25")'],
      ['佐藤', 30, '技術', '', ''],
      ['鈴木', 28, '人事', '', '']
    ]
  },
  {
    name: 'FLATTEN',
    category: '14. Web・インポート',
    description: '多次元配列を1次元に',
    data: [
      ['配列1', '', '配列2', '', 'フラット化'],
      [1, 2, 5, 6, '=FLATTEN(A2:D3)'],
      [3, 4, 7, 8, '']
    ]
  },
  {
    name: 'ARRAYFORMULA',
    category: '14. Web・インポート',
    description: '配列数式',
    data: [
      ['値', '結果'],
      [1, '=ARRAYFORMULA(A2:A5*2)'],
      [2, ''],
      [3, ''],
      [4, '']
    ],
    expectedValues: { 'B2': 2, 'B3': 4, 'B4': 6, 'B5': 8 }
  },
  {
    name: 'REGEXEXTRACT',
    category: '14. Web・インポート',
    description: '正規表現で抽出',
    data: [
      ['テキスト', 'パターン', '抽出結果'],
      ['abc123def', '[0-9]+', '=REGEXEXTRACT(A2,B2)'],
      ['test@example.com', '[^@]+', '=REGEXEXTRACT(A3,B3)']
    ],
    expectedValues: { 'C2': '123', 'C3': 'test' }
  },
  {
    name: 'REGEXMATCH',
    category: '14. Web・インポート',
    description: '正規表現でマッチ',
    data: [
      ['テキスト', 'パターン', 'マッチ'],
      ['abc123', '[0-9]+', '=REGEXMATCH(A2,B2)'],
      ['abcdef', '[0-9]+', '=REGEXMATCH(A3,B3)']
    ],
    expectedValues: { 'C2': true, 'C3': false }
  },
  {
    name: 'REGEXREPLACE',
    category: '14. Web・インポート',
    description: '正規表現で置換',
    data: [
      ['テキスト', 'パターン', '置換文字', '結果'],
      ['abc123def', '[0-9]+', 'XXX', '=REGEXREPLACE(A2,B2,C2)'],
      ['hello world', ' ', '_', '=REGEXREPLACE(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 'abcXXXdef', 'D3': 'hello_world' }
  },
  {
    name: 'SORTN',
    category: '14. Web・インポート',
    description: '上位N件をソート',
    data: [
      ['値', 'ソート結果'],
      [85, '=SORTN(A2:A6,3)'],
      [92, ''],
      [78, ''],
      [95, ''],
      [88, '']
    ]
  },
  {
    name: 'WEBSERVICE',
    category: '14. Web・インポート',
    description: 'Webサービスからデータ取得',
    data: [
      ['URL', '取得結果'],
      ['https://api.example.com/data', '=WEBSERVICE(A2)']
    ]
  },
  {
    name: 'FILTERXML',
    category: '14. Web・インポート',
    description: 'XMLからデータ抽出',
    data: [
      ['XMLデータ', 'XPath', '抽出結果'],
      ['<root><item>Value1</item><item>Value2</item></root>', '//item[1]', '=FILTERXML(A2,B2)']
    ],
    expectedValues: { 'C2': 'Value1' }
  },
  {
    name: 'SPARKLINE',
    category: '14. Web・インポート',
    description: 'ミニグラフを作成',
    data: [
      ['データ', '', '', '', 'スパークライン'],
      [1, 3, 2, 5, '=SPARKLINE(A2:D2)'],
      [4, 2, 6, 3, '=SPARKLINE(A3:D3,{"charttype","column"})']
    ]
  },
  // インポート関数
  {
    name: 'IMPORTDATA',
    category: '14. Web・インポート',
    description: 'URLからデータをインポート',
    data: [
      ['URL', 'インポート結果'],
      ['https://example.com/data.csv', '=IMPORTDATA(A2)']
    ]
  },
  {
    name: 'IMPORTFEED',
    category: '14. Web・インポート',
    description: 'RSSやAtomフィードをインポート',
    data: [
      ['フィードURL', 'クエリ', 'ヘッダー', 'アイテム数', 'インポート結果'],
      ['https://example.com/feed.rss', 'items title', true, 5, '=IMPORTFEED(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'IMPORTHTML',
    category: '14. Web・インポート',
    description: 'HTMLテーブルやリストをインポート',
    data: [
      ['URL', 'クエリ', 'インデックス', 'インポート結果'],
      ['https://example.com/page.html', 'table', 1, '=IMPORTHTML(A2,B2,C2)']
    ]
  },
  {
    name: 'IMPORTXML',
    category: '14. Web・インポート',
    description: 'XMLデータをインポート',
    data: [
      ['URL', 'XPathクエリ', 'インポート結果'],
      ['https://example.com/data.xml', '//item/title', '=IMPORTXML(A2,B2)']
    ]
  },
  {
    name: 'IMPORTRANGE',
    category: '14. Web・インポート',
    description: '他のスプレッドシートから範囲をインポート',
    data: [
      ['スプレッドシートURL', '範囲', 'インポート結果'],
      ['https://docs.google.com/spreadsheets/d/abc123', 'Sheet1!A1:C10', '=IMPORTRANGE(A2,B2)']
    ]
  },
  {
    name: 'IMAGE',
    category: '14. Web・インポート',
    description: '画像を挿入',
    data: [
      ['画像URL', 'モード', '高さ', '幅', '画像'],
      ['https://example.com/image.png', 1, 100, 100, '=IMAGE(A2,B2,C2,D2)']
    ]
  }
];

// Google Sheets専用関数の個別テスト
export const googleSheetsFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'GOOGLEFINANCE',
    category: '14. Web・インポート',
    description: '金融情報を取得',
    data: [
      ['銘柄', '属性', '開始日', '終了日', '間隔', '結果'],
      ['GOOG', 'price', '', '', '', '=GOOGLEFINANCE(A2,B2)'],
      ['AAPL', 'volume', '2024/1/1', '2024/1/31', 'DAILY', '=GOOGLEFINANCE(A3,B3,C3,D3,E3)']
    ]
  },
  {
    name: 'GOOGLETRANSLATE',
    category: '14. Web・インポート',
    description: 'テキストを翻訳',
    data: [
      ['テキスト', '元言語', '翻訳先言語', '翻訳結果'],
      ['Hello', 'en', 'ja', '=GOOGLETRANSLATE(A2,B2,C2)'],
      ['こんにちは', 'ja', 'en', '=GOOGLETRANSLATE(A3,B3,C3)']
    ]
  },
  {
    name: 'DETECTLANGUAGE',
    category: '14. Web・インポート',
    description: '言語を検出',
    data: [
      ['テキスト', '検出言語'],
      ['Hello World', '=DETECTLANGUAGE(A2)'],
      ['こんにちは世界', '=DETECTLANGUAGE(A3)'],
      ['Bonjour le monde', '=DETECTLANGUAGE(A4)']
    ],
    expectedValues: { 'B2': 'en', 'B3': 'ja', 'B4': 'fr' }
  },
  {
    name: 'TO_DATE',
    category: '14. Web・インポート',
    description: '値を日付に変換',
    data: [
      ['値', '日付'],
      [44926, '=TO_DATE(A2)'],
      ['2023/1/1', '=TO_DATE(A3)']
    ]
  },
  {
    name: 'TO_PERCENT',
    category: '14. Web・インポート',
    description: '値をパーセントに変換',
    data: [
      ['値', 'パーセント'],
      [0.25, '=TO_PERCENT(A2)'],
      [0.5, '=TO_PERCENT(A3)']
    ],
    expectedValues: { 'B2': '25%', 'B3': '50%' }
  },
  {
    name: 'TO_DOLLARS',
    category: '14. Web・インポート',
    description: '値をドル表記に変換',
    data: [
      ['値', 'ドル表記'],
      [1234.56, '=TO_DOLLARS(A2)'],
      [9876.54, '=TO_DOLLARS(A3)']
    ],
    expectedValues: { 'B2': '$1,234.56', 'B3': '$9,876.54' }
  },
  {
    name: 'TO_TEXT',
    category: '14. Web・インポート',
    description: '値をテキストに変換',
    data: [
      ['値', 'テキスト'],
      [123, '=TO_TEXT(A2)'],
      ['=TRUE()', '=TO_TEXT(A3)']
    ],
    expectedValues: { 'B2': '123', 'B3': 'TRUE' }
  }
];

// その他の関数の個別テスト
export const otherFunctionTests: IndividualFunctionTest[] = [
  {
    name: 'ISOMITTED',
    category: '15. その他',
    description: '引数が省略されているか判定',
    data: [
      ['値', '省略判定'],
      ['=LAMBDA(x,y,ISOMITTED(y))(A2)', '=A2'],
      ['=LAMBDA(x,y,ISOMITTED(y))(A3,B3)', '=A3']
    ]
  },
  {
    name: 'STOCKHISTORY',
    category: '15. その他',
    description: '株価履歴を取得',
    data: [
      ['銘柄', '開始日', '終了日', '間隔', 'ヘッダー', 'プロパティ', '履歴'],
      ['MSFT', '2024/1/1', '2024/1/31', 0, 1, 0, '=STOCKHISTORY(A2,B2,C2,D2,E2,F2)']
    ]
  },
  {
    name: 'GPT',
    category: '15. その他',
    description: 'GPTによるテキスト生成',
    data: [
      ['プロンプト', '生成結果'],
      ['Excelの便利な使い方を教えて', '=GPT(A2)']
    ]
  }
];

// すべての個別テストを結合
export const allIndividualFunctionTests: IndividualFunctionTest[] = [
  ...basicOperatorTests,
  ...mathFunctionTests,
  ...statisticalFunctionTests,
  ...textFunctionTests,
  ...dateFunctionTests,
  ...logicalFunctionTests,
  ...lookupFunctionTests,
  ...financialFunctionTests,
  ...matrixFunctionTests,
  ...informationFunctionTests,
  ...databaseFunctionTests,
  ...engineeringFunctionTests,
  ...dynamicArrayFunctionTests,
  ...cubeFunctionTests,
  ...webFunctionTests,
  ...googleSheetsFunctionTests,
  ...otherFunctionTests
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