// エンジニアリング関数の定義
import { type FunctionDefinition } from '../types/types';
import { COLOR_SCHEMES } from '../types/colorSchemes';

export const ENGINEERING_FUNCTIONS: Record<string, FunctionDefinition> = {
  // 単位変換関数
  CONVERT: {
    name: 'CONVERT',
    syntax: 'CONVERT(number, from_unit, to_unit)',
    params: [
      { name: 'number', desc: '変換する数値' },
      { name: 'from_unit', desc: '変換元の単位' },
      { name: 'to_unit', desc: '変換先の単位' }
    ],
    description: '数値の単位を変換します',
    category: 'engineering',
    examples: ['=CONVERT(1,"in","cm")', '=CONVERT(68,"F","C")', '=CONVERT(2.5,"ft","m")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  // 数値システム変換関数
  BIN2DEC: {
    name: 'BIN2DEC',
    syntax: 'BIN2DEC(number)',
    params: [
      { name: 'number', desc: '2進数' }
    ],
    description: '2進数を10進数に変換します',
    category: 'engineering',
    examples: ['=BIN2DEC(1010)', '=BIN2DEC(11111011)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  BIN2HEX: {
    name: 'BIN2HEX',
    syntax: 'BIN2HEX(number, [places])',
    params: [
      { name: 'number', desc: '2進数' },
      { name: 'places', desc: '桁数', optional: true }
    ],
    description: '2進数を16進数に変換します',
    category: 'engineering',
    examples: ['=BIN2HEX(11111011)', '=BIN2HEX(1110,8)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  BIN2OCT: {
    name: 'BIN2OCT',
    syntax: 'BIN2OCT(number, [places])',
    params: [
      { name: 'number', desc: '2進数' },
      { name: 'places', desc: '桁数', optional: true }
    ],
    description: '2進数を8進数に変換します',
    category: 'engineering',
    examples: ['=BIN2OCT(1001)', '=BIN2OCT(1100100,4)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  DEC2BIN: {
    name: 'DEC2BIN',
    syntax: 'DEC2BIN(number, [places])',
    params: [
      { name: 'number', desc: '10進数' },
      { name: 'places', desc: '桁数', optional: true }
    ],
    description: '10進数を2進数に変換します',
    category: 'engineering',
    examples: ['=DEC2BIN(9)', '=DEC2BIN(100,8)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  DEC2HEX: {
    name: 'DEC2HEX',
    syntax: 'DEC2HEX(number, [places])',
    params: [
      { name: 'number', desc: '10進数' },
      { name: 'places', desc: '桁数', optional: true }
    ],
    description: '10進数を16進数に変換します',
    category: 'engineering',
    examples: ['=DEC2HEX(100)', '=DEC2HEX(100,4)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  DEC2OCT: {
    name: 'DEC2OCT',
    syntax: 'DEC2OCT(number, [places])',
    params: [
      { name: 'number', desc: '10進数' },
      { name: 'places', desc: '桁数', optional: true }
    ],
    description: '10進数を8進数に変換します',
    category: 'engineering',
    examples: ['=DEC2OCT(58)', '=DEC2OCT(58,3)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  HEX2BIN: {
    name: 'HEX2BIN',
    syntax: 'HEX2BIN(number, [places])',
    params: [
      { name: 'number', desc: '16進数' },
      { name: 'places', desc: '桁数', optional: true }
    ],
    description: '16進数を2進数に変換します',
    category: 'engineering',
    examples: ['=HEX2BIN("F")', '=HEX2BIN("B7",8)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  HEX2DEC: {
    name: 'HEX2DEC',
    syntax: 'HEX2DEC(number)',
    params: [
      { name: 'number', desc: '16進数' }
    ],
    description: '16進数を10進数に変換します',
    category: 'engineering',
    examples: ['=HEX2DEC("A5")', '=HEX2DEC("FFFFFFFF5B")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  HEX2OCT: {
    name: 'HEX2OCT',
    syntax: 'HEX2OCT(number, [places])',
    params: [
      { name: 'number', desc: '16進数' },
      { name: 'places', desc: '桁数', optional: true }
    ],
    description: '16進数を8進数に変換します',
    category: 'engineering',
    examples: ['=HEX2OCT("F")', '=HEX2OCT("3B4E")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  OCT2BIN: {
    name: 'OCT2BIN',
    syntax: 'OCT2BIN(number, [places])',
    params: [
      { name: 'number', desc: '8進数' },
      { name: 'places', desc: '桁数', optional: true }
    ],
    description: '8進数を2進数に変換します',
    category: 'engineering',
    examples: ['=OCT2BIN(3)', '=OCT2BIN(7777777000)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  OCT2DEC: {
    name: 'OCT2DEC',
    syntax: 'OCT2DEC(number)',
    params: [
      { name: 'number', desc: '8進数' }
    ],
    description: '8進数を10進数に変換します',
    category: 'engineering',
    examples: ['=OCT2DEC(54)', '=OCT2DEC(7777777533)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  OCT2HEX: {
    name: 'OCT2HEX',
    syntax: 'OCT2HEX(number, [places])',
    params: [
      { name: 'number', desc: '8進数' },
      { name: 'places', desc: '桁数', optional: true }
    ],
    description: '8進数を16進数に変換します',
    category: 'engineering',
    examples: ['=OCT2HEX(100)', '=OCT2HEX(7777777533)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  // 複素数関数
  COMPLEX: {
    name: 'COMPLEX',
    syntax: 'COMPLEX(real_num, i_num, [suffix])',
    params: [
      { name: 'real_num', desc: '実部' },
      { name: 'i_num', desc: '虚部' },
      { name: 'suffix', desc: '虚数単位（"i"または"j"）', optional: true }
    ],
    description: '実部と虚部から複素数を作成します',
    category: 'engineering',
    examples: ['=COMPLEX(3,4)', '=COMPLEX(3,4,"j")', '=COMPLEX(0,1)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMABS: {
    name: 'IMABS',
    syntax: 'IMABS(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の絶対値を返します',
    category: 'engineering',
    examples: ['=IMABS("5+12i")', '=IMABS("3+4i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMAGINARY: {
    name: 'IMAGINARY',
    syntax: 'IMAGINARY(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の虚部を返します',
    category: 'engineering',
    examples: ['=IMAGINARY("3+4i")', '=IMAGINARY("0-j")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMREAL: {
    name: 'IMREAL',
    syntax: 'IMREAL(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の実部を返します',
    category: 'engineering',
    examples: ['=IMREAL("6-9i")', '=IMREAL("2+3j")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMARGUMENT: {
    name: 'IMARGUMENT',
    syntax: 'IMARGUMENT(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の偏角（ラジアン）を返します',
    category: 'engineering',
    examples: ['=IMARGUMENT("3+4i")', '=IMARGUMENT("1+i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMCONJUGATE: {
    name: 'IMCONJUGATE',
    syntax: 'IMCONJUGATE(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の複素共役を返します',
    category: 'engineering',
    examples: ['=IMCONJUGATE("3+4i")', '=IMCONJUGATE("5-3i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMSUM: {
    name: 'IMSUM',
    syntax: 'IMSUM(inumber1, [inumber2], ...)',
    params: [
      { name: 'inumber1', desc: '1つ目の複素数' },
      { name: 'inumber2', desc: '2つ目の複素数', optional: true }
    ],
    description: '複素数の和を返します',
    category: 'engineering',
    examples: ['=IMSUM("3+4i","5-3i")', '=IMSUM("1+i","2+2i","3+3i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMSUB: {
    name: 'IMSUB',
    syntax: 'IMSUB(inumber1, inumber2)',
    params: [
      { name: 'inumber1', desc: '被減数（複素数）' },
      { name: 'inumber2', desc: '減数（複素数）' }
    ],
    description: '複素数の差を返します',
    category: 'engineering',
    examples: ['=IMSUB("13+4i","5+3i")', '=IMSUB("3+2i","1+i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMPRODUCT: {
    name: 'IMPRODUCT',
    syntax: 'IMPRODUCT(inumber1, [inumber2], ...)',
    params: [
      { name: 'inumber1', desc: '1つ目の複素数' },
      { name: 'inumber2', desc: '2つ目の複素数', optional: true }
    ],
    description: '複素数の積を返します',
    category: 'engineering',
    examples: ['=IMPRODUCT("5+2i","6+7i")', '=IMPRODUCT("1+i","2+i","3+i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMDIV: {
    name: 'IMDIV',
    syntax: 'IMDIV(inumber1, inumber2)',
    params: [
      { name: 'inumber1', desc: '被除数（複素数）' },
      { name: 'inumber2', desc: '除数（複素数）' }
    ],
    description: '複素数の商を返します',
    category: 'engineering',
    examples: ['=IMDIV("-238+240i","10+24i")', '=IMDIV("4+2i","1+i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMPOWER: {
    name: 'IMPOWER',
    syntax: 'IMPOWER(inumber, number)',
    params: [
      { name: 'inumber', desc: '底となる複素数' },
      { name: 'number', desc: '指数' }
    ],
    description: '複素数のべき乗を返します',
    category: 'engineering',
    examples: ['=IMPOWER("2+3i",3)', '=IMPOWER("1+i",2)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMSQRT: {
    name: 'IMSQRT',
    syntax: 'IMSQRT(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の平方根を返します',
    category: 'engineering',
    examples: ['=IMSQRT("1+i")', '=IMSQRT("4+3i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMEXP: {
    name: 'IMEXP',
    syntax: 'IMEXP(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の指数関数を返します',
    category: 'engineering',
    examples: ['=IMEXP("1+i")', '=IMEXP("0+i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMLN: {
    name: 'IMLN',
    syntax: 'IMLN(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の自然対数を返します',
    category: 'engineering',
    examples: ['=IMLN("3+4i")', '=IMLN("1+i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMLOG10: {
    name: 'IMLOG10',
    syntax: 'IMLOG10(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の常用対数を返します',
    category: 'engineering',
    examples: ['=IMLOG10("3+4i")', '=IMLOG10("10+0i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMLOG2: {
    name: 'IMLOG2',
    syntax: 'IMLOG2(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の2を底とする対数を返します',
    category: 'engineering',
    examples: ['=IMLOG2("4+0i")', '=IMLOG2("2+0i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMSIN: {
    name: 'IMSIN',
    syntax: 'IMSIN(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の正弦を返します',
    category: 'engineering',
    examples: ['=IMSIN("4+3i")', '=IMSIN("1+i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMCOS: {
    name: 'IMCOS',
    syntax: 'IMCOS(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の余弦を返します',
    category: 'engineering',
    examples: ['=IMCOS("1+i")', '=IMCOS("4+3i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  IMTAN: {
    name: 'IMTAN',
    syntax: 'IMTAN(inumber)',
    params: [
      { name: 'inumber', desc: '複素数' }
    ],
    description: '複素数の正接を返します',
    category: 'engineering',
    examples: ['=IMTAN("4+3i")', '=IMTAN("1+i")'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  // ベッセル関数
  BESSELI: {
    name: 'BESSELI',
    syntax: 'BESSELI(x, n)',
    params: [
      { name: 'x', desc: '関数を評価する値' },
      { name: 'n', desc: 'ベッセル関数の次数' }
    ],
    description: '修正ベッセル関数In(x)を返します',
    category: 'engineering',
    examples: ['=BESSELI(1.5,1)', '=BESSELI(1,0)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  BESSELJ: {
    name: 'BESSELJ',
    syntax: 'BESSELJ(x, n)',
    params: [
      { name: 'x', desc: '関数を評価する値' },
      { name: 'n', desc: 'ベッセル関数の次数' }
    ],
    description: 'ベッセル関数Jn(x)を返します',
    category: 'engineering',
    examples: ['=BESSELJ(1.9,2)', '=BESSELJ(1,0)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  BESSELK: {
    name: 'BESSELK',
    syntax: 'BESSELK(x, n)',
    params: [
      { name: 'x', desc: '関数を評価する値' },
      { name: 'n', desc: 'ベッセル関数の次数' }
    ],
    description: '修正ベッセル関数Kn(x)を返します',
    category: 'engineering',
    examples: ['=BESSELK(1.5,1)', '=BESSELK(1,0)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  BESSELY: {
    name: 'BESSELY',
    syntax: 'BESSELY(x, n)',
    params: [
      { name: 'x', desc: '関数を評価する値' },
      { name: 'n', desc: 'ベッセル関数の次数' }
    ],
    description: 'ベッセル関数Yn(x)を返します',
    category: 'engineering',
    examples: ['=BESSELY(2.5,1)', '=BESSELY(1,0)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  // その他のエンジニアリング関数
  DELTA: {
    name: 'DELTA',
    syntax: 'DELTA(number1, [number2])',
    params: [
      { name: 'number1', desc: '1つ目の数値' },
      { name: 'number2', desc: '2つ目の数値', optional: true }
    ],
    description: 'クロネッカーのデルタ関数を返します',
    category: 'engineering',
    examples: ['=DELTA(5,4)', '=DELTA(5,5)', '=DELTA(0.5,0)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  GESTEP: {
    name: 'GESTEP',
    syntax: 'GESTEP(number, [step])',
    params: [
      { name: 'number', desc: '評価する数値' },
      { name: 'step', desc: 'しきい値', optional: true }
    ],
    description: 'ステップ関数を返します',
    category: 'engineering',
    examples: ['=GESTEP(5,4)', '=GESTEP(-4,-5)', '=GESTEP(-1)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  ERF: {
    name: 'ERF',
    syntax: 'ERF(lower_limit, [upper_limit])',
    params: [
      { name: 'lower_limit', desc: '下限' },
      { name: 'upper_limit', desc: '上限', optional: true }
    ],
    description: '誤差関数を返します',
    category: 'engineering',
    examples: ['=ERF(0.745)', '=ERF(1,2)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  ERFC: {
    name: 'ERFC',
    syntax: 'ERFC(x)',
    params: [
      { name: 'x', desc: '関数を評価する値' }
    ],
    description: '相補誤差関数を返します',
    category: 'engineering',
    examples: ['=ERFC(1)', '=ERFC(0.5)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  // ビット演算関数
  BITAND: {
    name: 'BITAND',
    syntax: 'BITAND(number1, number2)',
    params: [
      { name: 'number1', desc: '1つ目の数値' },
      { name: 'number2', desc: '2つ目の数値' }
    ],
    description: 'ビット単位AND演算を返します',
    category: 'engineering',
    examples: ['=BITAND(1,5)', '=BITAND(13,25)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  BITOR: {
    name: 'BITOR',
    syntax: 'BITOR(number1, number2)',
    params: [
      { name: 'number1', desc: '1つ目の数値' },
      { name: 'number2', desc: '2つ目の数値' }
    ],
    description: 'ビット単位OR演算を返します',
    category: 'engineering',
    examples: ['=BITOR(23,10)', '=BITOR(5,3)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  BITXOR: {
    name: 'BITXOR',
    syntax: 'BITXOR(number1, number2)',
    params: [
      { name: 'number1', desc: '1つ目の数値' },
      { name: 'number2', desc: '2つ目の数値' }
    ],
    description: 'ビット単位XOR演算を返します',
    category: 'engineering',
    examples: ['=BITXOR(5,3)', '=BITXOR(1,5)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  BITLSHIFT: {
    name: 'BITLSHIFT',
    syntax: 'BITLSHIFT(number, shift_amount)',
    params: [
      { name: 'number', desc: 'シフトする数値' },
      { name: 'shift_amount', desc: 'シフトするビット数' }
    ],
    description: 'ビット左シフトを返します',
    category: 'engineering',
    examples: ['=BITLSHIFT(4,2)', '=BITLSHIFT(3,1)'],
    colorScheme: COLOR_SCHEMES.engineering
  },

  BITRSHIFT: {
    name: 'BITRSHIFT',
    syntax: 'BITRSHIFT(number, shift_amount)',
    params: [
      { name: 'number', desc: 'シフトする数値' },
      { name: 'shift_amount', desc: 'シフトするビット数' }
    ],
    description: 'ビット右シフトを返します',
    category: 'engineering',
    examples: ['=BITRSHIFT(13,2)', '=BITRSHIFT(8,3)'],
    colorScheme: COLOR_SCHEMES.engineering
  }
};