// エンジニアリング関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// 単位変換の定義
const unitConversions: Record<string, Record<string, number>> = {
  // 長さ
  length: {
    m: 1,
    km: 0.001,
    cm: 100,
    mm: 1000,
    um: 1000000,
    nm: 1000000000,
    mi: 0.000621371,
    yd: 1.09361,
    ft: 3.28084,
    in: 39.3701,
    'Nautical mile': 0.000539957,
    'light-year': 1.057e-16,
    parsec: 3.24078e-17
  },
  // 重量
  weight: {
    g: 1,
    kg: 0.001,
    mg: 1000,
    ton: 0.000001,
    lbm: 0.00220462,
    ozm: 0.035274,
    grain: 15.4324,
    cwt: 0.0000220462,
    'stone': 0.000157473,
    uk_ton: 0.000000984207,
    'slug': 0.0000685218
  },
  // 時間
  time: {
    s: 1,
    sec: 1,
    min: 1/60,
    hr: 1/3600,
    day: 1/86400,
    yr: 1/31536000
  },
  // 温度（特殊処理が必要）
  temperature: {
    C: 1,
    F: 1,
    K: 1,
    Rank: 1,
    Reau: 1
  },
  // 体積
  volume: {
    l: 1,
    L: 1,
    ml: 1000,
    gal: 0.264172,
    qt: 1.05669,
    pt: 2.11338,
    cup: 4.22675,
    oz: 33.814,
    tbs: 67.628,
    tsp: 202.884,
    'uk_gal': 0.219969,
    'uk_qt': 0.879877,
    'uk_pt': 1.75975,
    m3: 0.001,
    cm3: 1000,
    mm3: 1000000,
    ft3: 0.0353147,
    in3: 61.0237,
    yd3: 0.00130795,
    'barrel': 0.00628981,
    'bushel': 0.0283776
  },
  // 面積
  area: {
    m2: 1,
    km2: 0.000001,
    cm2: 10000,
    mm2: 1000000,
    ha: 0.0001,
    acre: 0.000247105,
    ft2: 10.7639,
    in2: 1550,
    yd2: 1.19599,
    mi2: 3.861e-7
  },
  // 速度
  speed: {
    'm/s': 1,
    'km/h': 3.6,
    'mi/h': 2.23694,
    'ft/s': 3.28084,
    'knot': 1.94384
  },
  // 圧力
  pressure: {
    Pa: 1,
    atm: 9.86923e-6,
    bar: 0.00001,
    psi: 0.000145038,
    Torr: 0.00750062,
    mmHg: 0.00750062
  },
  // エネルギー
  energy: {
    J: 1,
    kJ: 0.001,
    cal: 0.239006,
    kcal: 0.000239006,
    Wh: 0.000277778,
    kWh: 2.77778e-7,
    BTU: 0.000947817,
    'ft-lb': 0.737562,
    eV: 6.242e18
  },
  // 力
  force: {
    N: 1,
    dyn: 100000,
    lbf: 0.224809,
    kgf: 0.101972,
    pond: 101.972
  },
  // 磁気
  magnetism: {
    T: 1,
    G: 10000,
    'Oe': 79.5775
  }
};

// 単位のカテゴリーを見つける関数
function findUnitCategory(unit: string): string | null {
  const lowerUnit = unit.toLowerCase();
  for (const [category, units] of Object.entries(unitConversions)) {
    if (lowerUnit in units || unit in units) {
      return category;
    }
  }
  return null;
}

// 温度変換の特殊処理
function convertTemperature(value: number, fromUnit: string, toUnit: string): number {
  // まずケルビンに変換
  let kelvin: number;
  
  switch (fromUnit) {
    case 'C':
      kelvin = value + 273.15;
      break;
    case 'F':
      kelvin = (value - 32) * 5/9 + 273.15;
      break;
    case 'K':
      kelvin = value;
      break;
    case 'Rank':
      kelvin = value * 5/9;
      break;
    case 'Reau':
      kelvin = value * 5/4 + 273.15;
      break;
    default:
      throw new Error('Invalid temperature unit');
  }
  
  // ケルビンから目的の単位に変換
  switch (toUnit) {
    case 'C':
      return kelvin - 273.15;
    case 'F':
      return (kelvin - 273.15) * 9/5 + 32;
    case 'K':
      return kelvin;
    case 'Rank':
      return kelvin * 9/5;
    case 'Reau':
      return (kelvin - 273.15) * 4/5;
    default:
      throw new Error('Invalid temperature unit');
  }
}

// CONVERT関数の実装
export const CONVERT: CustomFormula = {
  name: 'CONVERT',
  pattern: /CONVERT\(([^,]+),\s*"([^"]+)",\s*"([^"]+)"\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valueRef, fromUnit, toUnit] = matches;
    
    try {
      const value = parseFloat(getCellValue(valueRef.trim(), context)?.toString() ?? valueRef.trim());
      
      if (isNaN(value)) {
        return FormulaError.VALUE;
      }
      
      // 同じ単位の場合
      if (fromUnit === toUnit) {
        return value;
      }
      
      // 単位のカテゴリーを見つける
      const fromCategory = findUnitCategory(fromUnit);
      const toCategory = findUnitCategory(toUnit);
      
      if (!fromCategory || !toCategory) {
        return FormulaError.NA;
      }
      
      if (fromCategory !== toCategory) {
        return FormulaError.NA;
      }
      
      // 温度の特殊処理
      if (fromCategory === 'temperature') {
        return convertTemperature(value, fromUnit, toUnit);
      }
      
      // 通常の単位変換
      const fromFactor = unitConversions[fromCategory][fromUnit] || unitConversions[fromCategory][fromUnit.toLowerCase()];
      const toFactor = unitConversions[toCategory][toUnit] || unitConversions[toCategory][toUnit.toLowerCase()];
      
      if (!fromFactor || !toFactor) {
        return FormulaError.NA;
      }
      
      // 基準単位に変換してから目的の単位に変換
      return (value / fromFactor) * toFactor;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BIN2DEC関数の実装（2進数から10進数）
export const BIN2DEC: CustomFormula = {
  name: 'BIN2DEC',
  pattern: /BIN2DEC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, binaryRef] = matches;
    
    try {
      let binaryStr = getCellValue(binaryRef.trim(), context)?.toString() ?? binaryRef.trim();
      
      // 引用符を除去
      if (binaryStr.startsWith('"') && binaryStr.endsWith('"')) {
        binaryStr = binaryStr.slice(1, -1);
      }
      
      // 2進数の検証
      if (!/^[01]+$/.test(binaryStr)) {
        return FormulaError.NUM;
      }
      
      // 10文字以内（最大512ビット）
      if (binaryStr.length > 10) {
        return FormulaError.NUM;
      }
      
      // 負の数の処理（2の補数）
      if (binaryStr.length === 10 && binaryStr.startsWith('1')) {
        // 2の補数から10進数に変換
        let inverted = '';
        for (const bit of binaryStr) {
          inverted += bit === '0' ? '1' : '0';
        }
        return -(parseInt(inverted, 2) + 1);
      }
      
      return parseInt(binaryStr, 2);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DEC2BIN関数の実装（10進数から2進数）
export const DEC2BIN: CustomFormula = {
  name: 'DEC2BIN',
  pattern: /DEC2BIN\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, decimalRef, placesRef] = matches;
    
    try {
      const decimal = parseInt(getCellValue(decimalRef.trim(), context)?.toString() ?? decimalRef.trim());
      const places = placesRef ? parseInt(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim()) : undefined;
      
      if (isNaN(decimal)) {
        return FormulaError.VALUE;
      }
      
      // 範囲チェック（-512 から 511）
      if (decimal < -512 || decimal > 511) {
        return FormulaError.NUM;
      }
      
      if (places !== undefined && (isNaN(places) || places < 0 || places > 10)) {
        return FormulaError.VALUE;
      }
      
      let binary: string;
      
      if (decimal >= 0) {
        binary = decimal.toString(2);
      } else {
        // 負の数は2の補数で表現
        const positive = Math.abs(decimal);
        let binaryPositive = positive.toString(2);
        
        // 10ビットに拡張
        binaryPositive = binaryPositive.padStart(10, '0');
        
        // ビット反転
        let inverted = '';
        for (const bit of binaryPositive) {
          inverted += bit === '0' ? '1' : '0';
        }
        
        // 1を加算
        const invertedNum = parseInt(inverted, 2) + 1;
        binary = invertedNum.toString(2);
        
        // 10ビットに制限
        if (binary.length > 10) {
          binary = binary.slice(-10);
        }
      }
      
      // 指定された桁数にパディング
      if (places !== undefined && binary.length < places) {
        binary = binary.padStart(places, '0');
      }
      
      return binary;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// HEX2DEC関数の実装（16進数から10進数）
export const HEX2DEC: CustomFormula = {
  name: 'HEX2DEC',
  pattern: /HEX2DEC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, hexRef] = matches;
    
    try {
      let hexStr = getCellValue(hexRef.trim(), context)?.toString() ?? hexRef.trim();
      
      // 引用符を除去
      if (hexStr.startsWith('"') && hexStr.endsWith('"')) {
        hexStr = hexStr.slice(1, -1);
      }
      
      // 16進数の検証
      if (!/^[0-9A-Fa-f]+$/.test(hexStr)) {
        return FormulaError.NUM;
      }
      
      // 10文字以内
      if (hexStr.length > 10) {
        return FormulaError.NUM;
      }
      
      // 負の数の処理（最上位ビットが1の場合）
      if (hexStr.length === 10 && parseInt(hexStr[0], 16) >= 8) {
        // 2の補数から10進数に変換
        const maxValue = Math.pow(16, 10);
        const value = parseInt(hexStr, 16);
        return value - maxValue;
      }
      
      return parseInt(hexStr, 16);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DEC2HEX関数の実装（10進数から16進数）
export const DEC2HEX: CustomFormula = {
  name: 'DEC2HEX',
  pattern: /DEC2HEX\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, decimalRef, placesRef] = matches;
    
    try {
      const decimal = parseInt(getCellValue(decimalRef.trim(), context)?.toString() ?? decimalRef.trim());
      const places = placesRef ? parseInt(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim()) : undefined;
      
      if (isNaN(decimal)) {
        return FormulaError.VALUE;
      }
      
      // 範囲チェック
      const minValue = -Math.pow(16, 9); // -68719476736
      const maxValue = Math.pow(16, 9) - 1; // 68719476735
      
      if (decimal < minValue || decimal > maxValue) {
        return FormulaError.NUM;
      }
      
      if (places !== undefined && (isNaN(places) || places < 0 || places > 10)) {
        return FormulaError.VALUE;
      }
      
      let hex: string;
      
      if (decimal >= 0) {
        hex = decimal.toString(16).toUpperCase();
      } else {
        // 負の数は2の補数で表現
        const maxHexValue = Math.pow(16, 10);
        const positiveEquivalent = maxHexValue + decimal;
        hex = positiveEquivalent.toString(16).toUpperCase();
      }
      
      // 指定された桁数にパディング
      if (places !== undefined && hex.length < places) {
        hex = hex.padStart(places, '0');
      }
      
      return hex;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BIN2HEX関数の実装（2進数から16進数）
export const BIN2HEX: CustomFormula = {
  name: 'BIN2HEX',
  pattern: /BIN2HEX\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, binaryRef, placesRef] = matches;
    
    try {
      // まず10進数に変換
      const decimalResult = BIN2DEC.calculate([matches[0], binaryRef], context);
      
      if (typeof decimalResult === 'string' && decimalResult.startsWith('#')) {
        return decimalResult;
      }
      
      // 10進数から16進数に変換
      const decimal = decimalResult as number;
      const places = placesRef ? parseInt(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim()) : undefined;
      
      return DEC2HEX.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// HEX2BIN関数の実装（16進数から2進数）
export const HEX2BIN: CustomFormula = {
  name: 'HEX2BIN',
  pattern: /HEX2BIN\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, hexRef, placesRef] = matches;
    
    try {
      // まず10進数に変換
      const decimalResult = HEX2DEC.calculate([matches[0], hexRef], context);
      
      if (typeof decimalResult === 'string' && decimalResult.startsWith('#')) {
        return decimalResult;
      }
      
      // 10進数から2進数に変換
      const decimal = decimalResult as number;
      
      // 2進数の範囲チェック（-512 から 511）
      if (decimal < -512 || decimal > 511) {
        return FormulaError.NUM;
      }
      
      const places = placesRef ? parseInt(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim()) : undefined;
      
      return DEC2BIN.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// OCT2DEC関数の実装（8進数から10進数）
export const OCT2DEC: CustomFormula = {
  name: 'OCT2DEC',
  pattern: /OCT2DEC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, octRef] = matches;
    
    try {
      let octStr = getCellValue(octRef.trim(), context)?.toString() ?? octRef.trim();
      
      // 引用符を除去
      if (octStr.startsWith('"') && octStr.endsWith('"')) {
        octStr = octStr.slice(1, -1);
      }
      
      // 8進数の検証
      if (!/^[0-7]+$/.test(octStr)) {
        return FormulaError.NUM;
      }
      
      // 10文字以内
      if (octStr.length > 10) {
        return FormulaError.NUM;
      }
      
      // 負の数の処理（最上位ビットが1の場合）
      if (octStr.length === 10 && octStr[0] >= '4') {
        // 2の補数から10進数に変換
        const maxValue = Math.pow(8, 10);
        const value = parseInt(octStr, 8);
        return value - maxValue;
      }
      
      return parseInt(octStr, 8);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DEC2OCT関数の実装（10進数から8進数）
export const DEC2OCT: CustomFormula = {
  name: 'DEC2OCT',
  pattern: /DEC2OCT\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, decimalRef, placesRef] = matches;
    
    try {
      const decimal = parseInt(getCellValue(decimalRef.trim(), context)?.toString() ?? decimalRef.trim());
      const places = placesRef ? parseInt(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim()) : undefined;
      
      if (isNaN(decimal)) {
        return FormulaError.VALUE;
      }
      
      // 範囲チェック
      const minValue = -Math.pow(8, 9); // -134217728
      const maxValue = Math.pow(8, 9) - 1; // 134217727
      
      if (decimal < minValue || decimal > maxValue) {
        return FormulaError.NUM;
      }
      
      if (places !== undefined && (isNaN(places) || places < 0 || places > 10)) {
        return FormulaError.VALUE;
      }
      
      let oct: string;
      
      if (decimal >= 0) {
        oct = decimal.toString(8);
      } else {
        // 負の数は2の補数で表現
        const maxOctValue = Math.pow(8, 10);
        const positiveEquivalent = maxOctValue + decimal;
        oct = positiveEquivalent.toString(8);
      }
      
      // 指定された桁数にパディング
      if (places !== undefined && oct.length < places) {
        oct = oct.padStart(places, '0');
      }
      
      return oct;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BIN2OCT関数の実装（2進数から8進数）
export const BIN2OCT: CustomFormula = {
  name: 'BIN2OCT',
  pattern: /BIN2OCT\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, binaryRef, placesRef] = matches;
    
    try {
      // まず10進数に変換
      const decimalResult = BIN2DEC.calculate([matches[0], binaryRef], context);
      
      if (typeof decimalResult === 'string' && decimalResult.startsWith('#')) {
        return decimalResult;
      }
      
      // 10進数から8進数に変換
      const decimal = decimalResult as number;
      const places = placesRef ? parseInt(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim()) : undefined;
      
      return DEC2OCT.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// HEX2OCT関数の実装（16進数から8進数）
export const HEX2OCT: CustomFormula = {
  name: 'HEX2OCT',
  pattern: /HEX2OCT\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, hexRef, placesRef] = matches;
    
    try {
      // まず10進数に変換
      const decimalResult = HEX2DEC.calculate([matches[0], hexRef], context);
      
      if (typeof decimalResult === 'string' && decimalResult.startsWith('#')) {
        return decimalResult;
      }
      
      // 10進数から8進数に変換
      const decimal = decimalResult as number;
      const places = placesRef ? parseInt(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim()) : undefined;
      
      return DEC2OCT.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// OCT2BIN関数の実装（8進数から2進数）
export const OCT2BIN: CustomFormula = {
  name: 'OCT2BIN',
  pattern: /OCT2BIN\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, octRef, placesRef] = matches;
    
    try {
      // まず10進数に変換
      const decimalResult = OCT2DEC.calculate([matches[0], octRef], context);
      
      if (typeof decimalResult === 'string' && decimalResult.startsWith('#')) {
        return decimalResult;
      }
      
      // 10進数から2進数に変換
      const decimal = decimalResult as number;
      
      // 2進数の範囲チェック（-512 から 511）
      if (decimal < -512 || decimal > 511) {
        return FormulaError.NUM;
      }
      
      const places = placesRef ? parseInt(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim()) : undefined;
      
      return DEC2BIN.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// OCT2HEX関数の実装（8進数から16進数）
export const OCT2HEX: CustomFormula = {
  name: 'OCT2HEX',
  pattern: /OCT2HEX\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, octRef, placesRef] = matches;
    
    try {
      // まず10進数に変換
      const decimalResult = OCT2DEC.calculate([matches[0], octRef], context);
      
      if (typeof decimalResult === 'string' && decimalResult.startsWith('#')) {
        return decimalResult;
      }
      
      // 10進数から16進数に変換
      const decimal = decimalResult as number;
      const places = placesRef ? parseInt(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim()) : undefined;
      
      return DEC2HEX.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};