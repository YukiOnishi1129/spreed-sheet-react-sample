// エンジニアリング関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// 単位変換の定義（Excel完全互換）
const unitConversions: Record<string, Record<string, number>> = {
  // 長さ - Excel exact conversion factors
  length: {
    'm': 1,
    'km': 0.001,
    'cm': 100,
    'mm': 1000,
    'in': 39.3700787,
    'ft': 3.2808399,
    'yd': 1.0936133,
    'mi': 0.0006213712,
    'Nmi': 0.0005399568,
    'um': 1000000,
    'Pica': 2834.6457,
    'Picapt': 2834.6457,
    'pica': 2834.6457,
    'ang': 10000000000,
    'Angstrom': 10000000000,
    'ly': 1.057000766e-16,
    'parsec': 3.240779289e-17,
    'pc': 3.240779289e-17
  },
  // 重量/質量 - Excel exact conversion factors
  weight: {
    'g': 1,
    'kg': 0.001,
    'mg': 1000,
    'u': 6.02214076e26,
    'ozm': 0.035273962,
    'lbm': 0.002204623,
    'stone': 0.000157473,
    'ton': 0.000001,
    'grain': 15.432358,
    'pweight': 0.00006479891,
    'hweight': 0.000001968413,
    'shweight': 0.000002204623,
    'brton': 0.0000009842065,
    'cwt': 0.00002204623,
    'shcwt': 0.00002204623,
    'lcwt': 0.00001968413,
    'uk_ton': 0.0000009842065,
    'LTON': 0.0000009842065
  },
  // 時間 - Excel exact conversion factors
  time: {
    'sec': 1,
    's': 1,
    'min': 1/60,
    'hr': 1/3600,
    'day': 1/86400,
    'yr': 1/31557600
  },
  // 温度（特殊処理が必要）- Excel exact conversion
  temperature: {
    'C': 1,
    'cel': 1,
    'F': 1,
    'fah': 1,
    'K': 1,
    'kel': 1,
    'Rank': 1,
    'Reau': 1
  },
  // 体積 - Excel exact conversion factors
  volume: {
    'l': 1,
    'L': 1,
    'lt': 1,
    'm3': 0.001,
    'cm3': 1000,
    'cc': 1000,
    'in3': 61.02374409,
    'ft3': 0.035314667,
    'yd3': 0.001307951,
    'gal': 0.264172052,
    'qt': 1.056688209,
    'pt': 2.113376419,
    'cup': 4.226752838,
    'fl_oz': 33.8140227,
    'tbs': 67.6280454,
    'tsp': 202.8841362,
    'uk_gal': 0.219969248,
    'uk_qt': 0.879876993,
    'uk_pt': 1.759753986,
    'BARL': 0.006289811,
    'bushel': 0.028377593,
    'regton': 0.000353147
  },
  // 面積 - Excel exact conversion factors
  area: {
    'm2': 1,
    'km2': 0.000001,
    'cm2': 10000,
    'mm2': 1000000,
    'in2': 1550.0031,
    'ft2': 10.76391042,
    'yd2': 1.19599005,
    'acre': 0.000247105,
    'ha': 0.0001,
    'mi2': 3.861021585e-7
  },
  // 速度 - Excel exact conversion factors
  speed: {
    'm/s': 1,
    'm/sec': 1,
    'km/h': 3.6,
    'kn': 1.943844492,
    'mph': 2.236936292,
    'ft/s': 3.280839895,
    'ft/sec': 3.280839895
  },
  // 圧力 - Excel exact conversion factors
  pressure: {
    'Pa': 1,
    'p': 1,
    'atm': 9.869232667e-6,
    'at': 1.019716213e-5,
    'bar': 0.00001,
    'Torr': 0.007500617,
    'psi': 0.000145038,
    'mmHg': 0.007500617
  },
  // エネルギー - Excel exact conversion factors
  energy: {
    'J': 1,
    'j': 1,
    'kJ': 0.001,
    'e': 1,
    'c': 0.238902958,
    'cal': 0.238902958,
    'eV': 6.241457e18,
    'ev': 6.241457e18,
    'HPh': 3.725061e-7,
    'hh': 3.725061e-7,
    'Wh': 0.000277778,
    'wh': 0.000277778,
    'flb': 0.737562149,
    'BTU': 0.000947817,
    'btu': 0.000947817
  },
  // 力 - Excel exact conversion factors
  force: {
    'N': 1,
    'n': 1,
    'dyn': 100000,
    'dy': 100000,
    'lbf': 0.224808924,
    'pond': 101.9716213
  },
  // 磁気 - Excel exact conversion factors
  magnetism: {
    'T': 1,
    'ga': 10000,
    'Wb': 1,
    'maxwell': 100000000,
    'Oe': 79.57747155,
    'H': 1,
    'iH': 1000000000,
    'mH': 1000,
    'kH': 0.001,
    'A': 1
  },
  // データ単位 - バイナリプレフィックス付き
  data: {
    'byte': 1,
    'bit': 8,
    'kB': 1/1000,
    'MB': 1/1000000,
    'GB': 1/1000000000,
    'TB': 1/1000000000000,
    'PB': 1/1000000000000000,
    'Kibibyte': 1/1024,
    'Mebibyte': 1/(1024*1024),
    'Gibibyte': 1/(1024*1024*1024),
    'Gibyte': 1/(1024*1024*1024), // 同じ単位の別名
    'Tebibyte': 1/(1024*1024*1024*1024),
    'Pebibyte': 1/(1024*1024*1024*1024*1024)
  }
};

// 単位のカテゴリーを見つける関数（Excel完全互換：大文字小文字区別）
function findUnitCategory(unit: string): string | null {
  for (const [category, units] of Object.entries(unitConversions)) {
    if (unit in units) {
      return category;
    }
  }
  return null;
}

// 温度変換の特殊処理（Excel完全互換）
function convertTemperature(value: number, fromUnit: string, toUnit: string): number {
  // まずケルビンに変換
  let kelvin: number;
  
  switch (fromUnit) {
    case 'C':
    case 'cel':
      kelvin = value + 273.15;
      break;
    case 'F':
    case 'fah':
      kelvin = (value - 32) * 5/9 + 273.15;
      break;
    case 'K':
    case 'kel':
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
    case 'cel':
      return kelvin - 273.15;
    case 'F':
    case 'fah':
      return (kelvin - 273.15) * 9/5 + 32;
    case 'K':
    case 'kel':
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
  pattern: /\bCONVERT\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valueRef, fromUnitRef, toUnitRef] = matches;
    
    try {
      const value = parseFloat(getCellValue(valueRef.trim(), context)?.toString() ?? valueRef.trim());
      
      if (isNaN(value)) {
        return FormulaError.VALUE;
      }
      
      // Get unit strings from cells or literal values
      let fromUnitStr = getCellValue(fromUnitRef.trim(), context)?.toString();
      if (!fromUnitStr || fromUnitStr === fromUnitRef.trim()) {
        // If cell value not found, treat as literal string
        fromUnitStr = fromUnitRef.trim().replace(/^["']|["']$/g, '');
      }
      
      let toUnitStr = getCellValue(toUnitRef.trim(), context)?.toString();
      if (!toUnitStr || toUnitStr === toUnitRef.trim()) {
        // If cell value not found, treat as literal string  
        toUnitStr = toUnitRef.trim().replace(/^["']|["']$/g, '');
      }
      
      // Clean unit strings
      const fromUnitClean = fromUnitStr.trim();
      const toUnitClean = toUnitStr.trim();
      
      // Same unit case
      if (fromUnitClean === toUnitClean) {
        return value;
      }
      
      // Find unit categories
      const fromCategory = findUnitCategory(fromUnitClean);
      const toCategory = findUnitCategory(toUnitClean);
      
      if (!fromCategory || !toCategory) {
        return FormulaError.NA;
      }
      
      if (fromCategory !== toCategory) {
        return FormulaError.NA;
      }
      
      // Special temperature handling
      if (fromCategory === 'temperature') {
        return convertTemperature(value, fromUnitClean, toUnitClean);
      }
      
      // Normal unit conversion
      const fromFactor = unitConversions[fromCategory][fromUnitClean];
      const toFactor = unitConversions[toCategory][toUnitClean];
      
      if (fromFactor === undefined || toFactor === undefined) {
        return FormulaError.NA;
      }
      
      // Convert to base unit then to target unit
      const result = (value / fromFactor) * toFactor;
      
      // Round to avoid floating point precision issues
      return Math.round(result * 1000000) / 1000000;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// BIN2DEC関数の実装（2進数から10進数）
export const BIN2DEC: CustomFormula = {
  name: 'BIN2DEC',
  pattern: /\bBIN2DEC\(([^)]+)\)/i,
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
  pattern: /\bDEC2BIN\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, decimalRef, placesRef] = matches;
    
    try {
      const decimal = Math.trunc(parseFloat(getCellValue(decimalRef.trim(), context)?.toString() ?? decimalRef.trim()));
      const places = placesRef ? Math.trunc(parseFloat(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim())) : undefined;
      
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
  pattern: /\bHEX2DEC\(([^)]+)\)/i,
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
  pattern: /\bDEC2HEX\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, decimalRef, placesRef] = matches;
    
    try {
      const decimal = Math.trunc(parseFloat(getCellValue(decimalRef.trim(), context)?.toString() ?? decimalRef.trim()));
      const places = placesRef ? Math.trunc(parseFloat(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim())) : undefined;
      
      if (isNaN(decimal)) {
        return FormulaError.VALUE;
      }
      
      // 範囲チェック（40-bit signed range）
      const minValue = -549755813888; // -2^39
      const maxValue = 549755813887; // 2^39 - 1
      
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
  pattern: /\bBIN2HEX\(([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      const places = placesRef ? Math.trunc(parseFloat(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim())) : undefined;
      
      // 10進数から16進数への変換を直接実行して文字列長をチェック
      let hex: string;
      if (decimal >= 0) {
        hex = decimal.toString(16).toUpperCase();
      } else {
        // 負の数は2の補数で表現
        const maxHexValue = Math.pow(16, 10);
        const positiveEquivalent = maxHexValue + decimal;
        hex = positiveEquivalent.toString(16).toUpperCase();
      }
      
      // places が指定されており、16進数の結果より小さい場合はNUM error
      if (places !== undefined && hex.length > places) {
        return FormulaError.NUM;
      }
      
      return DEC2HEX.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// HEX2BIN関数の実装（16進数から2進数）
export const HEX2BIN: CustomFormula = {
  name: 'HEX2BIN',
  pattern: /\bHEX2BIN\(([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      
      const places = placesRef ? Math.trunc(parseFloat(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim())) : undefined;
      
      return DEC2BIN.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// OCT2DEC関数の実装（8進数から10進数）
export const OCT2DEC: CustomFormula = {
  name: 'OCT2DEC',
  pattern: /\bOCT2DEC\(([^)]+)\)/i,
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
  pattern: /\bDEC2OCT\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, decimalRef, placesRef] = matches;
    
    try {
      const decimal = Math.trunc(parseFloat(getCellValue(decimalRef.trim(), context)?.toString() ?? decimalRef.trim()));
      const places = placesRef ? Math.trunc(parseFloat(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim())) : undefined;
      
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
  pattern: /\bBIN2OCT\(([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      const places = placesRef ? Math.trunc(parseFloat(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim())) : undefined;
      
      return DEC2OCT.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// HEX2OCT関数の実装（16進数から8進数）
export const HEX2OCT: CustomFormula = {
  name: 'HEX2OCT',
  pattern: /\bHEX2OCT\(([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      const places = placesRef ? Math.trunc(parseFloat(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim())) : undefined;
      
      return DEC2OCT.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// OCT2BIN関数の実装（8進数から2進数）
export const OCT2BIN: CustomFormula = {
  name: 'OCT2BIN',
  pattern: /\bOCT2BIN\(([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      
      const places = placesRef ? Math.trunc(parseFloat(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim())) : undefined;
      
      return DEC2BIN.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// OCT2HEX関数の実装（8進数から16進数）
export const OCT2HEX: CustomFormula = {
  name: 'OCT2HEX',
  pattern: /\bOCT2HEX\(([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      const places = placesRef ? Math.trunc(parseFloat(getCellValue(placesRef.trim(), context)?.toString() ?? placesRef.trim())) : undefined;
      
      return DEC2HEX.calculate([matches[0], String(decimal), places ? String(places) : ''], context);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DELTA関数の実装（クロネッカーのデルタ）
export const DELTA: CustomFormula = {
  name: 'DELTA',
  pattern: /\bDELTA\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, num1Ref, num2Ref] = matches;
    
    try {
      const num1 = parseFloat(getCellValue(num1Ref.trim(), context)?.toString() ?? num1Ref.trim());
      const num2 = num2Ref ? 
        parseFloat(getCellValue(num2Ref.trim(), context)?.toString() ?? num2Ref.trim()) : 0;
      
      if (isNaN(num1) || isNaN(num2)) {
        return FormulaError.VALUE;
      }
      
      // 数値が等しい場合は1、そうでない場合は0を返す
      return num1 === num2 ? 1 : 0;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// GESTEP関数の実装（ステップ関数）
export const GESTEP: CustomFormula = {
  name: 'GESTEP',
  pattern: /\bGESTEP\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, numRef, stepRef] = matches;
    
    try {
      const num = parseFloat(getCellValue(numRef.trim(), context)?.toString() ?? numRef.trim());
      const step = stepRef ? 
        parseFloat(getCellValue(stepRef.trim(), context)?.toString() ?? stepRef.trim()) : 0;
      
      if (isNaN(num) || isNaN(step)) {
        return FormulaError.VALUE;
      }
      
      // 数値がステップ以上の場合は1、そうでない場合は0を返す
      return num >= step ? 1 : 0;
    } catch {
      return FormulaError.VALUE;
    }
  }
};