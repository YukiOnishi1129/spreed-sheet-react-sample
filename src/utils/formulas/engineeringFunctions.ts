// エンジニアリング関数の実装

import type { CustomFormula, FormulaContext } from './types';
import { FormulaError } from './types';
import { getCellValue } from './utils';

// BIN2DEC関数の実装（2進数→10進数）
export const BIN2DEC: CustomFormula = {
  name: 'BIN2DEC',
  pattern: /BIN2DEC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const binaryStr = String(getCellValue(numberRef, context) ?? numberRef).trim();
    
    // 2進数として有効かチェック（0と1のみ、最大10桁）
    if (!/^[01]{1,10}$/.test(binaryStr)) {
      return FormulaError.NUM;
    }
    
    // 2進数を10進数に変換
    const result = parseInt(binaryStr, 2);
    
    // 10桁の2進数は-512から511の範囲
    if (binaryStr.length === 10 && binaryStr.startsWith('1')) {
      // 負数の場合（2の補数）
      return result - 1024;
    }
    
    return result;
  }
};

// DEC2BIN関数の実装（10進数→2進数）
export const DEC2BIN: CustomFormula = {
  name: 'DEC2BIN',
  pattern: /DEC2BIN\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, placesRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const places = placesRef ? Number(getCellValue(placesRef, context) ?? placesRef) : undefined;
    
    if (isNaN(number)) {
      return FormulaError.VALUE;
    }
    
    // -512から511の範囲チェック
    if (number < -512 || number > 511 || !Number.isInteger(number)) {
      return FormulaError.NUM;
    }
    
    if (places !== undefined && (isNaN(places) || places < 1 || places > 10 || !Number.isInteger(places))) {
      return FormulaError.NUM;
    }
    
    let binaryStr: string;
    
    if (number >= 0) {
      binaryStr = number.toString(2);
    } else {
      // 負数の場合（2の補数）
      binaryStr = (1024 + number).toString(2);
    }
    
    // 桁数指定がある場合
    if (places !== undefined) {
      if (binaryStr.length > places) {
        return FormulaError.NUM;
      }
      binaryStr = binaryStr.padStart(places, '0');
    }
    
    return binaryStr;
  }
};

// BIN2HEX関数の実装（2進数→16進数）
export const BIN2HEX: CustomFormula = {
  name: 'BIN2HEX',
  pattern: /BIN2HEX\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, placesRef] = matches;
    
    const binaryStr = String(getCellValue(numberRef, context) ?? numberRef).trim();
    const places = placesRef ? Number(getCellValue(placesRef, context) ?? placesRef) : undefined;
    
    if (!/^[01]{1,10}$/.test(binaryStr)) {
      return FormulaError.NUM;
    }
    
    if (places !== undefined && (isNaN(places) || places < 1 || places > 10 || !Number.isInteger(places))) {
      return FormulaError.NUM;
    }
    
    let decimal = parseInt(binaryStr, 2);
    
    // 10桁の2進数の場合の負数処理
    if (binaryStr.length === 10 && binaryStr.startsWith('1')) {
      decimal = decimal - 1024;
    }
    
    let hexStr: string;
    
    if (decimal >= 0) {
      hexStr = decimal.toString(16).toUpperCase();
    } else {
      // 負数の場合
      hexStr = (Math.pow(2, 40) + decimal).toString(16).toUpperCase().slice(-10);
    }
    
    if (places !== undefined) {
      if (hexStr.length > places) {
        return FormulaError.NUM;
      }
      hexStr = hexStr.padStart(places, '0');
    }
    
    return hexStr;
  }
};

// BIN2OCT関数の実装（2進数→8進数）
export const BIN2OCT: CustomFormula = {
  name: 'BIN2OCT',
  pattern: /BIN2OCT\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, placesRef] = matches;
    
    const binaryStr = String(getCellValue(numberRef, context) ?? numberRef).trim();
    const places = placesRef ? Number(getCellValue(placesRef, context) ?? placesRef) : undefined;
    
    if (!/^[01]{1,10}$/.test(binaryStr)) {
      return FormulaError.NUM;
    }
    
    if (places !== undefined && (isNaN(places) || places < 1 || places > 10 || !Number.isInteger(places))) {
      return FormulaError.NUM;
    }
    
    let decimal = parseInt(binaryStr, 2);
    
    // 10桁の2進数の場合の負数処理
    if (binaryStr.length === 10 && binaryStr.startsWith('1')) {
      decimal = decimal - 1024;
    }
    
    let octStr: string;
    
    if (decimal >= 0) {
      octStr = decimal.toString(8);
    } else {
      // 負数の場合
      octStr = (Math.pow(2, 30) + decimal).toString(8).slice(-10);
    }
    
    if (places !== undefined) {
      if (octStr.length > places) {
        return FormulaError.NUM;
      }
      octStr = octStr.padStart(places, '0');
    }
    
    return octStr;
  }
};

// DEC2HEX関数の実装（10進数→16進数）
export const DEC2HEX: CustomFormula = {
  name: 'DEC2HEX',
  pattern: /DEC2HEX\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, placesRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const places = placesRef ? Number(getCellValue(placesRef, context) ?? placesRef) : undefined;
    
    if (isNaN(number) || !Number.isInteger(number)) {
      return FormulaError.VALUE;
    }
    
    // -549755813888から549755813887の範囲チェック
    if (number < -549755813888 || number > 549755813887) {
      return FormulaError.NUM;
    }
    
    if (places !== undefined && (isNaN(places) || places < 1 || places > 10 || !Number.isInteger(places))) {
      return FormulaError.NUM;
    }
    
    let hexStr: string;
    
    if (number >= 0) {
      hexStr = number.toString(16).toUpperCase();
    } else {
      // 負数の場合
      hexStr = (Math.pow(2, 40) + number).toString(16).toUpperCase().slice(-10);
    }
    
    if (places !== undefined) {
      if (hexStr.length > places) {
        return FormulaError.NUM;
      }
      hexStr = hexStr.padStart(places, '0');
    }
    
    return hexStr;
  }
};

// DEC2OCT関数の実装（10進数→8進数）
export const DEC2OCT: CustomFormula = {
  name: 'DEC2OCT',
  pattern: /DEC2OCT\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, placesRef] = matches;
    
    const number = Number(getCellValue(numberRef, context) ?? numberRef);
    const places = placesRef ? Number(getCellValue(placesRef, context) ?? placesRef) : undefined;
    
    if (isNaN(number) || !Number.isInteger(number)) {
      return FormulaError.VALUE;
    }
    
    // -536870912から536870911の範囲チェック
    if (number < -536870912 || number > 536870911) {
      return FormulaError.NUM;
    }
    
    if (places !== undefined && (isNaN(places) || places < 1 || places > 10 || !Number.isInteger(places))) {
      return FormulaError.NUM;
    }
    
    let octStr: string;
    
    if (number >= 0) {
      octStr = number.toString(8);
    } else {
      // 負数の場合  
      octStr = (Math.pow(2, 30) + number).toString(8).slice(-10);
    }
    
    if (places !== undefined) {
      if (octStr.length > places) {
        return FormulaError.NUM;
      }
      octStr = octStr.padStart(places, '0');
    }
    
    return octStr;
  }
};

// HEX2BIN関数の実装（16進数→2進数）
export const HEX2BIN: CustomFormula = {
  name: 'HEX2BIN',
  pattern: /HEX2BIN\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, placesRef] = matches;
    
    const hexStr = String(getCellValue(numberRef, context) ?? numberRef).trim().toUpperCase();
    const places = placesRef ? Number(getCellValue(placesRef, context) ?? placesRef) : undefined;
    
    if (!/^[0-9A-F]{1,10}$/.test(hexStr)) {
      return FormulaError.NUM;
    }
    
    if (places !== undefined && (isNaN(places) || places < 1 || places > 10 || !Number.isInteger(places))) {
      return FormulaError.NUM;
    }
    
    const decimal = parseInt(hexStr, 16);
    
    // 2進数として表現可能な範囲チェック（-512から511）
    let finalDecimal = decimal;
    if (hexStr.length === 10 && parseInt(hexStr[0], 16) >= 8) {
      // 負数の場合
      finalDecimal = decimal - Math.pow(16, 10);
    }
    
    if (finalDecimal < -512 || finalDecimal > 511) {
      return FormulaError.NUM;
    }
    
    let binaryStr: string;
    
    if (finalDecimal >= 0) {
      binaryStr = finalDecimal.toString(2);
    } else {
      // 負数の場合（2の補数）
      binaryStr = (1024 + finalDecimal).toString(2);
    }
    
    if (places !== undefined) {
      if (binaryStr.length > places) {
        return FormulaError.NUM;
      }
      binaryStr = binaryStr.padStart(places, '0');
    }
    
    return binaryStr;
  }
};

// HEX2DEC関数の実装（16進数→10進数）
export const HEX2DEC: CustomFormula = {
  name: 'HEX2DEC',
  pattern: /HEX2DEC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const hexStr = String(getCellValue(numberRef, context) ?? numberRef).trim().toUpperCase();
    
    if (!/^[0-9A-F]{1,10}$/.test(hexStr)) {
      return FormulaError.NUM;
    }
    
    const decimal = parseInt(hexStr, 16);
    
    // 10桁の16進数の場合の負数処理
    if (hexStr.length === 10 && parseInt(hexStr[0], 16) >= 8) {
      return decimal - Math.pow(16, 10);
    }
    
    return decimal;
  }
};

// HEX2OCT関数の実装（16進数→8進数）
export const HEX2OCT: CustomFormula = {
  name: 'HEX2OCT',
  pattern: /HEX2OCT\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, placesRef] = matches;
    
    const hexStr = String(getCellValue(numberRef, context) ?? numberRef).trim().toUpperCase();
    const places = placesRef ? Number(getCellValue(placesRef, context) ?? placesRef) : undefined;
    
    if (!/^[0-9A-F]{1,10}$/.test(hexStr)) {
      return FormulaError.NUM;
    }
    
    if (places !== undefined && (isNaN(places) || places < 1 || places > 10 || !Number.isInteger(places))) {
      return FormulaError.NUM;
    }
    
    let decimal = parseInt(hexStr, 16);
    
    // 10桁の16進数の場合の負数処理
    if (hexStr.length === 10 && parseInt(hexStr[0], 16) >= 8) {
      decimal = decimal - Math.pow(16, 10);
    }
    
    // 8進数として表現可能な範囲チェック
    if (decimal < -536870912 || decimal > 536870911) {
      return FormulaError.NUM;
    }
    
    let octStr: string;
    
    if (decimal >= 0) {
      octStr = decimal.toString(8);
    } else {
      // 負数の場合
      octStr = (Math.pow(2, 30) + decimal).toString(8).slice(-10);
    }
    
    if (places !== undefined) {
      if (octStr.length > places) {
        return FormulaError.NUM;
      }
      octStr = octStr.padStart(places, '0');
    }
    
    return octStr;
  }
};

// OCT2BIN関数の実装（8進数→2進数）
export const OCT2BIN: CustomFormula = {
  name: 'OCT2BIN',
  pattern: /OCT2BIN\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, placesRef] = matches;
    
    const octStr = String(getCellValue(numberRef, context) ?? numberRef).trim();
    const places = placesRef ? Number(getCellValue(placesRef, context) ?? placesRef) : undefined;
    
    if (!/^[0-7]{1,10}$/.test(octStr)) {
      return FormulaError.NUM;
    }
    
    if (places !== undefined && (isNaN(places) || places < 1 || places > 10 || !Number.isInteger(places))) {
      return FormulaError.NUM;
    }
    
    let decimal = parseInt(octStr, 8);
    
    // 10桁の8進数の場合の負数処理
    if (octStr.length === 10 && parseInt(octStr[0]) >= 4) {
      decimal = decimal - Math.pow(8, 10);
    }
    
    // 2進数として表現可能な範囲チェック（-512から511）
    if (decimal < -512 || decimal > 511) {
      return FormulaError.NUM;
    }
    
    let binaryStr: string;
    
    if (decimal >= 0) {
      binaryStr = decimal.toString(2);
    } else {
      // 負数の場合（2の補数）
      binaryStr = (1024 + decimal).toString(2);
    }
    
    if (places !== undefined) {
      if (binaryStr.length > places) {
        return FormulaError.NUM;
      }
      binaryStr = binaryStr.padStart(places, '0');
    }
    
    return binaryStr;
  }
};

// OCT2DEC関数の実装（8進数→10進数）
export const OCT2DEC: CustomFormula = {
  name: 'OCT2DEC',
  pattern: /OCT2DEC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef] = matches;
    
    const octStr = String(getCellValue(numberRef, context) ?? numberRef).trim();
    
    if (!/^[0-7]{1,10}$/.test(octStr)) {
      return FormulaError.NUM;
    }
    
    const decimal = parseInt(octStr, 8);
    
    // 10桁の8進数の場合の負数処理
    if (octStr.length === 10 && parseInt(octStr[0]) >= 4) {
      return decimal - Math.pow(8, 10);
    }
    
    return decimal;
  }
};

// OCT2HEX関数の実装（8進数→16進数）
export const OCT2HEX: CustomFormula = {
  name: 'OCT2HEX',
  pattern: /OCT2HEX\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, numberRef, placesRef] = matches;
    
    const octStr = String(getCellValue(numberRef, context) ?? numberRef).trim();
    const places = placesRef ? Number(getCellValue(placesRef, context) ?? placesRef) : undefined;
    
    if (!/^[0-7]{1,10}$/.test(octStr)) {
      return FormulaError.NUM;
    }
    
    if (places !== undefined && (isNaN(places) || places < 1 || places > 10 || !Number.isInteger(places))) {
      return FormulaError.NUM;
    }
    
    let decimal = parseInt(octStr, 8);
    
    // 10桁の8進数の場合の負数処理
    if (octStr.length === 10 && parseInt(octStr[0]) >= 4) {
      decimal = decimal - Math.pow(8, 10);
    }
    
    let hexStr: string;
    
    if (decimal >= 0) {
      hexStr = decimal.toString(16).toUpperCase();
    } else {
      // 負数の場合
      hexStr = (Math.pow(2, 40) + decimal).toString(16).toUpperCase().slice(-10);
    }
    
    if (places !== undefined) {
      if (hexStr.length > places) {
        return FormulaError.NUM;
      }
      hexStr = hexStr.padStart(places, '0');
    }
    
    return hexStr;
  }
};