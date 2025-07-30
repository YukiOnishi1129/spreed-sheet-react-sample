// 複素数の三角関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// 複素数を表す型
interface ComplexNumber {
  real: number;
  imaginary: number;
}

// 複素数を文字列から解析
function parseComplex(value: string): ComplexNumber | null {
  const str = value.toString().trim().replace(/"/g, '');
  
  // 純実数の場合
  const realOnly = parseFloat(str);
  if (!isNaN(realOnly) && !str.includes('i') && !str.includes('j')) {
    return { real: realOnly, imaginary: 0 };
  }
  
  // 単独のi/jの場合
  if (str === 'i' || str === 'j') {
    return { real: 0, imaginary: 1 };
  }
  if (str === '-i' || str === '-j') {
    return { real: 0, imaginary: -1 };
  }
  if (str === '+i' || str === '+j') {
    return { real: 0, imaginary: 1 };
  }
  
  // 純虚数の場合（例：3i, -2j, 1.5i）
  const pureImagMatch = str.match(/^([+-]?\d*\.?\d*)([ij])$/);
  if (pureImagMatch) {
    const imagPart = pureImagMatch[1];
    let imag: number;
    if (imagPart === '' || imagPart === '+') {
      imag = 1;
    } else if (imagPart === '-') {
      imag = -1;
    } else {
      imag = parseFloat(imagPart);
    }
    return { real: 0, imaginary: imag };
  }
  
  // 複素数の場合（例：3+4i, -2-3j, 1.5-2.3i）
  const complexMatch = str.match(/^([+-]?\d*\.?\d+)\s*([+-])\s*(\d*\.?\d*)([ij])$/);
  if (complexMatch) {
    const real = parseFloat(complexMatch[1]);
    const sign = complexMatch[2] === '-' ? -1 : 1;
    const imagPart = complexMatch[3];
    let imag: number;
    if (imagPart === '' || imagPart === '+') {
      imag = sign * 1;
    } else {
      imag = sign * parseFloat(imagPart);
    }
    return { real, imaginary: imag };
  }
  
  // より複雑なパターンのための追加マッチング
  const altComplexMatch = str.match(/^([+-]?\d*\.?\d*)\s*([+-])\s*([+-]?\d*\.?\d*)([ij])$/);
  if (altComplexMatch) {
    const realPart = altComplexMatch[1];
    const sign = altComplexMatch[2] === '-' ? -1 : 1;
    const imagPart = altComplexMatch[3];
    
    let real = 0;
    if (realPart !== '' && realPart !== '+' && realPart !== '-') {
      real = parseFloat(realPart);
    }
    
    let imag: number;
    if (imagPart === '' || imagPart === '+') {
      imag = sign * 1;
    } else if (imagPart === '-') {
      imag = sign * -1;
    } else {
      imag = sign * parseFloat(imagPart);
    }
    
    return { real, imaginary: imag };
  }
  
  return null;
}

// 複素数を文字列に変換
function complexToString(complex: ComplexNumber, suffix: string = 'i'): string {
  if (complex.imaginary === 0) {
    return complex.real.toString();
  }
  
  if (complex.real === 0) {
    if (complex.imaginary === 1) return suffix;
    if (complex.imaginary === -1) return `-${suffix}`;
    return `${complex.imaginary}${suffix}`;
  }
  
  const sign = complex.imaginary >= 0 ? '+' : '';
  if (Math.abs(complex.imaginary) === 1) {
    return `${complex.real}${sign}${complex.imaginary < 0 ? '-' : ''}${suffix}`;
  }
  return `${complex.real}${sign}${complex.imaginary}${suffix}`;
}

// IMSIN関数の実装（複素数の正弦）
export const IMSIN: CustomFormula = {
  name: 'IMSIN',
  pattern: /IMSIN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // sin(a + bi) = sin(a)cosh(b) + i*cos(a)sinh(b)
      const real = Math.sin(complex.real) * Math.cosh(complex.imaginary);
      const imag = Math.cos(complex.real) * Math.sinh(complex.imaginary);
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMCOS関数の実装（複素数の余弦）
export const IMCOS: CustomFormula = {
  name: 'IMCOS',
  pattern: /IMCOS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // cos(a + bi) = cos(a)cosh(b) - i*sin(a)sinh(b)
      const real = Math.cos(complex.real) * Math.cosh(complex.imaginary);
      const imag = -Math.sin(complex.real) * Math.sinh(complex.imaginary);
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMTAN関数の実装（複素数の正接）
export const IMTAN: CustomFormula = {
  name: 'IMTAN',
  pattern: /IMTAN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // tan(z) = sin(z) / cos(z)
      // 分母の計算
      const denom = Math.cos(2 * complex.real) + Math.cosh(2 * complex.imaginary);
      
      if (Math.abs(denom) < 1e-15) {
        return FormulaError.NUM;
      }
      
      const real = Math.sin(2 * complex.real) / denom;
      const imag = Math.sinh(2 * complex.imaginary) / denom;
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMSINH関数の実装（複素数の双曲線正弦）
export const IMSINH: CustomFormula = {
  name: 'IMSINH',
  pattern: /IMSINH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // sinh(a + bi) = sinh(a)cos(b) + i*cosh(a)sin(b)
      const real = Math.sinh(complex.real) * Math.cos(complex.imaginary);
      const imag = Math.cosh(complex.real) * Math.sin(complex.imaginary);
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMCOSH関数の実装（複素数の双曲線余弦）
export const IMCOSH: CustomFormula = {
  name: 'IMCOSH',
  pattern: /IMCOSH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // cosh(a + bi) = cosh(a)cos(b) + i*sinh(a)sin(b)
      const real = Math.cosh(complex.real) * Math.cos(complex.imaginary);
      const imag = Math.sinh(complex.real) * Math.sin(complex.imaginary);
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMCSC関数の実装（複素数の余割）
export const IMCSC: CustomFormula = {
  name: 'IMCSC',
  pattern: /IMCSC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // csc(z) = 1 / sin(z)
      const sinReal = Math.sin(complex.real) * Math.cosh(complex.imaginary);
      const sinImag = Math.cos(complex.real) * Math.sinh(complex.imaginary);
      
      const denom = sinReal * sinReal + sinImag * sinImag;
      
      if (Math.abs(denom) < 1e-15) {
        return FormulaError.NUM;
      }
      
      const real = sinReal / denom;
      const imag = -sinImag / denom;
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMSEC関数の実装（複素数の正割）
export const IMSEC: CustomFormula = {
  name: 'IMSEC',
  pattern: /IMSEC\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // sec(z) = 1 / cos(z)
      const cosReal = Math.cos(complex.real) * Math.cosh(complex.imaginary);
      const cosImag = -Math.sin(complex.real) * Math.sinh(complex.imaginary);
      
      const denom = cosReal * cosReal + cosImag * cosImag;
      
      if (Math.abs(denom) < 1e-15) {
        return FormulaError.NUM;
      }
      
      const real = cosReal / denom;
      const imag = -cosImag / denom;
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMCOT関数の実装（複素数の余接）
export const IMCOT: CustomFormula = {
  name: 'IMCOT',
  pattern: /IMCOT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // cot(z) = cos(z) / sin(z)
      // Using the identity: cot(a+bi) = (sin(2a) - i*sinh(2b)) / (cosh(2b) - cos(2a))
      const denom = Math.cosh(2 * complex.imaginary) - Math.cos(2 * complex.real);
      
      if (Math.abs(denom) < 1e-15) {
        return FormulaError.NUM;
      }
      
      const real = Math.sin(2 * complex.real) / denom;
      const imag = -Math.sinh(2 * complex.imaginary) / denom;
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMCSCH関数の実装（複素数の双曲線余割）
export const IMCSCH: CustomFormula = {
  name: 'IMCSCH',
  pattern: /IMCSCH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // csch(z) = 1 / sinh(z)
      const sinhReal = Math.sinh(complex.real) * Math.cos(complex.imaginary);
      const sinhImag = Math.cosh(complex.real) * Math.sin(complex.imaginary);
      
      const denom = sinhReal * sinhReal + sinhImag * sinhImag;
      
      if (Math.abs(denom) < 1e-15) {
        return FormulaError.NUM;
      }
      
      const real = sinhReal / denom;
      const imag = -sinhImag / denom;
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMTANH関数の実装（複素数の双曲線正接）
export const IMTANH: CustomFormula = {
  name: 'IMTANH',
  pattern: /IMTANH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // tanh(z) = sinh(z) / cosh(z)
      // Using the identity: tanh(a+bi) = (sinh(2a) + i*sin(2b)) / (cosh(2a) + cos(2b))
      const denom = Math.cosh(2 * complex.real) + Math.cos(2 * complex.imaginary);
      
      if (Math.abs(denom) < 1e-15) {
        return FormulaError.NUM;
      }
      
      const real = Math.sinh(2 * complex.real) / denom;
      const imag = Math.sin(2 * complex.imaginary) / denom;
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMSECH関数の実装（複素数の双曲線正割）
export const IMSECH: CustomFormula = {
  name: 'IMSECH',
  pattern: /IMSECH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // sech(z) = 1 / cosh(z)
      const coshReal = Math.cosh(complex.real) * Math.cos(complex.imaginary);
      const coshImag = Math.sinh(complex.real) * Math.sin(complex.imaginary);
      
      const denom = coshReal * coshReal + coshImag * coshImag;
      
      if (Math.abs(denom) < 1e-15) {
        return FormulaError.NUM;
      }
      
      const real = coshReal / denom;
      const imag = -coshImag / denom;
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMCOTH関数の実装（複素数の双曲線余接）
export const IMCOTH: CustomFormula = {
  name: 'IMCOTH',
  pattern: /IMCOTH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // coth(z) = cosh(z) / sinh(z)
      // Using the identity: coth(a+bi) = (sinh(2a) - i*sin(2b)) / (cosh(2a) - cos(2b))
      const denom = Math.cosh(2 * complex.real) - Math.cos(2 * complex.imaginary);
      
      if (Math.abs(denom) < 1e-15) {
        return FormulaError.NUM;
      }
      
      const real = Math.sinh(2 * complex.real) / denom;
      const imag = -Math.sin(2 * complex.imaginary) / denom;
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};