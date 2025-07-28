// 複素数関連のエンジニアリング関数

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
  const str = value.toString().trim();
  
  // 純実数の場合
  const realOnly = parseFloat(str);
  if (!isNaN(realOnly) && !str.includes('i') && !str.includes('j')) {
    return { real: realOnly, imaginary: 0 };
  }
  
  // 純虚数の場合（例：3i, -2j）
  const pureImagMatch = str.match(/^([+-]?\d*\.?\d*)([ij])$/);
  if (pureImagMatch) {
    const imagPart = pureImagMatch[1] || '1';
    const imag = imagPart === '-' ? -1 : (imagPart === '+' || imagPart === '' ? 1 : parseFloat(imagPart));
    return { real: 0, imaginary: imag };
  }
  
  // 複素数の場合（例：3+4i, -2-3j）
  const complexMatch = str.match(/^([+-]?\d*\.?\d+)\s*([+-])\s*(\d*\.?\d*)([ij])$/);
  if (complexMatch) {
    const real = parseFloat(complexMatch[1]);
    const sign = complexMatch[2] === '-' ? -1 : 1;
    const imagPart = complexMatch[3] || '1';
    const imag = sign * (imagPart === '' ? 1 : parseFloat(imagPart));
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

// COMPLEX関数の実装（複素数を作成）
export const COMPLEX: CustomFormula = {
  name: 'COMPLEX',
  pattern: /COMPLEX\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, realRef, imagRef, suffixRef] = matches;
    
    try {
      const real = parseFloat(getCellValue(realRef.trim(), context)?.toString() ?? realRef.trim());
      const imaginary = parseFloat(getCellValue(imagRef.trim(), context)?.toString() ?? imagRef.trim());
      let suffix = suffixRef ? getCellValue(suffixRef.trim(), context)?.toString() ?? suffixRef.trim() : 'i';
      
      if (isNaN(real) || isNaN(imaginary)) {
        return FormulaError.VALUE;
      }
      
      // サフィックスの検証
      suffix = suffix.replace(/['"]/g, '').toLowerCase();
      if (suffix !== 'i' && suffix !== 'j') {
        return FormulaError.VALUE;
      }
      
      return complexToString({ real, imaginary }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMABS関数の実装（複素数の絶対値）
export const IMABS: CustomFormula = {
  name: 'IMABS',
  pattern: /IMABS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      // |a + bi| = √(a² + b²)
      return Math.sqrt(complex.real * complex.real + complex.imaginary * complex.imaginary);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMAGINARY関数の実装（虚部を返す）
export const IMAGINARY: CustomFormula = {
  name: 'IMAGINARY',
  pattern: /IMAGINARY\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      return complex.imaginary;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMREAL関数の実装（実部を返す）
export const IMREAL: CustomFormula = {
  name: 'IMREAL',
  pattern: /IMREAL\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      return complex.real;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMARGUMENT関数の実装（偏角を返す）
export const IMARGUMENT: CustomFormula = {
  name: 'IMARGUMENT',
  pattern: /IMARGUMENT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      if (complex.real === 0 && complex.imaginary === 0) {
        return FormulaError.DIV0;
      }
      
      // arg(a + bi) = atan2(b, a)
      return Math.atan2(complex.imaginary, complex.real);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMCONJUGATE関数の実装（複素共役）
export const IMCONJUGATE: CustomFormula = {
  name: 'IMCONJUGATE',
  pattern: /IMCONJUGATE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      // 共役 a + bi → a - bi
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      return complexToString({ real: complex.real, imaginary: -complex.imaginary }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMSUM関数の実装（複素数の和）
export const IMSUM: CustomFormula = {
  name: 'IMSUM',
  pattern: /IMSUM\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, argsStr] = matches;
    
    try {
      const args = argsStr.split(',').map(arg => arg.trim());
      let sumReal = 0;
      let sumImag = 0;
      let suffix = 'i';
      
      for (const arg of args) {
        const complexStr = getCellValue(arg, context)?.toString() ?? arg;
        const complex = parseComplex(complexStr);
        
        if (!complex) {
          return FormulaError.VALUE;
        }
        
        sumReal += complex.real;
        sumImag += complex.imaginary;
        
        if (complexStr.includes('j')) {
          suffix = 'j';
        }
      }
      
      return complexToString({ real: sumReal, imaginary: sumImag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMSUB関数の実装（複素数の減算）
export const IMSUB: CustomFormula = {
  name: 'IMSUB',
  pattern: /IMSUB\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complex1Ref, complex2Ref] = matches;
    
    try {
      const complex1Str = getCellValue(complex1Ref.trim(), context)?.toString() ?? complex1Ref.trim();
      const complex2Str = getCellValue(complex2Ref.trim(), context)?.toString() ?? complex2Ref.trim();
      
      const complex1 = parseComplex(complex1Str);
      const complex2 = parseComplex(complex2Str);
      
      if (!complex1 || !complex2) {
        return FormulaError.NUM;
      }
      
      const suffix = complex1Str.includes('j') || complex2Str.includes('j') ? 'j' : 'i';
      
      return complexToString({
        real: complex1.real - complex2.real,
        imaginary: complex1.imaginary - complex2.imaginary
      }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMPRODUCT関数の実装（複素数の積）
export const IMPRODUCT: CustomFormula = {
  name: 'IMPRODUCT',
  pattern: /IMPRODUCT\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, argsStr] = matches;
    
    try {
      const args = argsStr.split(',').map(arg => arg.trim());
      let productReal = 1;
      let productImag = 0;
      let suffix = 'i';
      let first = true;
      
      for (const arg of args) {
        const complexStr = getCellValue(arg, context)?.toString() ?? arg;
        const complex = parseComplex(complexStr);
        
        if (!complex) {
          return FormulaError.VALUE;
        }
        
        if (first) {
          productReal = complex.real;
          productImag = complex.imaginary;
          first = false;
        } else {
          // (a + bi)(c + di) = (ac - bd) + (ad + bc)i
          const newReal = productReal * complex.real - productImag * complex.imaginary;
          const newImag = productReal * complex.imaginary + productImag * complex.real;
          productReal = newReal;
          productImag = newImag;
        }
        
        if (complexStr.includes('j')) {
          suffix = 'j';
        }
      }
      
      return complexToString({ real: productReal, imaginary: productImag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMDIV関数の実装（複素数の除算）
export const IMDIV: CustomFormula = {
  name: 'IMDIV',
  pattern: /IMDIV\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, numerRef, denomRef] = matches;
    
    try {
      const numerStr = getCellValue(numerRef.trim(), context)?.toString() ?? numerRef.trim();
      const denomStr = getCellValue(denomRef.trim(), context)?.toString() ?? denomRef.trim();
      
      const numer = parseComplex(numerStr);
      const denom = parseComplex(denomStr);
      
      if (!numer || !denom) {
        return FormulaError.NUM;
      }
      
      // ゼロ除算チェック
      if (denom.real === 0 && denom.imaginary === 0) {
        return FormulaError.DIV0;
      }
      
      const suffix = numerStr.includes('j') || denomStr.includes('j') ? 'j' : 'i';
      
      // (a + bi) / (c + di) = ((ac + bd) + (bc - ad)i) / (c² + d²)
      const denomSquared = denom.real * denom.real + denom.imaginary * denom.imaginary;
      const real = (numer.real * denom.real + numer.imaginary * denom.imaginary) / denomSquared;
      const imag = (numer.imaginary * denom.real - numer.real * denom.imaginary) / denomSquared;
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMPOWER関数の実装（複素数のべき乗）
export const IMPOWER: CustomFormula = {
  name: 'IMPOWER',
  pattern: /IMPOWER\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef, powerRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const power = parseFloat(getCellValue(powerRef.trim(), context)?.toString() ?? powerRef.trim());
      
      const complex = parseComplex(complexStr);
      
      if (!complex || isNaN(power)) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // 極形式に変換: r * e^(iθ)
      const r = Math.sqrt(complex.real * complex.real + complex.imaginary * complex.imaginary);
      const theta = Math.atan2(complex.imaginary, complex.real);
      
      // べき乗を計算: (r * e^(iθ))^n = r^n * e^(inθ)
      const newR = Math.pow(r, power);
      const newTheta = theta * power;
      
      // 直交形式に戻す
      const real = newR * Math.cos(newTheta);
      const imag = newR * Math.sin(newTheta);
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMSQRT関数の実装（複素数の平方根）
export const IMSQRT: CustomFormula = {
  name: 'IMSQRT',
  pattern: /IMSQRT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // √(a + bi) の計算
      const r = Math.sqrt(complex.real * complex.real + complex.imaginary * complex.imaginary);
      const real = Math.sqrt((r + complex.real) / 2);
      const imag = Math.sign(complex.imaginary) * Math.sqrt((r - complex.real) / 2);
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMEXP関数の実装（複素数の指数関数）
export const IMEXP: CustomFormula = {
  name: 'IMEXP',
  pattern: /IMEXP\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // e^(a + bi) = e^a * (cos(b) + i*sin(b))
      const expReal = Math.exp(complex.real);
      const real = expReal * Math.cos(complex.imaginary);
      const imag = expReal * Math.sin(complex.imaginary);
      
      return complexToString({ real, imaginary: imag }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMLN関数の実装（複素数の自然対数）
export const IMLN: CustomFormula = {
  name: 'IMLN',
  pattern: /IMLN\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      if (complex.real === 0 && complex.imaginary === 0) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // ln(a + bi) = ln|a + bi| + i*arg(a + bi)
      const r = Math.sqrt(complex.real * complex.real + complex.imaginary * complex.imaginary);
      const theta = Math.atan2(complex.imaginary, complex.real);
      
      return complexToString({ real: Math.log(r), imaginary: theta }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMLOG10関数の実装（複素数の常用対数）
export const IMLOG10: CustomFormula = {
  name: 'IMLOG10',
  pattern: /IMLOG10\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      if (complex.real === 0 && complex.imaginary === 0) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // log10(z) = ln(z) / ln(10)
      const r = Math.sqrt(complex.real * complex.real + complex.imaginary * complex.imaginary);
      const theta = Math.atan2(complex.imaginary, complex.real);
      const ln10 = Math.log(10);
      
      return complexToString({ real: Math.log(r) / ln10, imaginary: theta / ln10 }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMLOG2関数の実装（複素数の2を底とする対数）
export const IMLOG2: CustomFormula = {
  name: 'IMLOG2',
  pattern: /IMLOG2\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, complexRef] = matches;
    
    try {
      const complexStr = getCellValue(complexRef.trim(), context)?.toString() ?? complexRef.trim();
      const complex = parseComplex(complexStr);
      
      if (!complex) {
        return FormulaError.NUM;
      }
      
      if (complex.real === 0 && complex.imaginary === 0) {
        return FormulaError.NUM;
      }
      
      const suffix = complexStr.includes('j') ? 'j' : 'i';
      
      // log2(z) = ln(z) / ln(2)
      const r = Math.sqrt(complex.real * complex.real + complex.imaginary * complex.imaginary);
      const theta = Math.atan2(complex.imaginary, complex.real);
      const ln2 = Math.log(2);
      
      return complexToString({ real: Math.log(r) / ln2, imaginary: theta / ln2 }, suffix);
    } catch {
      return FormulaError.VALUE;
    }
  }
};