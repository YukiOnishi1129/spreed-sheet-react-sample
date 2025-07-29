// クーポン（利札）関連の財務関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';
import { parseDate } from '../shared/dateUtils';

// 日数計算規約（統一版）
function calculateDaysFraction(startDate: Date, endDate: Date, basis: number = 0): number {
  const earlier = startDate <= endDate ? startDate : endDate;
  const later = startDate <= endDate ? endDate : startDate;
  
  switch (basis) {
    case 0: // 30/360 US (NASD)
      return days30_360US(earlier, later);
    case 1: // Actual/actual - 実際の日数を返す
      return Math.floor((later.getTime() - earlier.getTime()) / (24 * 60 * 60 * 1000));
    case 2: // Actual/360 - 実際の日数を返す
      return Math.floor((later.getTime() - earlier.getTime()) / (24 * 60 * 60 * 1000));
    case 3: // Actual/365 - 実際の日数を返す
      return Math.floor((later.getTime() - earlier.getTime()) / (24 * 60 * 60 * 1000));
    case 4: // 30/360 European
      return days30_360European(earlier, later);
    default:
      return days30_360US(earlier, later);
  }
}

// 30/360 US (NASD) 方式の日数計算
function days30_360US(startDate: Date, endDate: Date): number {
  let startYear = startDate.getFullYear();
  let startMonth = startDate.getMonth() + 1;
  let startDay = startDate.getDate();
  
  let endYear = endDate.getFullYear();
  let endMonth = endDate.getMonth() + 1;
  let endDay = endDate.getDate();
  
  // US 30/360の特殊ルール
  if (startDay === getDaysInMonth(startYear, startMonth)) {
    startDay = 30;
  }
  
  if (endDay === getDaysInMonth(endYear, endMonth) && startDay < 30) {
    endDay = getDaysInMonth(endYear, endMonth);
  } else if (endDay === getDaysInMonth(endYear, endMonth) && startDay === 30) {
    endDay = 30;
  }
  
  if (startDay === 31 || (startMonth === 2 && startDay === getDaysInMonth(startYear, startMonth))) {
    startDay = 30;
  }
  
  if (endDay === 31 && startDay === 30) {
    endDay = 30;
  }
  
  return (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay);
}

// 30/360 European方式の日数計算
function days30_360European(startDate: Date, endDate: Date): number {
  let startYear = startDate.getFullYear();
  let startMonth = startDate.getMonth() + 1;
  let startDay = startDate.getDate();
  
  let endYear = endDate.getFullYear();
  let endMonth = endDate.getMonth() + 1;
  let endDay = endDate.getDate();
  
  // European 30/360のルール
  if (startDay === 31) startDay = 30;
  if (endDay === 31) endDay = 30;
  
  return (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay);
}

// ヘルパー関数：月の日数を取得
function getDaysInMonth(year: number, month: number): number {
  return new Date(year, month, 0).getDate();
}

// 年間日数を取得
function getDaysInYear(basis: number, year?: number): number {
  switch (basis) {
    case 0: // 30/360 US
    case 4: // 30/360 European
      return 360;
    case 1: // Actual/actual
      if (year) {
        return ((year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0)) ? 366 : 365;
      }
      return 365; // デフォルト
    case 2: // Actual/360
      return 360;
    case 3: // Actual/365
      return 365;
    default:
      return 360;
  }
}

// 次の利払日を計算
function getNextCouponDate(settlement: Date, maturity: Date, frequency: number): Date {
  const monthsPerPeriod = 12 / frequency;
  const currentDate = new Date(maturity);
  
  while (currentDate > settlement) {
    currentDate.setMonth(currentDate.getMonth() - monthsPerPeriod);
  }
  
  currentDate.setMonth(currentDate.getMonth() + monthsPerPeriod);
  return currentDate;
}

// 直前の利払日を計算
function getPreviousCouponDate(settlement: Date, maturity: Date, frequency: number): Date {
  const monthsPerPeriod = 12 / frequency;
  const currentDate = new Date(maturity);
  
  while (currentDate > settlement) {
    currentDate.setMonth(currentDate.getMonth() - monthsPerPeriod);
  }
  
  return currentDate;
}

// COUPDAYBS関数の実装（直前利払日から受渡日までの日数）
export const COUPDAYBS: CustomFormula = {
  name: 'COUPDAYBS',
  pattern: /\bCOUPDAYBS\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, freqRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      const previousCouponDate = getPreviousCouponDate(settlement, maturity, frequency);
      
      return calculateDaysFraction(previousCouponDate, settlement, basis);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// COUPDAYS関数の実装（利払期間の日数）
export const COUPDAYS: CustomFormula = {
  name: 'COUPDAYS',
  pattern: /\bCOUPDAYS\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, freqRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      const daysInYear = getDaysInYear(basis, maturity.getFullYear());
      
      // 利払期間の日数 = 年間日数 / 頻度
      return daysInYear / frequency;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// COUPDAYSNC関数の実装（受渡日から次回利払日までの日数）
export const COUPDAYSNC: CustomFormula = {
  name: 'COUPDAYSNC',
  pattern: /\bCOUPDAYSNC\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, freqRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      const nextCouponDate = getNextCouponDate(settlement, maturity, frequency);
      
      return calculateDaysFraction(settlement, nextCouponDate, basis);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// COUPNCD関数の実装（次回利払日）
export const COUPNCD: CustomFormula = {
  name: 'COUPNCD',
  pattern: /\bCOUPNCD\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, freqRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      // const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      const nextCouponDate = getNextCouponDate(settlement, maturity, frequency);
      
      // Excel形式の日付シリアル値として返す
      const excelEpoch = new Date(1900, 0, 1);
      const diffTime = nextCouponDate.getTime() - excelEpoch.getTime();
      const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24)) + 1;
      
      // Excel has a bug where it considers 1900 as a leap year
      if (diffDays > 59) {
        return diffDays + 1;
      }
      return diffDays;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// COUPNUM関数の実装（利払回数）
export const COUPNUM: CustomFormula = {
  name: 'COUPNUM',
  pattern: /\bCOUPNUM\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, freqRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      // const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      // 利払回数を計算
      const monthsPerPeriod = 12 / frequency;
      let count = 0;
      const currentDate = new Date(maturity);
      
      while (currentDate > settlement) {
        count++;
        currentDate.setMonth(currentDate.getMonth() - monthsPerPeriod);
      }
      
      return count;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// COUPPCD関数の実装（直前利払日）
export const COUPPCD: CustomFormula = {
  name: 'COUPPCD',
  pattern: /\bCOUPPCD\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, freqRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      // const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      const previousCouponDate = getPreviousCouponDate(settlement, maturity, frequency);
      
      // Excel形式の日付シリアル値として返す
      const excelEpoch = new Date(1900, 0, 1);
      const diffTime = previousCouponDate.getTime() - excelEpoch.getTime();
      const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24)) + 1;
      
      // Excel has a bug where it considers 1900 as a leap year
      if (diffDays > 59) {
        return diffDays + 1;
      }
      return diffDays;
    } catch {
      return FormulaError.VALUE;
    }
  }
};