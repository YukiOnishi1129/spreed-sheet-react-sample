// クーポン（利札）関連の財務関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';
import { parseDate } from '../shared/dateUtils';

// 日付ベースの日数計算
function getDayCountBasis(basis: number = 0): { daysInYear: number; daysInMonth: number } {
  switch (basis) {
    case 0: // 30/360 US
    case 4: // 30/360 European
      return { daysInYear: 360, daysInMonth: 30 };
    case 1: // Actual/actual
      return { daysInYear: 365, daysInMonth: 0 }; // 実際の日数
    case 2: // Actual/360
      return { daysInYear: 360, daysInMonth: 0 };
    case 3: // Actual/365
      return { daysInYear: 365, daysInMonth: 0 };
    default:
      return { daysInYear: 360, daysInMonth: 30 };
  }
}

// 30/360方式での日数計算
function days360(startDate: Date, endDate: Date, european: boolean = false): number {
  let startDay = startDate.getDate();
  let startMonth = startDate.getMonth() + 1;
  let startYear = startDate.getFullYear();
  let endDay = endDate.getDate();
  let endMonth = endDate.getMonth() + 1;
  let endYear = endDate.getFullYear();

  if (!european) {
    // US方式
    if (startDay === 31) startDay = 30;
    if (endDay === 31 && startDay === 30) endDay = 30;
  } else {
    // European方式
    if (startDay === 31) startDay = 30;
    if (endDay === 31) endDay = 30;
  }

  return (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay);
}

// 実際の日数計算
function actualDays(startDate: Date, endDate: Date): number {
  const diffTime = endDate.getTime() - startDate.getTime();
  return Math.floor(diffTime / (1000 * 60 * 60 * 24));
}

// 次の利払日を計算
function getNextCouponDate(settlement: Date, maturity: Date, frequency: number): Date {
  const monthsPerPeriod = 12 / frequency;
  let currentDate = new Date(maturity);
  
  while (currentDate > settlement) {
    currentDate.setMonth(currentDate.getMonth() - monthsPerPeriod);
  }
  
  currentDate.setMonth(currentDate.getMonth() + monthsPerPeriod);
  return currentDate;
}

// 直前の利払日を計算
function getPreviousCouponDate(settlement: Date, maturity: Date, frequency: number): Date {
  const monthsPerPeriod = 12 / frequency;
  let currentDate = new Date(maturity);
  
  while (currentDate > settlement) {
    currentDate.setMonth(currentDate.getMonth() - monthsPerPeriod);
  }
  
  return currentDate;
}

// COUPDAYBS関数の実装（直前利払日から受渡日までの日数）
export const COUPDAYBS: CustomFormula = {
  name: 'COUPDAYBS',
  pattern: /COUPDAYBS\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      
      if (basis === 0 || basis === 4) {
        // 30/360方式
        return days360(previousCouponDate, settlement, basis === 4);
      } else {
        // 実際の日数
        return actualDays(previousCouponDate, settlement);
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// COUPDAYS関数の実装（利払期間の日数）
export const COUPDAYS: CustomFormula = {
  name: 'COUPDAYS',
  pattern: /COUPDAYS\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      
      const { daysInYear } = getDayCountBasis(basis);
      
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
  pattern: /COUPDAYSNC\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      
      if (basis === 0 || basis === 4) {
        // 30/360方式
        return days360(settlement, nextCouponDate, basis === 4);
      } else {
        // 実際の日数
        return actualDays(settlement, nextCouponDate);
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// COUPNCD関数の実装（次回利払日）
export const COUPNCD: CustomFormula = {
  name: 'COUPNCD',
  pattern: /COUPNCD\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
  pattern: /COUPNUM\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      let currentDate = new Date(maturity);
      
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
  pattern: /COUPPCD\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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