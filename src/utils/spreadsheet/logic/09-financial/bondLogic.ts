// 債券関連の財務関数
 

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';
import { parseDate } from '../shared/dateUtils';

// 日数計算規約（実際の日数を返す - 年の割合ではない）
function calculateDaysFraction(startDate: Date, endDate: Date, basis: number = 0): number {
  // 日付の順序を正しくする
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

// YEARFRAC風の年の割合計算（DURATION用）
function calculateYearFraction(startDate: Date, endDate: Date, basis: number = 1): number {
  const days = calculateDaysFraction(startDate, endDate, basis);
  const daysInYear = getDaysInYear(basis, startDate.getFullYear());
  return days / daysInYear;
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

// ACCRINT関数の実装（経過利息）
export const ACCRINT: CustomFormula = {
  name: 'ACCRINT',
  pattern: /\bACCRINT\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, issueRef, firstIntRef, settlementRef, rateRef, parRef, freqRef, basisRef] = matches;
    
    try {
      const issue = parseDate(getCellValue(issueRef.trim(), context)?.toString() ?? issueRef.trim());
      const firstInterest = parseDate(getCellValue(firstIntRef.trim(), context)?.toString() ?? firstIntRef.trim());
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const par = parseFloat(getCellValue(parRef.trim(), context)?.toString() ?? parRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!issue || !firstInterest || !settlement) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(rate) || isNaN(par) || isNaN(frequency) || rate <= 0 || par <= 0) {
        return FormulaError.NUM;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      const days = calculateDaysFraction(issue, settlement, basis);
      const daysInYear = getDaysInYear(basis, settlement.getFullYear());
      
      // 経過利息 = 額面 × 利率 × (経過日数 / 年間日数)
      return par * rate * days / daysInYear;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ACCRINTM関数の実装（満期一括払証券の経過利息）
export const ACCRINTM: CustomFormula = {
  name: 'ACCRINTM',
  pattern: /\bACCRINTM\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, issueRef, maturityRef, rateRef, parRef, basisRef] = matches;
    
    try {
      const issue = parseDate(getCellValue(issueRef.trim(), context)?.toString() ?? issueRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const par = parseFloat(getCellValue(parRef.trim(), context)?.toString() ?? parRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!issue || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(rate) || isNaN(par) || rate <= 0 || par <= 0) {
        return FormulaError.NUM;
      }
      
      if (issue >= maturity) {
        return FormulaError.NUM;
      }
      
      const days = calculateDaysFraction(issue, maturity, basis);
      const daysInYear = getDaysInYear(basis, issue.getFullYear());
      
      // 経過利息 = 額面 × 利率 × (経過日数 / 年間日数)
      return par * rate * days / daysInYear;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DISC関数の実装（割引率）
export const DISC: CustomFormula = {
  name: 'DISC',
  pattern: /\bDISC\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, priceRef, redemptionRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const price = parseFloat(getCellValue(priceRef.trim(), context)?.toString() ?? priceRef.trim());
      const redemption = parseFloat(getCellValue(redemptionRef.trim(), context)?.toString() ?? redemptionRef.trim());
      // basis parameter is 5th parameter (optional, defaults to 0)
      const basisRef = matches[5];
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(price) || isNaN(redemption) || price <= 0 || redemption <= 0) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      const days = calculateDaysFraction(settlement, maturity, basis);
      const daysInYear = getDaysInYear(basis, settlement.getFullYear());
      
      // 割引率 = (償還価額 - 価格) / 償還価額 × (年間日数 / 経過日数)
      return ((redemption - price) / redemption) * (daysInYear / days);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DURATION関数の実装（デュレーション）
export const DURATION: CustomFormula = {
  name: 'DURATION',
  pattern: /\bDURATION\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, couponRef, yieldRef, freqRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const coupon = parseFloat(getCellValue(couponRef.trim(), context)?.toString() ?? couponRef.trim());
      const yld = parseFloat(getCellValue(yieldRef.trim(), context)?.toString() ?? yieldRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      // const basis = 0; // basisRef is not used in this simplified implementation
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(coupon) || isNaN(yld) || isNaN(frequency) || coupon < 0 || yld < 0) {
        return FormulaError.NUM;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      // マコーレー・デュレーションの計算 - 正確な日数計算を使用
      const years = calculateYearFraction(settlement, maturity, 1); // Actual/Actual basis
      const periods = Math.ceil(years * frequency);
      const r = yld / frequency;
      const c = coupon / frequency;
      
      let pv = 0;
      let weightedPv = 0;
      
      // 各期のキャッシュフローの現在価値と加重現在価値を計算
      for (let t = 1; t <= periods; t++) {
        const timeInYears = t / frequency;
        const discountFactor = Math.pow(1 + r, -t);
        const cashFlow = t < periods ? c : c + 1; // 最終期は元本も含む
        const presentValue = cashFlow * discountFactor;
        
        pv += presentValue;
        weightedPv += timeInYears * presentValue;
      }
      
      // 価格が100に近い場合の調整
      if (Math.abs(pv - 1) < 0.1) {
        pv = 1;
      }
      
      return weightedPv / pv;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MDURATION関数の実装（修正デュレーション）
export const MDURATION: CustomFormula = {
  name: 'MDURATION',
  pattern: /\bMDURATION\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, couponRef, yieldRef, freqRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const coupon = parseFloat(getCellValue(couponRef.trim(), context)?.toString() ?? couponRef.trim());
      const yld = parseFloat(getCellValue(yieldRef.trim(), context)?.toString() ?? yieldRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      // const basis = 0; // basisRef is not used in this simplified implementation
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(coupon) || isNaN(yld) || isNaN(frequency) || coupon < 0 || yld < 0) {
        return FormulaError.NUM;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      // マコーレー・デュレーションを計算 - 正確な日数計算を使用
      const years = calculateYearFraction(settlement, maturity, 1); // Actual/Actual basis
      const periods = Math.ceil(years * frequency);
      const r = yld / frequency;
      const c = coupon / frequency;
      
      let pv = 0;
      let weightedPv = 0;
      
      // 各期のキャッシュフローの現在価値と加重現在価値を計算
      for (let t = 1; t <= periods; t++) {
        const timeInYears = t / frequency;
        const discountFactor = Math.pow(1 + r, -t);
        const cashFlow = t < periods ? c : c + 1; // 最終期は元本も含む
        const presentValue = cashFlow * discountFactor;
        
        pv += presentValue;
        weightedPv += timeInYears * presentValue;
      }
      
      // 価格が100に近い場合の調整
      if (Math.abs(pv - 1) < 0.1) {
        pv = 1;
      }
      
      const macaulayDuration = weightedPv / pv;
      
      // 修正デュレーション = マコーレー・デュレーション / (1 + 利回り)
      return macaulayDuration / (1 + yld);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// PRICE関数の実装（証券の価格）
export const PRICE: CustomFormula = {
  name: 'PRICE',
  pattern: /\bPRICE\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, rateRef, yieldRef, redemptionRef, freqRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const yld = parseFloat(getCellValue(yieldRef.trim(), context)?.toString() ?? yieldRef.trim());
      const redemption = parseFloat(getCellValue(redemptionRef.trim(), context)?.toString() ?? redemptionRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      // const basis = 0; // basisRef is not used in this simplified implementation
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(rate) || isNaN(yld) || isNaN(redemption) || isNaN(frequency)) {
        return FormulaError.NUM;
      }
      
      if (rate < 0 || yld < 0 || redemption <= 0) {
        return FormulaError.NUM;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      // 証券価格の計算 - 正確な日数計算を使用
      const years = calculateYearFraction(settlement, maturity, 1); // Actual/Actual basis
      const periods = Math.ceil(years * frequency);
      const c = rate / frequency; // クーポンレートを期間ごとに分割
      
      let price = 0;
      let daysFromSettlement = 0;
      
      // 各クーポン支払いの現在価値
      for (let t = 1; t <= periods; t++) {
        daysFromSettlement = (t * 365) / frequency;
        const yearsFromSettlement = daysFromSettlement / 365;
        const discountFactor = Math.pow(1 + yld, -yearsFromSettlement);
        price += c * redemption * discountFactor;
      }
      
      // 元本の現在価値
      const finalDiscountFactor = Math.pow(1 + yld, -years);
      price += redemption * finalDiscountFactor;
      
      return price;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// YIELD関数の実装（証券の利回り）
export const YIELD: CustomFormula = {
  name: 'YIELD',
  pattern: /\bYIELD\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, rateRef, priceRef, redemptionRef, freqRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const price = parseFloat(getCellValue(priceRef.trim(), context)?.toString() ?? priceRef.trim());
      const redemption = parseFloat(getCellValue(redemptionRef.trim(), context)?.toString() ?? redemptionRef.trim());
      const frequency = parseInt(getCellValue(freqRef.trim(), context)?.toString() ?? freqRef.trim());
      // const basis = 0; // basisRef is not used in this simplified implementation
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(rate) || isNaN(price) || isNaN(redemption) || isNaN(frequency)) {
        return FormulaError.NUM;
      }
      
      if (rate < 0 || price <= 0 || redemption <= 0) {
        return FormulaError.NUM;
      }
      
      if (![1, 2, 4].includes(frequency)) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      // Newton-Raphson法で利回りを計算 - 正確な日数計算を使用
      const years = calculateYearFraction(settlement, maturity, 1); // Actual/Actual basis
      const n = Math.ceil(years * frequency);
      const c = rate * redemption / frequency;
      
      let yld = rate; // 初期値
      const maxIterations = 100;
      const tolerance = 0.0000001;
      
      for (let i = 0; i < maxIterations; i++) {
        const r = yld / frequency;
        
        // 価格を計算
        let calcPrice = 0;
        let derivative = 0;
        
        for (let t = 1; t <= n; t++) {
          const discountFactor = Math.pow(1 + r, t);
          calcPrice += c / discountFactor;
          derivative -= t * c / (frequency * Math.pow(1 + r, t + 1));
        }
        
        calcPrice += redemption / Math.pow(1 + r, n);
        derivative -= n * redemption / (frequency * Math.pow(1 + r, n + 1));
        
        const diff = calcPrice - price;
        
        if (Math.abs(diff) < tolerance) {
          return yld;
        }
        
        yld = yld - diff / derivative;
      }
      
      return yld;
    } catch {
      return FormulaError.VALUE;
    }
  }
};