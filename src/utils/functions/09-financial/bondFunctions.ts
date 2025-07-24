// 債券関連の財務関数
/* eslint-disable @typescript-eslint/no-unused-vars */

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

// ACCRINT関数の実装（経過利息）
export const ACCRINT: CustomFormula = {
  name: 'ACCRINT',
  pattern: /ACCRINT\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
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
      
      const { daysInYear } = getDayCountBasis(basis);
      let days: number;
      
      if (basis === 0 || basis === 4) {
        // 30/360方式
        days = days360(issue, settlement, basis === 4);
      } else {
        // 実際の日数
        days = actualDays(issue, settlement);
      }
      
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
  pattern: /ACCRINTM\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      
      const { daysInYear } = getDayCountBasis(basis);
      let days: number;
      
      if (basis === 0 || basis === 4) {
        // 30/360方式
        days = days360(issue, maturity, basis === 4);
      } else {
        // 実際の日数
        days = actualDays(issue, maturity);
      }
      
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
  pattern: /DISC\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, priceRef, redemptionRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const price = parseFloat(getCellValue(priceRef.trim(), context)?.toString() ?? priceRef.trim());
      const redemption = parseFloat(getCellValue(redemptionRef.trim(), context)?.toString() ?? redemptionRef.trim());
      const basis = 0; // using default basis value for simplified implementation
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(price) || isNaN(redemption) || price <= 0 || redemption <= 0) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      const { daysInYear } = getDayCountBasis(basis);
      let days: number;
      
      if (0 === 0) {
        // 30/360方式 (simplified implementation)
        days = days360(settlement, maturity, false);
      } else {
        // 実際の日数
        days = actualDays(settlement, maturity);
      }
      
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
  pattern: /DURATION\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      
      // マコーレー・デュレーションの計算（簡易実装）
      const n = Math.ceil(actualDays(settlement, maturity) / (365 / frequency));
      const r = yld / frequency;
      const c = coupon / frequency;
      
      let pv = 0;
      let duration = 0;
      
      for (let t = 1; t <= n; t++) {
        const discountFactor = Math.pow(1 + r, -t);
        const cashFlow = t < n ? c : c + 1;
        const presentValue = cashFlow * discountFactor;
        
        pv += presentValue;
        duration += t * presentValue;
      }
      
      return duration / (pv * frequency);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MDURATION関数の実装（修正デュレーション）
export const MDURATION: CustomFormula = {
  name: 'MDURATION',
  pattern: /MDURATION\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      
      // まずマコーレー・デュレーションを計算
      const n = Math.ceil(actualDays(settlement, maturity) / (365 / frequency));
      const r = yld / frequency;
      const c = coupon / frequency;
      
      let pv = 0;
      let duration = 0;
      
      for (let t = 1; t <= n; t++) {
        const discountFactor = Math.pow(1 + r, -t);
        const cashFlow = t < n ? c : c + 1;
        const presentValue = cashFlow * discountFactor;
        
        pv += presentValue;
        duration += t * presentValue;
      }
      
      const macaulayDuration = duration / (pv * frequency);
      
      // 修正デュレーション = マコーレー・デュレーション / (1 + 利回り/頻度)
      return macaulayDuration / (1 + yld / frequency);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// PRICE関数の実装（証券の価格）
export const PRICE: CustomFormula = {
  name: 'PRICE',
  pattern: /PRICE\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      
      // 簡易的な価格計算
      const n = Math.ceil(actualDays(settlement, maturity) / (365 / frequency));
      const r = yld / frequency;
      const c = rate * redemption / frequency;
      
      let price = 0;
      
      // クーポン支払いの現在価値
      for (let t = 1; t <= n; t++) {
        price += c / Math.pow(1 + r, t);
      }
      
      // 額面の現在価値
      price += redemption / Math.pow(1 + r, n);
      
      return price;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// YIELD関数の実装（証券の利回り）
export const YIELD: CustomFormula = {
  name: 'YIELD',
  pattern: /YIELD\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
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
      
      // Newton-Raphson法で利回りを計算（簡易実装）
      const n = Math.ceil(actualDays(settlement, maturity) / (365 / frequency));
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