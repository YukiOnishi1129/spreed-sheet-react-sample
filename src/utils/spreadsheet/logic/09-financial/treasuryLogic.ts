// 国債関連の財務関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';
import { parseDate } from '../shared/dateUtils';

// 実際の日数計算（Excel互換）
function actualDays(startDate: Date, endDate: Date): number {
  return Math.floor((endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000));
}

// TBILLEQ関数の実装（米国短期国債の債券換算利回り）
export const TBILLEQ: CustomFormula = {
  name: 'TBILLEQ',
  pattern: /TBILLEQ\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, discountRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const discount = parseFloat(getCellValue(discountRef.trim(), context)?.toString() ?? discountRef.trim());
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(discount) || discount < 0 || discount >= 1) {
        return FormulaError.NUM;
      }
      
      const days = actualDays(settlement, maturity);
      
      if (days <= 0 || days > 365) {
        return FormulaError.NUM;
      }
      
      // 債券換算利回り = (365 × 割引率) / (360 - 割引率 × 満期までの日数)
      return (365 * discount) / (360 - discount * days);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TBILLPRICE関数の実装（米国短期国債の価格）
export const TBILLPRICE: CustomFormula = {
  name: 'TBILLPRICE',
  pattern: /TBILLPRICE\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, discountRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const discount = parseFloat(getCellValue(discountRef.trim(), context)?.toString() ?? discountRef.trim());
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(discount) || discount < 0) {
        return FormulaError.NUM;
      }
      
      const days = actualDays(settlement, maturity);
      
      if (days <= 0 || days > 365) {
        return FormulaError.NUM;
      }
      
      // 価格 = 100 × (1 - 割引率 × 満期までの日数 / 360)
      return 100 * (1 - discount * days / 360);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TBILLYIELD関数の実装（米国短期国債の利回り）
export const TBILLYIELD: CustomFormula = {
  name: 'TBILLYIELD',
  pattern: /TBILLYIELD\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, priceRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const price = parseFloat(getCellValue(priceRef.trim(), context)?.toString() ?? priceRef.trim());
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(price) || price <= 0 || price > 100) {
        return FormulaError.NUM;
      }
      
      const days = actualDays(settlement, maturity);
      
      if (days <= 0 || days > 365) {
        return FormulaError.NUM;
      }
      
      // 利回り = ((100 - 価格) / 価格) × (360 / 満期までの日数)
      return ((100 - price) / price) * (360 / days);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DOLLARDE関数の実装（分数表記のドル価格を小数表記に変換）
export const DOLLARDE: CustomFormula = {
  name: 'DOLLARDE',
  pattern: /DOLLARDE\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, fractionalDollarRef, fractionRef] = matches;
    
    try {
      const fractionalDollar = parseFloat(getCellValue(fractionalDollarRef.trim(), context)?.toString() ?? fractionalDollarRef.trim());
      const fraction = parseInt(getCellValue(fractionRef.trim(), context)?.toString() ?? fractionRef.trim());
      
      if (isNaN(fractionalDollar) || isNaN(fraction)) {
        return FormulaError.VALUE;
      }
      
      if (fraction < 0) {
        return FormulaError.NUM;
      }
      
      if (fraction === 0) {
        return FormulaError.DIV0;
      }
      
      // 整数部分と小数部分を分離
      const integerPart = Math.floor(Math.abs(fractionalDollar));
      const fractionalPart = Math.abs(fractionalDollar) - integerPart;
      
      // 分数表記を小数に変換
      const decimalPart = fractionalPart * 100 / fraction;
      
      const result = integerPart + decimalPart;
      
      return fractionalDollar < 0 ? -result : result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DOLLARFR関数の実装（小数表記のドル価格を分数表記に変換）
export const DOLLARFR: CustomFormula = {
  name: 'DOLLARFR',
  pattern: /DOLLARFR\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, decimalDollarRef, fractionRef] = matches;
    
    try {
      const decimalDollar = parseFloat(getCellValue(decimalDollarRef.trim(), context)?.toString() ?? decimalDollarRef.trim());
      const fraction = parseInt(getCellValue(fractionRef.trim(), context)?.toString() ?? fractionRef.trim());
      
      if (isNaN(decimalDollar) || isNaN(fraction)) {
        return FormulaError.VALUE;
      }
      
      if (fraction < 0) {
        return FormulaError.NUM;
      }
      
      if (fraction === 0) {
        return FormulaError.DIV0;
      }
      
      // 整数部分と小数部分を分離
      const integerPart = Math.floor(Math.abs(decimalDollar));
      const decimalPart = Math.abs(decimalDollar) - integerPart;
      
      // 小数を分数表記に変換
      const fractionalPart = decimalPart * fraction / 100;
      
      const result = integerPart + fractionalPart;
      
      return decimalDollar < 0 ? -result : result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// EFFECT関数の実装（実効年利率）
export const EFFECT: CustomFormula = {
  name: 'EFFECT',
  pattern: /EFFECT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, nominalRateRef, nperRef] = matches;
    
    try {
      const nominalRate = parseFloat(getCellValue(nominalRateRef.trim(), context)?.toString() ?? nominalRateRef.trim());
      const npery = parseInt(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      
      if (isNaN(nominalRate) || isNaN(npery)) {
        return FormulaError.VALUE;
      }
      
      if (nominalRate <= 0 || npery < 1) {
        return FormulaError.NUM;
      }
      
      // 実効年利率 = (1 + 名目利率/期間数)^期間数 - 1
      return Math.pow(1 + nominalRate / npery, npery) - 1;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// NOMINAL関数の実装（名目年利率）
export const NOMINAL: CustomFormula = {
  name: 'NOMINAL',
  pattern: /NOMINAL\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, effectRateRef, nperRef] = matches;
    
    try {
      const effectRate = parseFloat(getCellValue(effectRateRef.trim(), context)?.toString() ?? effectRateRef.trim());
      const npery = parseInt(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      
      if (isNaN(effectRate) || isNaN(npery)) {
        return FormulaError.VALUE;
      }
      
      if (effectRate <= 0 || npery < 1) {
        return FormulaError.NUM;
      }
      
      // 名目利率 = 期間数 × ((1 + 実効利率)^(1/期間数) - 1)
      return npery * (Math.pow(1 + effectRate, 1 / npery) - 1);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// PRICEDISC関数の実装（割引証券の価格）
export const PRICEDISC: CustomFormula = {
  name: 'PRICEDISC',
  pattern: /PRICEDISC\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, discountRef, redemptionRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const discount = parseFloat(getCellValue(discountRef.trim(), context)?.toString() ?? discountRef.trim());
      const redemption = parseFloat(getCellValue(redemptionRef.trim(), context)?.toString() ?? redemptionRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(discount) || isNaN(redemption)) {
        return FormulaError.NUM;
      }
      
      if (discount <= 0 || redemption <= 0) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      const days = actualDays(settlement, maturity);
      const daysInYear = basis === 0 || basis === 4 ? 360 : 365;
      
      // 価格 = 償還価額 - 割引率 × 償還価額 × (満期までの日数 / 年間日数)
      return redemption - discount * redemption * days / daysInYear;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// RECEIVED関数の実装（満期受取額）
export const RECEIVED: CustomFormula = {
  name: 'RECEIVED',
  pattern: /RECEIVED\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, investmentRef, discountRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const investment = parseFloat(getCellValue(investmentRef.trim(), context)?.toString() ?? investmentRef.trim());
      const discount = parseFloat(getCellValue(discountRef.trim(), context)?.toString() ?? discountRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(investment) || isNaN(discount)) {
        return FormulaError.NUM;
      }
      
      if (investment <= 0 || discount <= 0) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      const days = actualDays(settlement, maturity);
      const daysInYear = basis === 0 || basis === 4 ? 360 : 365;
      
      // 受取額 = 投資額 / (1 - 割引率 × 満期までの日数 / 年間日数)
      return investment / (1 - discount * days / daysInYear);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// INTRATE関数の実装（満期一括証券の利率）
export const INTRATE: CustomFormula = {
  name: 'INTRATE',
  pattern: /INTRATE\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, investmentRef, redemptionRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const investment = parseFloat(getCellValue(investmentRef.trim(), context)?.toString() ?? investmentRef.trim());
      const redemption = parseFloat(getCellValue(redemptionRef.trim(), context)?.toString() ?? redemptionRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(investment) || isNaN(redemption) || investment <= 0 || redemption <= 0) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity) {
        return FormulaError.NUM;
      }
      
      const days = actualDays(settlement, maturity);
      let daysInYear: number;
      
      switch (basis) {
        case 0: // 30/360 US
        case 4: // 30/360 European
          daysInYear = 360;
          break;
        case 1: // Actual/actual
        case 3: // Actual/365
          daysInYear = 365;
          break;
        case 2: // Actual/360
          daysInYear = 360;
          break;
        default:
          daysInYear = 360;
      }
      
      // 利率 = ((償還額 - 投資額) / 投資額) × (年間日数 / 期間日数)
      return ((redemption - investment) / investment) * (daysInYear / days);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// PRICEMAT関数の実装（満期利付証券の価格）
export const PRICEMAT: CustomFormula = {
  name: 'PRICEMAT',
  pattern: /PRICEMAT\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, issueRef, rateRef, yieldRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const issue = parseDate(getCellValue(issueRef.trim(), context)?.toString() ?? issueRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const yld = parseFloat(getCellValue(yieldRef.trim(), context)?.toString() ?? yieldRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity || !issue) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(rate) || isNaN(yld) || rate < 0 || yld < 0) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity || issue >= settlement) {
        return FormulaError.NUM;
      }
      
      const daysIM = actualDays(issue, maturity);
      const daysSM = actualDays(settlement, maturity);
      const daysIS = actualDays(issue, settlement);
      
      let daysInYear: number;
      switch (basis) {
        case 0: // 30/360 US
        case 4: // 30/360 European
          daysInYear = 360;
          break;
        case 1: // Actual/actual
        case 3: // Actual/365
          daysInYear = 365;
          break;
        case 2: // Actual/360
          daysInYear = 360;
          break;
        default:
          daysInYear = 360;
      }
      
      // 価格計算
      const A = daysIS / daysInYear;
      const B = daysSM / daysInYear;
      const price = (100 + (rate * 100 * daysIM / daysInYear)) / (1 + (yld * B)) - (rate * 100 * A);
      
      return price;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// YIELDDISC関数の実装（割引証券の利回り）
export const YIELDDISC: CustomFormula = {
  name: 'YIELDDISC',
  pattern: /YIELDDISC\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, priceRef, redemptionRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const price = parseFloat(getCellValue(priceRef.trim(), context)?.toString() ?? priceRef.trim());
      const redemption = parseFloat(getCellValue(redemptionRef.trim(), context)?.toString() ?? redemptionRef.trim());
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
      
      const days = actualDays(settlement, maturity);
      let daysInYear: number;
      
      switch (basis) {
        case 0: // 30/360 US
        case 4: // 30/360 European
          daysInYear = 360;
          break;
        case 1: // Actual/actual
        case 3: // Actual/365
          daysInYear = 365;
          break;
        case 2: // Actual/360
          daysInYear = 360;
          break;
        default:
          daysInYear = 360;
      }
      
      // 利回り = ((償還額 - 価格) / 価格) × (年間日数 / 期間日数)
      return ((redemption - price) / price) * (daysInYear / days);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// YIELDMAT関数の実装（満期利付証券の利回り）
export const YIELDMAT: CustomFormula = {
  name: 'YIELDMAT',
  pattern: /YIELDMAT\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, settlementRef, maturityRef, issueRef, rateRef, priceRef, basisRef] = matches;
    
    try {
      const settlement = parseDate(getCellValue(settlementRef.trim(), context)?.toString() ?? settlementRef.trim());
      const maturity = parseDate(getCellValue(maturityRef.trim(), context)?.toString() ?? maturityRef.trim());
      const issue = parseDate(getCellValue(issueRef.trim(), context)?.toString() ?? issueRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const price = parseFloat(getCellValue(priceRef.trim(), context)?.toString() ?? priceRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!settlement || !maturity || !issue) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(rate) || isNaN(price) || rate < 0 || price <= 0) {
        return FormulaError.NUM;
      }
      
      if (settlement >= maturity || issue >= settlement) {
        return FormulaError.NUM;
      }
      
      const daysIM = actualDays(issue, maturity);
      const daysSM = actualDays(settlement, maturity);
      const daysIS = actualDays(issue, settlement);
      
      let daysInYear: number;
      switch (basis) {
        case 0: // 30/360 US
        case 4: // 30/360 European
          daysInYear = 360;
          break;
        case 1: // Actual/actual
        case 3: // Actual/365
          daysInYear = 365;
          break;
        case 2: // Actual/360
          daysInYear = 360;
          break;
        default:
          daysInYear = 360;
      }
      
      // 利回り計算（簡易実装）
      const A = daysIS / daysInYear;
      const B = daysSM / daysInYear;
      const yld = ((100 + (rate * 100 * daysIM / daysInYear)) / (price + (rate * 100 * A)) - 1) / B;
      
      return yld;
    } catch {
      return FormulaError.VALUE;
    }
  }
};