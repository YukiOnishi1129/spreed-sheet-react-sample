// 減価償却関連の財務関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';
import { parseDate } from '../shared/dateUtils';

// AMORDEGRC関数の実装（フランス式減価償却）
export const AMORDEGRC: CustomFormula = {
  name: 'AMORDEGRC',
  pattern: /AMORDEGRC\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, costRef, purchaseDateRef, firstPeriodRef, salvageRef, periodRef, rateRef, basisRef] = matches;
    
    try {
      const cost = parseFloat(getCellValue(costRef.trim(), context)?.toString() ?? costRef.trim());
      const purchaseDate = parseDate(getCellValue(purchaseDateRef.trim(), context)?.toString() ?? purchaseDateRef.trim());
      const firstPeriod = parseDate(getCellValue(firstPeriodRef.trim(), context)?.toString() ?? firstPeriodRef.trim());
      const salvage = parseFloat(getCellValue(salvageRef.trim(), context)?.toString() ?? salvageRef.trim());
      const period = parseInt(getCellValue(periodRef.trim(), context)?.toString() ?? periodRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!purchaseDate || !firstPeriod) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(cost) || isNaN(salvage) || isNaN(period) || isNaN(rate)) {
        return FormulaError.NUM;
      }
      
      if (cost < 0 || salvage < 0 || period < 0 || rate <= 0) {
        return FormulaError.NUM;
      }
      
      if (salvage >= cost) {
        return FormulaError.NUM;
      }
      
      // フランス式減価償却係数を取得
      let depreciationCoef: number;
      if (rate >= 0 && rate <= 0.04) {
        depreciationCoef = 1.5;
      } else if (rate > 0.04 && rate <= 0.05) {
        depreciationCoef = 2;
      } else if (rate > 0.05 && rate <= 0.0625) {
        depreciationCoef = 2.5;
      } else {
        depreciationCoef = 1 / (1 - Math.pow(salvage / cost, 1 / Math.ceil(1 / rate)));
      }
      
      // 第1期の日数を計算
      const daysInYear = basis === 0 || basis === 4 ? 360 : 365;
      const firstPeriodDays = Math.floor((firstPeriod.getTime() - purchaseDate.getTime()) / (1000 * 60 * 60 * 24));
      const firstPeriodFraction = firstPeriodDays / daysInYear;
      
      // 累積減価償却額を計算
      let totalDepreciation = 0;
      let remainingValue = cost;
      
      for (let i = 0; i <= period; i++) {
        let periodDepreciation: number;
        
        if (i === 0) {
          // 第1期
          periodDepreciation = cost * rate * depreciationCoef * firstPeriodFraction;
        } else {
          // 第2期以降
          periodDepreciation = remainingValue * rate * depreciationCoef;
        }
        
        // 残存価額を下回らないように調整
        if (remainingValue - periodDepreciation < salvage) {
          periodDepreciation = remainingValue - salvage;
        }
        
        if (i === period) {
          return periodDepreciation;
        }
        
        totalDepreciation += periodDepreciation;
        remainingValue = cost - totalDepreciation;
        
        if (remainingValue <= salvage) {
          return 0;
        }
      }
      
      return 0;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// AMORLINC関数の実装（フランス式定額償却）
export const AMORLINC: CustomFormula = {
  name: 'AMORLINC',
  pattern: /AMORLINC\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, costRef, purchaseDateRef, firstPeriodRef, salvageRef, periodRef, rateRef, basisRef] = matches;
    
    try {
      const cost = parseFloat(getCellValue(costRef.trim(), context)?.toString() ?? costRef.trim());
      const purchaseDate = parseDate(getCellValue(purchaseDateRef.trim(), context)?.toString() ?? purchaseDateRef.trim());
      const firstPeriod = parseDate(getCellValue(firstPeriodRef.trim(), context)?.toString() ?? firstPeriodRef.trim());
      const salvage = parseFloat(getCellValue(salvageRef.trim(), context)?.toString() ?? salvageRef.trim());
      const period = parseInt(getCellValue(periodRef.trim(), context)?.toString() ?? periodRef.trim());
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const basis = basisRef ? parseInt(getCellValue(basisRef.trim(), context)?.toString() ?? basisRef.trim()) : 0;
      
      if (!purchaseDate || !firstPeriod) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(cost) || isNaN(salvage) || isNaN(period) || isNaN(rate)) {
        return FormulaError.NUM;
      }
      
      if (cost < 0 || salvage < 0 || period < 0 || rate <= 0 || rate > 1) {
        return FormulaError.NUM;
      }
      
      if (salvage >= cost) {
        return FormulaError.NUM;
      }
      
      // 定額減価償却額
      const yearlyDepreciation = (cost - salvage) * rate;
      
      if (period === 0) {
        // 第1期の日数を計算
        const daysInYear = basis === 0 || basis === 4 ? 360 : 365;
        const firstPeriodDays = Math.floor((firstPeriod.getTime() - purchaseDate.getTime()) / (1000 * 60 * 60 * 24));
        const firstPeriodFraction = firstPeriodDays / daysInYear;
        
        return yearlyDepreciation * firstPeriodFraction;
      } else {
        // 第2期以降
        const totalLife = 1 / rate;
        
        if (period >= totalLife) {
          return 0;
        }
        
        return yearlyDepreciation;
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CUMIPMT関数の実装（累積利息）
export const CUMIPMT: CustomFormula = {
  name: 'CUMIPMT',
  pattern: /CUMIPMT\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, nperRef, pvRef, startPeriodRef, endPeriodRef, typeRef] = matches;
    
    try {
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const nper = parseInt(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      const pv = parseFloat(getCellValue(pvRef.trim(), context)?.toString() ?? pvRef.trim());
      const startPeriod = parseInt(getCellValue(startPeriodRef.trim(), context)?.toString() ?? startPeriodRef.trim());
      const endPeriod = parseInt(getCellValue(endPeriodRef.trim(), context)?.toString() ?? endPeriodRef.trim());
      const type = parseInt(getCellValue(typeRef.trim(), context)?.toString() ?? typeRef.trim());
      
      if (isNaN(rate) || isNaN(nper) || isNaN(pv) || isNaN(startPeriod) || isNaN(endPeriod) || isNaN(type)) {
        return FormulaError.NUM;
      }
      
      if (rate <= 0 || nper <= 0 || pv <= 0) {
        return FormulaError.NUM;
      }
      
      if (startPeriod < 1 || endPeriod < 1 || startPeriod > endPeriod || endPeriod > nper) {
        return FormulaError.NUM;
      }
      
      if (type !== 0 && type !== 1) {
        return FormulaError.NUM;
      }
      
      // 定期支払額を計算
      const payment = pv * rate * Math.pow(1 + rate, nper) / (Math.pow(1 + rate, nper) - 1);
      
      // 累積利息を計算
      let cumulativeInterest = 0;
      let remainingBalance = pv;
      
      for (let period = 1; period <= endPeriod; period++) {
        let interestPayment: number;
        
        if (type === 0) {
          // 期末払い
          interestPayment = remainingBalance * rate;
        } else {
          // 期首払い
          if (period === 1) {
            interestPayment = 0;
          } else {
            interestPayment = remainingBalance * rate;
          }
        }
        
        const principalPayment = payment - interestPayment;
        
        if (period >= startPeriod) {
          cumulativeInterest += interestPayment;
        }
        
        remainingBalance -= principalPayment;
      }
      
      return -cumulativeInterest;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CUMPRINC関数の実装（累積元本）
export const CUMPRINC: CustomFormula = {
  name: 'CUMPRINC',
  pattern: /CUMPRINC\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rateRef, nperRef, pvRef, startPeriodRef, endPeriodRef, typeRef] = matches;
    
    try {
      const rate = parseFloat(getCellValue(rateRef.trim(), context)?.toString() ?? rateRef.trim());
      const nper = parseInt(getCellValue(nperRef.trim(), context)?.toString() ?? nperRef.trim());
      const pv = parseFloat(getCellValue(pvRef.trim(), context)?.toString() ?? pvRef.trim());
      const startPeriod = parseInt(getCellValue(startPeriodRef.trim(), context)?.toString() ?? startPeriodRef.trim());
      const endPeriod = parseInt(getCellValue(endPeriodRef.trim(), context)?.toString() ?? endPeriodRef.trim());
      const type = parseInt(getCellValue(typeRef.trim(), context)?.toString() ?? typeRef.trim());
      
      if (isNaN(rate) || isNaN(nper) || isNaN(pv) || isNaN(startPeriod) || isNaN(endPeriod) || isNaN(type)) {
        return FormulaError.NUM;
      }
      
      if (rate <= 0 || nper <= 0 || pv <= 0) {
        return FormulaError.NUM;
      }
      
      if (startPeriod < 1 || endPeriod < 1 || startPeriod > endPeriod || endPeriod > nper) {
        return FormulaError.NUM;
      }
      
      if (type !== 0 && type !== 1) {
        return FormulaError.NUM;
      }
      
      // 定期支払額を計算
      const payment = pv * rate * Math.pow(1 + rate, nper) / (Math.pow(1 + rate, nper) - 1);
      
      // 累積元本を計算
      let cumulativePrincipal = 0;
      let remainingBalance = pv;
      
      for (let period = 1; period <= endPeriod; period++) {
        let interestPayment: number;
        
        if (type === 0) {
          // 期末払い
          interestPayment = remainingBalance * rate;
        } else {
          // 期首払い
          if (period === 1) {
            interestPayment = 0;
          } else {
            interestPayment = remainingBalance * rate;
          }
        }
        
        const principalPayment = payment - interestPayment;
        
        if (period >= startPeriod) {
          cumulativePrincipal += principalPayment;
        }
        
        remainingBalance -= principalPayment;
      }
      
      return -cumulativePrincipal;
    } catch {
      return FormulaError.VALUE;
    }
  }
};