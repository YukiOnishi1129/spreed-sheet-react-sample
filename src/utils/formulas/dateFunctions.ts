// 日付関数の実装

import type { CustomFormula, FormulaContext } from './types';
import { FormulaError } from './types';
import { getCellValue, parseDate } from './utils';
import dayjs from 'dayjs';
import isSameOrBefore from 'dayjs/plugin/isSameOrBefore';
import isSameOrAfter from 'dayjs/plugin/isSameOrAfter';

// dayjsプラグインを有効化
dayjs.extend(isSameOrBefore);
dayjs.extend(isSameOrAfter);

// DATEDIF関数の実装
export const DATEDIF: CustomFormula = {
  name: 'DATEDIF',
  pattern: /DATEDIF\(([^,]+),\s*([^,]+),\s*"([^"]+)"\)/i,
  isSupported: false,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, startRef, endRef, unit] = matches;
    
    // セル参照か直接値かを判定
    let startValue, endValue;
    
    // 開始日の取得
    if (startRef.startsWith('"') && startRef.endsWith('"')) {
      // 直接文字列の場合
      startValue = startRef.slice(1, -1);
    } else if (startRef.match(/^[A-Z]+\d+$/)) {
      // セル参照の場合
      startValue = getCellValue(startRef, context);
    } else {
      startValue = startRef;
    }
    
    // 終了日の取得
    if (endRef.startsWith('"') && endRef.endsWith('"')) {
      // 直接文字列の場合
      endValue = endRef.slice(1, -1);
    } else if (endRef.match(/^[A-Z]+\d+$/)) {
      // セル参照の場合
      endValue = getCellValue(endRef, context);
    } else {
      endValue = endRef;
    }
    
    console.log('DATEDIF計算:', { startValue, endValue, unit });
    
    const startDate = parseDate(startValue);
    const endDate = parseDate(endValue);
    
    if (!startDate || !endDate) {
      console.error('DATEDIF: 無効な日付', { startValue, endValue });
      return FormulaError.VALUE;
    }
    
    if (startDate.isAfter(endDate)) {
      console.error('DATEDIF: 開始日が終了日より後', { startDate: startDate.format(), endDate: endDate.format() });
      return FormulaError.NUM;
    }
    
    switch (unit.toUpperCase()) {
      case 'D':
        // 日数の差
        return endDate.diff(startDate, 'day');
      case 'M':
        // 月数の差
        return endDate.diff(startDate, 'month');
      case 'Y':
        // 年数の差
        return endDate.diff(startDate, 'year');
      case 'YM':
        // 年を無視した月数の差
        const monthDiff = endDate.month() - startDate.month();
        return monthDiff < 0 ? monthDiff + 12 : monthDiff;
      case 'YD':
        // 年を無視した日数の差
        const sameYearEnd = endDate.year(startDate.year());
        let dayDiff = sameYearEnd.diff(startDate, 'day');
        if (dayDiff < 0) {
          const nextYearEnd = endDate.year(startDate.year() + 1);
          dayDiff = nextYearEnd.diff(startDate, 'day');
        }
        return dayDiff;
      case 'MD':
        // 月を無視した日数の差
        const dayOfMonthDiff = endDate.date() - startDate.date();
        if (dayOfMonthDiff < 0) {
          const prevMonthDays = endDate.subtract(1, 'month').daysInMonth();
          return prevMonthDays - startDate.date() + endDate.date();
        }
        return dayOfMonthDiff;
      default:
        console.error('DATEDIF: 無効な単位', unit);
        return FormulaError.VALUE;
    }
  }
};

// NETWORKDAYS関数の実装
export const NETWORKDAYS: CustomFormula = {
  name: 'NETWORKDAYS',
  pattern: /NETWORKDAYS\(([^,)]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, startRef, endRef] = matches;
    
    // セル参照か直接値かを判定
    let startValue, endValue;
    
    // 開始日の取得
    if (startRef.match(/^[A-Z]+\d+$/)) {
      startValue = getCellValue(startRef, context);
    } else {
      startValue = startRef.replace(/^"|"$/g, '');
    }
    
    // 終了日の取得
    if (endRef.match(/^[A-Z]+\d+$/)) {
      endValue = getCellValue(endRef, context);
    } else {
      endValue = endRef.replace(/^"|"$/g, '');
    }
    
    console.log('NETWORKDAYS計算:', { startValue, endValue });
    
    const startDate = parseDate(startValue);
    const endDate = parseDate(endValue);
    
    if (!startDate || !endDate) {
      console.error('NETWORKDAYS: 無効な日付', { startValue, endValue });
      return FormulaError.VALUE;
    }
    
    let workdays = 0;
    const isForward = startDate.isSameOrBefore(endDate);
    let currentDate = startDate.clone();
    
    while (isForward ? currentDate.isSameOrBefore(endDate) : currentDate.isSameOrAfter(endDate)) {
      const dayOfWeek = currentDate.day();
      // 月曜日(1)から金曜日(5)までが平日
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        workdays++;
      }
      currentDate = currentDate.add(isForward ? 1 : -1, 'day');
    }
    
    return isForward ? workdays : -workdays;
  }
};

// TODAY関数の実装（HyperFormulaでサポートされているが、日付フォーマットのため）
export const TODAY: CustomFormula = {
  name: 'TODAY',
  pattern: /TODAY\(\)/i,
  isSupported: true,
  calculate: () => {
    // Excelシリアル値を返す（HyperFormulaと同じ）
    const today = dayjs();
    const excelEpoch = dayjs('1899-12-30');
    return today.diff(excelEpoch, 'day');
  }
};