// 日付関数の実装

import type { CustomFormula, FormulaContext } from './types';
import { FormulaError } from './types';
import { getCellValue } from './utils';
import { 
  parseDate as parseDateNew, 
  diffDays, 
  diffMonths, 
  diffYears, 
  diffMonthsYM, 
  diffDaysYD, 
  diffDaysMD,
  now,
  today,
  dateToExcelSerial,
  isValidDate
} from './dateUtils';

// DATEDIF関数の実装（Date版）
export const DATEDIF: CustomFormula = {
  name: 'DATEDIF',
  pattern: /DATEDIF\(([^,]+),\s*([^,]+),\s*"([^"]+)"\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, startRef, endRef, unit] = matches;
    
    // セル参照か直接値かを判定
    let startValue, endValue;
    
    // 開始日の取得
    if (startRef.startsWith('"') && startRef.endsWith('"')) {
      startValue = startRef.slice(1, -1);
    } else if (startRef.match(/^[A-Z]+\d+$/)) {
      startValue = getCellValue(startRef, context);
    } else {
      startValue = startRef;
    }
    
    // 終了日の取得
    if (endRef.startsWith('"') && endRef.endsWith('"')) {
      endValue = endRef.slice(1, -1);
    } else if (endRef.match(/^[A-Z]+\d+$/)) {
      endValue = getCellValue(endRef, context);
    } else {
      endValue = endRef;
    }
    
    // Date版のparseDate関数を使用
    const startDate = parseDateNew(startValue);
    const endDate = parseDateNew(endValue);
    
    if (!startDate || !endDate) {
      return FormulaError.VALUE;
    }
    
    if (startDate.getTime() > endDate.getTime()) {
      return FormulaError.NUM;
    }
    
    switch (unit.toUpperCase()) {
      case 'D': {
        // 日数の差
        const days = diffDays(startDate, endDate);
        return days;
      }
      case 'M': {
        // 月数の差
        const months = diffMonths(startDate, endDate);
        return months;
      }
      case 'Y': {
        // 年数の差
        const years = diffYears(startDate, endDate);
        return years;
      }
      case 'YM': {
        // 年を無視した月数の差
        const result = diffMonthsYM(startDate, endDate);
        return result;
      }
      case 'YD': {
        // 年を無視した日数の差
        const result = diffDaysYD(startDate, endDate);
        return result;
      }
      case 'MD': {
        // 月を無視した日数の差
        const result = diffDaysMD(startDate, endDate);
        return result;
      }
      default:
        return FormulaError.VALUE;
    }
  }
};

// NETWORKDAYS関数の実装
export const NETWORKDAYS: CustomFormula = {
  name: 'NETWORKDAYS',
  pattern: /NETWORKDAYS\(([^,)]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// NOW関数の実装（現在の日時）
export const NOW: CustomFormula = {
  name: 'NOW',
  pattern: /NOW\(\)/i,
  calculate: () => {
    const currentDate = now();
    return dateToExcelSerial(currentDate) + (currentDate.getHours() * 3600 + currentDate.getMinutes() * 60 + currentDate.getSeconds()) / 86400;
  }
};

// DATE関数の実装（日付作成）
export const DATE: CustomFormula = {
  name: 'DATE',
  pattern: /DATE\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, yearRef, monthRef, dayRef] = matches;
    
    const year = Number(getCellValue(yearRef, context) ?? yearRef);
    const month = Number(getCellValue(monthRef, context) ?? monthRef);
    const day = Number(getCellValue(dayRef, context) ?? dayRef);
    
    if (isNaN(year) || isNaN(month) || isNaN(day)) {
      return FormulaError.VALUE;
    }
    
    const date = new Date(year, month - 1, day); // 月は0ベース
    if (!isValidDate(date)) {
      return FormulaError.VALUE;
    }
    
    return dateToExcelSerial(date);
  }
};

// YEAR関数の実装（年を抽出）
export const YEAR: CustomFormula = {
  name: 'YEAR',
  pattern: /YEAR\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, dateRef] = matches;
    
    const dateValue = getCellValue(dateRef, context) ?? dateRef;
    const date = parseDateNew(dateValue);
    
    if (!date) {
      return FormulaError.VALUE;
    }
    
    return date.getFullYear();
  }
};

// MONTH関数の実装（月を抽出）
export const MONTH: CustomFormula = {
  name: 'MONTH',
  pattern: /MONTH\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, dateRef] = matches;
    
    const dateValue = getCellValue(dateRef, context) ?? dateRef;
    const date = parseDateNew(dateValue);
    
    if (!date) {
      return FormulaError.VALUE;
    }
    
    return date.getMonth() + 1; // 月は1ベース
  }
};

// DAY関数の実装（日を抽出）
export const DAY: CustomFormula = {
  name: 'DAY',
  pattern: /DAY\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, dateRef] = matches;
    
    const dateValue = getCellValue(dateRef, context) ?? dateRef;
    const date = parseDateNew(dateValue);
    
    if (!date) {
      return FormulaError.VALUE;
    }
    
    return date.getDate();
  }
};

// WEEKDAY関数の実装（曜日を数値で返す）
export const WEEKDAY: CustomFormula = {
  name: 'WEEKDAY',
  pattern: /WEEKDAY\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// DAYS関数の実装（日数差）
export const DAYS: CustomFormula = {
  name: 'DAYS',
  pattern: /DAYS\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// EDATE関数の実装（月数後の日付）
export const EDATE: CustomFormula = {
  name: 'EDATE',
  pattern: /EDATE\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// TODAY関数の実装（今日の日付を返す）
export const TODAY: CustomFormula = {
  name: 'TODAY',
  pattern: /TODAY\(\)/i,
  calculate: () => {
    const todayDate = today();
    return dateToExcelSerial(todayDate);
  }
};

// EOMONTH関数の実装（月末日を返す）
export const EOMONTH: CustomFormula = {
  name: 'EOMONTH',
  pattern: /EOMONTH\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// TIME関数の実装（時刻を作成）
export const TIME: CustomFormula = {
  name: 'TIME',
  pattern: /TIME\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// HOUR関数の実装（時を抽出）
export const HOUR: CustomFormula = {
  name: 'HOUR',
  pattern: /HOUR\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// MINUTE関数の実装（分を抽出）
export const MINUTE: CustomFormula = {
  name: 'MINUTE',
  pattern: /MINUTE\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// SECOND関数の実装（秒を抽出）
export const SECOND: CustomFormula = {
  name: 'SECOND',
  pattern: /SECOND\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// WEEKNUM関数の実装（週番号を返す）
export const WEEKNUM: CustomFormula = {
  name: 'WEEKNUM',
  pattern: /WEEKNUM\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// DAYS360関数の実装（360日基準の日数）
export const DAYS360: CustomFormula = {
  name: 'DAYS360',
  pattern: /DAYS360\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// YEARFRAC関数の実装（年の割合を計算）
export const YEARFRAC: CustomFormula = {
  name: 'YEARFRAC',
  pattern: /YEARFRAC\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// DATEVALUE関数の実装（日付文字列を日付値に変換）
export const DATEVALUE: CustomFormula = {
  name: 'DATEVALUE',
  pattern: /DATEVALUE\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// TIMEVALUE関数の実装（時刻文字列を時刻値に変換）
export const TIMEVALUE: CustomFormula = {
  name: 'TIMEVALUE',
  pattern: /TIMEVALUE\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// ISOWEEKNUM関数の実装（ISO週番号を返す）
export const ISOWEEKNUM: CustomFormula = {
  name: 'ISOWEEKNUM',
  pattern: /ISOWEEKNUM\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// NETWORKDAYS.INTL関数の実装（稼働日数・国際版）
export const NETWORKDAYS_INTL: CustomFormula = {
  name: 'NETWORKDAYS.INTL',
  pattern: /NETWORKDAYS\.INTL\(([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// WORKDAY関数の実装（稼働日を計算）
export const WORKDAY: CustomFormula = {
  name: 'WORKDAY',
  pattern: /WORKDAY\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// WORKDAY.INTL関数の実装（稼働日・国際版）
export const WORKDAY_INTL: CustomFormula = {
  name: 'WORKDAY.INTL',
  pattern: /WORKDAY\.INTL\(([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};