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
    
    console.log('DATEDIF: 計算開始', { startRef, endRef, unit });
    console.log('DATEDIF: コンテキスト', { 
      dataLength: context.data.length, 
      firstRowLength: context.data[0]?.length,
      currentRow: context.row,
      currentCol: context.col
    });
    
    // セル参照か直接値かを判定
    let startValue, endValue;
    
    // 開始日の取得
    if (startRef.startsWith('"') && startRef.endsWith('"')) {
      // 直接文字列の場合
      startValue = startRef.slice(1, -1);
      console.log('DATEDIF: 開始日（直接文字列）', startValue);
    } else if (startRef.match(/^[A-Z]+\d+$/)) {
      // セル参照の場合
      startValue = getCellValue(startRef, context);
      console.log('DATEDIF: 開始日（セル参照）', { cellRef: startRef, value: startValue });
    } else {
      startValue = startRef;
      console.log('DATEDIF: 開始日（その他）', startValue);
    }
    
    // 終了日の取得
    if (endRef.startsWith('"') && endRef.endsWith('"')) {
      // 直接文字列の場合
      endValue = endRef.slice(1, -1);
      console.log('DATEDIF: 終了日（直接文字列）', endValue);
    } else if (endRef.match(/^[A-Z]+\d+$/)) {
      // セル参照の場合
      endValue = getCellValue(endRef, context);
      console.log('DATEDIF: 終了日（セル参照）', { cellRef: endRef, value: endValue });
    } else {
      endValue = endRef;
      console.log('DATEDIF: 終了日（その他）', endValue);
    }
    
    console.log('DATEDIF: 日付解析前の値', { startValue, endValue, startType: typeof startValue, endType: typeof endValue });
    
    const startDate = parseDate(startValue);
    const endDate = parseDate(endValue);
    
    console.log('DATEDIF: 日付解析結果', { 
      startDate: startDate ? startDate.format() : 'null', 
      endDate: endDate ? endDate.format() : 'null' 
    });
    
    if (!startDate || !endDate) {
      console.error('DATEDIF: 日付解析失敗', { startValue, endValue, startDate, endDate });
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
      case 'YM': {
        // 年を無視した月数の差
        const monthDiff = endDate.month() - startDate.month();
        return monthDiff < 0 ? monthDiff + 12 : monthDiff;
      }
      case 'YD': {
        // 年を無視した日数の差
        const sameYearEnd = endDate.year(startDate.year());
        let dayDiff = sameYearEnd.diff(startDate, 'day');
        if (dayDiff < 0) {
          const nextYearEnd = endDate.year(startDate.year() + 1);
          dayDiff = nextYearEnd.diff(startDate, 'day');
        }
        return dayDiff;
      }
      case 'MD': {
        // 月を無視した日数の差
        const dayOfMonthDiff = endDate.date() - startDate.date();
        if (dayOfMonthDiff < 0) {
          const prevMonthDays = endDate.subtract(1, 'month').daysInMonth();
          return prevMonthDays - startDate.date() + endDate.date();
        }
        return dayOfMonthDiff;
      }
      default: {
        console.error('DATEDIF: 無効な単位', unit);
        return FormulaError.VALUE;
      }
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
    
    const startDate = parseDate(startValue);
    const endDate = parseDate(endValue);
    
    if (!startDate || !endDate) {
      console.warn('NETWORKDAYS: 日付解析失敗', { startValue, endValue });
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

// NOW関数の実装（現在の日時）
export const NOW: CustomFormula = {
  name: 'NOW',
  pattern: /NOW\(\)/i,
  isSupported: false,
  calculate: () => {
    // Excelシリアル値を返す
    const now = dayjs();
    const excelEpoch = dayjs('1899-12-30');
    const days = now.diff(excelEpoch, 'day');
    const timeOfDay = (now.hour() * 3600 + now.minute() * 60 + now.second()) / 86400;
    return days + timeOfDay;
  }
};

// DATE関数の実装（日付作成）
export const DATE: CustomFormula = {
  name: 'DATE',
  pattern: /DATE\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const year = parseInt(matches[1]);
    const month = parseInt(matches[2]);
    const day = parseInt(matches[3]);
    
    if (isNaN(year) || isNaN(month) || isNaN(day)) return FormulaError.VALUE;
    if (year < 1900 || year > 9999) return FormulaError.NUM;
    if (month < 1 || month > 12) return FormulaError.NUM;
    if (day < 1 || day > 31) return FormulaError.NUM;
    
    const date = dayjs(`${year}-${month.toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}`);
    if (!date.isValid()) return FormulaError.NUM;
    
    // Excelシリアル値に変換
    const excelEpoch = dayjs('1899-12-30');
    return date.diff(excelEpoch, 'day');
  }
};

// YEAR関数の実装（年を抽出）
export const YEAR: CustomFormula = {
  name: 'YEAR',
  pattern: /YEAR\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const serialDate = parseFloat(matches[1]);
    if (isNaN(serialDate)) return FormulaError.VALUE;
    
    // Excelシリアル値から日付に変換
    const excelEpoch = dayjs('1899-12-30');
    const date = excelEpoch.add(Math.floor(serialDate), 'day');
    
    return date.year();
  }
};

// MONTH関数の実装（月を抽出）
export const MONTH: CustomFormula = {
  name: 'MONTH',
  pattern: /MONTH\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const serialDate = parseFloat(matches[1]);
    if (isNaN(serialDate)) return FormulaError.VALUE;
    
    // Excelシリアル値から日付に変換
    const excelEpoch = dayjs('1899-12-30');
    const date = excelEpoch.add(Math.floor(serialDate), 'day');
    
    return date.month() + 1; // dayjsは0ベース、Excelは1ベース
  }
};

// DAY関数の実装（日を抽出）
export const DAY: CustomFormula = {
  name: 'DAY',
  pattern: /DAY\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const serialDate = parseFloat(matches[1]);
    if (isNaN(serialDate)) return FormulaError.VALUE;
    
    // Excelシリアル値から日付に変換
    const excelEpoch = dayjs('1899-12-30');
    const date = excelEpoch.add(Math.floor(serialDate), 'day');
    
    return date.date();
  }
};

// WEEKDAY関数の実装（曜日を数値で返す）
export const WEEKDAY: CustomFormula = {
  name: 'WEEKDAY',
  pattern: /WEEKDAY\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches) => {
    const serialDate = parseFloat(matches[1]);
    const returnType = matches[2] ? parseInt(matches[2]) : 1;
    
    if (isNaN(serialDate)) return FormulaError.VALUE;
    if (isNaN(returnType) || returnType < 1 || returnType > 3) return FormulaError.NUM;
    
    // Excelシリアル値から日付に変換
    const excelEpoch = dayjs('1899-12-30');
    const date = excelEpoch.add(Math.floor(serialDate), 'day');
    
    const dayOfWeek = date.day(); // 0=日曜日, 1=月曜日, ...
    
    switch (returnType) {
      case 1: // 1=日曜日, 2=月曜日, ..., 7=土曜日
        return dayOfWeek + 1;
      case 2: // 1=月曜日, 2=火曜日, ..., 7=日曜日
        return dayOfWeek === 0 ? 7 : dayOfWeek;
      case 3: // 0=月曜日, 1=火曜日, ..., 6=日曜日
        return dayOfWeek === 0 ? 6 : dayOfWeek - 1;
      default:
        return FormulaError.NUM;
    }
  }
};

// DAYS関数の実装（日数差）
export const DAYS: CustomFormula = {
  name: 'DAYS',
  pattern: /DAYS\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches) => {
    const endDate = parseFloat(matches[1]);
    const startDate = parseFloat(matches[2]);
    
    if (isNaN(endDate) || isNaN(startDate)) return FormulaError.VALUE;
    
    return Math.floor(endDate - startDate);
  }
};

// EDATE関数の実装（月数後の日付）
export const EDATE: CustomFormula = {
  name: 'EDATE',
  pattern: /EDATE\(([^,]+),\s*([^)]+)\)/i,
  isSupported: false,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, dateRef, monthsRef] = matches;
    const months = parseInt(monthsRef.trim());
    
    console.log('EDATE: 計算開始', { dateRef: dateRef.trim(), monthsRef: monthsRef.trim(), months });
    
    if (isNaN(months)) {
      console.error('EDATE: 月数が無効', { monthsRef, months });
      return FormulaError.VALUE;
    }
    
    let startDate;
    const cleanDateRef = dateRef.trim();
    
    // セル参照かチェック
    if (cleanDateRef.match(/^[A-Z]+\d+$/)) {
      console.log('EDATE: セル参照を検出', cleanDateRef);
      const dateValue = getCellValue(cleanDateRef, context);
      console.log('EDATE: セル値取得', { cellRef: cleanDateRef, value: dateValue, type: typeof dateValue });
      startDate = parseDate(dateValue);
    } else {
      // 直接値の場合（文字列の日付や数値）
      const unquotedRef = cleanDateRef.replace(/^["']|["']$/g, ''); // 引用符を除去
      console.log('EDATE: 直接値を処理', { original: cleanDateRef, unquoted: unquotedRef });
      
      const numericDate = parseFloat(unquotedRef);
      
      if (!isNaN(numericDate)) {
        // Excelシリアル値の場合
        console.log('EDATE: Excelシリアル値として処理', numericDate);
        const excelEpoch = dayjs('1899-12-30');
        startDate = excelEpoch.add(Math.floor(numericDate), 'day');
      } else {
        // 文字列日付の場合
        console.log('EDATE: 文字列日付として処理', unquotedRef);
        startDate = parseDate(unquotedRef);
      }
    }
    
    console.log('EDATE: 日付解析結果', { 
      startDate: startDate ? startDate.format() : 'null', 
      isValid: startDate?.isValid() 
    });
    
    if (!startDate?.isValid()) {
      console.error('EDATE: 日付解析失敗', { 
        dateRef: cleanDateRef, 
        startDate, 
        context: { dataLength: context.data.length, row: context.row, col: context.col }
      });
      return FormulaError.VALUE;
    }
    
    const newDate = startDate.add(months, 'month');
    const result = newDate.format('YYYY-MM-DD');
    
    console.log('EDATE: 計算完了', { 
      originalDate: startDate.format(), 
      months, 
      newDate: newDate.format(), 
      result 
    });
    
    return result;
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