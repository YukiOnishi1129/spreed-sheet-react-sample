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
  calculate: (matches, context) => {
    const dateRef = matches[1].trim();
    let dateValue;
    
    // セル参照かチェック
    if (dateRef.match(/^[A-Z]+\d+$/)) {
      dateValue = getCellValue(dateRef, context);
    } else {
      dateValue = dateRef.replace(/"/g, '');
    }
    
    // 日付を解析
    const date = parseDate(dateValue);
    if (!date?.isValid()) {
      console.error('YEAR: 日付解析失敗', { dateRef, dateValue });
      return FormulaError.VALUE;
    }
    
    return date.year();
  }
};

// MONTH関数の実装（月を抽出）
export const MONTH: CustomFormula = {
  name: 'MONTH',
  pattern: /MONTH\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const dateRef = matches[1].trim();
    let dateValue;
    
    // セル参照かチェック
    if (dateRef.match(/^[A-Z]+\d+$/)) {
      dateValue = getCellValue(dateRef, context);
    } else {
      dateValue = dateRef.replace(/"/g, '');
    }
    
    // 日付を解析
    const date = parseDate(dateValue);
    if (!date?.isValid()) {
      console.error('MONTH: 日付解析失敗', { dateRef, dateValue });
      return FormulaError.VALUE;
    }
    
    return date.month() + 1; // dayjsは0ベース、Excelは1ベース
  }
};

// DAY関数の実装（日を抽出）
export const DAY: CustomFormula = {
  name: 'DAY',
  pattern: /DAY\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const dateRef = matches[1].trim();
    let dateValue;
    
    // セル参照かチェック
    if (dateRef.match(/^[A-Z]+\d+$/)) {
      dateValue = getCellValue(dateRef, context);
    } else {
      dateValue = dateRef.replace(/"/g, '');
    }
    
    // 日付を解析
    const date = parseDate(dateValue);
    if (!date?.isValid()) {
      console.error('DAY: 日付解析失敗', { dateRef, dateValue });
      return FormulaError.VALUE;
    }
    
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

// EOMONTH関数の実装（月末日を返す）
export const EOMONTH: CustomFormula = {
  name: 'EOMONTH',
  pattern: /EOMONTH\(([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// TIME関数の実装（時刻を作成）
export const TIME: CustomFormula = {
  name: 'TIME',
  pattern: /TIME\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// HOUR関数の実装（時を抽出）
export const HOUR: CustomFormula = {
  name: 'HOUR',
  pattern: /HOUR\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const timeRef = matches[1].trim();
    let timeValue;
    
    if (timeRef.match(/^[A-Z]+\d+$/)) {
      timeValue = getCellValue(timeRef, context);
    } else if (timeRef.startsWith('"') && timeRef.endsWith('"')) {
      timeValue = timeRef.slice(1, -1);
    } else {
      timeValue = timeRef;
    }
    
    // Excelシリアル値の場合（数値）
    if (typeof timeValue === 'number' || !isNaN(parseFloat(timeValue))) {
      const num = typeof timeValue === 'number' ? timeValue : parseFloat(timeValue);
      const hours = Math.floor((num % 1) * 24);
      return hours;
    }
    
    // 時刻文字列の場合
    const timeStr = String(timeValue);
    const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (timeMatch) {
      return parseInt(timeMatch[1]);
    }
    
    return FormulaError.VALUE;
  }
};

// MINUTE関数の実装（分を抽出）
export const MINUTE: CustomFormula = {
  name: 'MINUTE',
  pattern: /MINUTE\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const timeRef = matches[1].trim();
    let timeValue;
    
    if (timeRef.match(/^[A-Z]+\d+$/)) {
      timeValue = getCellValue(timeRef, context);
    } else if (timeRef.startsWith('"') && timeRef.endsWith('"')) {
      timeValue = timeRef.slice(1, -1);
    } else {
      timeValue = timeRef;
    }
    
    // Excelシリアル値の場合（数値）
    if (typeof timeValue === 'number' || !isNaN(parseFloat(timeValue))) {
      const num = typeof timeValue === 'number' ? timeValue : parseFloat(timeValue);
      const totalMinutes = Math.floor((num % 1) * 24 * 60);
      const minutes = totalMinutes % 60;
      return minutes;
    }
    
    // 時刻文字列の場合
    const timeStr = String(timeValue);
    const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (timeMatch) {
      return parseInt(timeMatch[2]);
    }
    
    return FormulaError.VALUE;
  }
};

// SECOND関数の実装（秒を抽出）
export const SECOND: CustomFormula = {
  name: 'SECOND',
  pattern: /SECOND\(([^)]+)\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const timeRef = matches[1].trim();
    let timeValue;
    
    if (timeRef.match(/^[A-Z]+\d+$/)) {
      timeValue = getCellValue(timeRef, context);
    } else if (timeRef.startsWith('"') && timeRef.endsWith('"')) {
      timeValue = timeRef.slice(1, -1);
    } else {
      timeValue = timeRef;
    }
    
    // Excelシリアル値の場合（数値）
    if (typeof timeValue === 'number' || !isNaN(parseFloat(timeValue))) {
      const num = typeof timeValue === 'number' ? timeValue : parseFloat(timeValue);
      const totalSeconds = Math.floor((num % 1) * 24 * 60 * 60);
      const seconds = totalSeconds % 60;
      return seconds;
    }
    
    // 時刻文字列の場合
    const timeStr = String(timeValue);
    const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (timeMatch) {
      return parseInt(timeMatch[3] || '0');
    }
    
    return FormulaError.VALUE;
  }
};

// WEEKNUM関数の実装（週番号を返す）
export const WEEKNUM: CustomFormula = {
  name: 'WEEKNUM',
  pattern: /WEEKNUM\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const dateRef = matches[1].trim();
    const returnTypeRef = matches[2]?.trim() || '1';
    
    let dateValue;
    if (dateRef.match(/^[A-Z]+\d+$/)) {
      dateValue = getCellValue(dateRef, context);
    } else if (dateRef.startsWith('"') && dateRef.endsWith('"')) {
      dateValue = dateRef.slice(1, -1);
    } else {
      dateValue = dateRef;
    }
    
    let returnType: number;
    if (returnTypeRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(returnTypeRef, context);
      returnType = parseInt(String(cellValue ?? '1'));
    } else {
      returnType = parseInt(returnTypeRef);
    }
    
    const date = parseDate(dateValue);
    if (!date?.isValid()) return FormulaError.VALUE;
    
    // 年の最初の日
    const yearStart = date.startOf('year');
    
    // 週の開始日を設定（1=日曜日、2=月曜日）
    const startOfWeek = returnType === 2 ? 1 : 0; // 0=日曜日、1=月曜日
    
    // 年の最初の週の開始日
    const firstWeekStart = yearStart.day(startOfWeek);
    
    // 日付が年始より前の場合は前年の最終週
    if (date.isBefore(firstWeekStart)) {
      return 1;
    }
    
    // 週番号を計算
    const weekNum = Math.floor(date.diff(firstWeekStart, 'day') / 7) + 1;
    return weekNum;
  }
};

// DAYS360関数の実装（360日基準の日数）
export const DAYS360: CustomFormula = {
  name: 'DAYS360',
  pattern: /DAYS360\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const startDateRef = matches[1].trim();
    const endDateRef = matches[2].trim();
    const methodRef = matches[3]?.trim() || 'FALSE';
    
    let startDateValue, endDateValue;
    
    // 開始日を取得
    if (startDateRef.match(/^[A-Z]+\d+$/)) {
      startDateValue = getCellValue(startDateRef, context);
    } else if (startDateRef.startsWith('"') && startDateRef.endsWith('"')) {
      startDateValue = startDateRef.slice(1, -1);
    } else {
      startDateValue = startDateRef;
    }
    
    // 終了日を取得
    if (endDateRef.match(/^[A-Z]+\d+$/)) {
      endDateValue = getCellValue(endDateRef, context);
    } else if (endDateRef.startsWith('"') && endDateRef.endsWith('"')) {
      endDateValue = endDateRef.slice(1, -1);
    } else {
      endDateValue = endDateRef;
    }
    
    // メソッドを取得（US=FALSE, European=TRUE）
    let method: boolean;
    if (methodRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(methodRef, context);
      method = Boolean(cellValue);
    } else {
      method = methodRef.toUpperCase() === 'TRUE';
    }
    
    const startDate = parseDate(startDateValue);
    const endDate = parseDate(endDateValue);
    
    if (!startDate?.isValid() || !endDate?.isValid()) {
      return FormulaError.VALUE;
    }
    
    let startY = startDate.year();
    let startM = startDate.month() + 1; // dayjsは0ベース
    let startD = startDate.date();
    
    let endY = endDate.year();
    let endM = endDate.month() + 1;
    let endD = endDate.date();
    
    // 360日基準での日数計算
    if (method) {
      // European method
      if (startD === 31) startD = 30;
      if (endD === 31) endD = 30;
    } else {
      // US method
      if (startD === 31) startD = 30;
      if (endD === 31 && startD >= 30) endD = 30;
    }
    
    return (endY - startY) * 360 + (endM - startM) * 30 + (endD - startD);
  }
};

// YEARFRAC関数の実装（年の割合を計算）
export const YEARFRAC: CustomFormula = {
  name: 'YEARFRAC',
  pattern: /YEARFRAC\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: false,
  calculate: (matches, context) => {
    const startDateRef = matches[1].trim();
    const endDateRef = matches[2].trim();
    const basisRef = matches[3]?.trim() || '0';
    
    let startDateValue, endDateValue;
    
    // 開始日を取得
    if (startDateRef.match(/^[A-Z]+\d+$/)) {
      startDateValue = getCellValue(startDateRef, context);
    } else if (startDateRef.startsWith('"') && startDateRef.endsWith('"')) {
      startDateValue = startDateRef.slice(1, -1);
    } else {
      startDateValue = startDateRef;
    }
    
    // 終了日を取得
    if (endDateRef.match(/^[A-Z]+\d+$/)) {
      endDateValue = getCellValue(endDateRef, context);
    } else if (endDateRef.startsWith('"') && endDateRef.endsWith('"')) {
      endDateValue = endDateRef.slice(1, -1);
    } else {
      endDateValue = endDateRef;
    }
    
    // 基準を取得
    let basis: number;
    if (basisRef.match(/^[A-Z]+\d+$/)) {
      const cellValue = getCellValue(basisRef, context);
      basis = parseInt(String(cellValue ?? '0'));
    } else {
      basis = parseInt(basisRef);
    }
    
    const startDate = parseDate(startDateValue);
    const endDate = parseDate(endDateValue);
    
    if (!startDate?.isValid() || !endDate?.isValid()) {
      return FormulaError.VALUE;
    }
    
    const daysDiff = endDate.diff(startDate, 'day');
    
    switch (basis) {
      case 0: // 30/360 US
        return daysDiff / 360;
      case 1: // Actual/actual
        const years = endDate.diff(startDate, 'year', true);
        return years;
      case 2: // Actual/360
        return daysDiff / 360;
      case 3: // Actual/365
        return daysDiff / 365;
      case 4: // 30/360 European
        return daysDiff / 360;
      default:
        return FormulaError.NUM;
    }
  }
};