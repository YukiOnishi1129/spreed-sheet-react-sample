// 日付関数の実装

import type { CustomFormula, FormulaContext } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';
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
  isValidDate,
  getWeekNumber,
  getISOWeekNumber,
  parseTimeString
} from './dateUtils';

// DATEDIF関数の実装（Date版）
export const DATEDIF: CustomFormula = {
  name: 'DATEDIF',
  pattern: /DATEDIF\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, startRef, endRef, unitRef] = matches;
    
    // セル参照か直接値かを判定
    let startValue, endValue;
    
    // 開始日の取得
    if (startRef.startsWith('"') && startRef.endsWith('"')) {
      startValue = startRef.slice(1, -1);
    } else if (startRef.match(/^[A-Z]+\d+$/)) {
      startValue = getCellValue(startRef.trim(), context);
      // エラー値を直接返す
      if (typeof startValue === 'string' && startValue.startsWith('#')) {
        return startValue as FormulaResult;
      }
    } else {
      startValue = startRef;
    }
    
    // 終了日の取得
    if (endRef.startsWith('"') && endRef.endsWith('"')) {
      endValue = endRef.slice(1, -1);
    } else if (endRef.match(/^[A-Z]+\d+$/)) {
      endValue = getCellValue(endRef.trim(), context);
      // エラー値を直接返す
      if (typeof endValue === 'string' && endValue.startsWith('#')) {
        return endValue as FormulaResult;
      }
    } else {
      endValue = endRef;
    }
    
    // 単位の取得（セル参照または直接文字列）
    let unit;
    if (unitRef.startsWith('"') && unitRef.endsWith('"')) {
      unit = unitRef.slice(1, -1);
    } else if (unitRef.match(/^[A-Z]+\d+$/)) {
      unit = getCellValue(unitRef.trim(), context);
      // エラー値を直接返す
      if (typeof unit === 'string' && unit.startsWith('#')) {
        return unit as FormulaResult;
      }
    } else {
      unit = unitRef;
    }
    
    // Date版のparseDate関数を使用
    const startDate = parseDateNew(startValue);
    const endDate = parseDateNew(endValue);
    
    if (!startDate || !endDate) {
      return FormulaError.NUM;
    }
    
    if (startDate.getTime() > endDate.getTime()) {
      return FormulaError.NUM;
    }
    
    switch (String(unit).toUpperCase()) {
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
  pattern: /\bDAY\(([^)]+)\)/i,
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
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, dateRef, typeRef] = matches;
    
    const dateValue = getCellValue(dateRef, context) ?? dateRef;
    const date = parseDateNew(dateValue);
    
    if (!date) {
      return FormulaError.VALUE;
    }
    
    // タイプ指定（デフォルトは1）
    const type = typeRef ? Number(getCellValue(typeRef, context) ?? typeRef) : 1;
    
    const dayOfWeek = date.getDay(); // 0=日曜日, 1=月曜日, ..., 6=土曜日
    
    switch (type) {
      case 1: // 1=日曜日, 2=月曜日, ..., 7=土曜日
        return dayOfWeek + 1;
      case 2: // 1=月曜日, 2=火曜日, ..., 7=日曜日
        return dayOfWeek === 0 ? 7 : dayOfWeek;
      case 3: // 0=月曜日, 1=火曜日, ..., 6=日曜日
        return dayOfWeek === 0 ? 6 : dayOfWeek - 1;
      default:
        return FormulaError.VALUE;
    }
  }
};

// DAYS関数の実装（日数差）
export const DAYS: CustomFormula = {
  name: 'DAYS',
  pattern: /DAYS\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, endDateRef, startDateRef] = matches;
    
    const startValue = getCellValue(startDateRef, context) ?? startDateRef;
    const endValue = getCellValue(endDateRef, context) ?? endDateRef;
    
    const startDate = parseDateNew(startValue);
    const endDate = parseDateNew(endValue);
    
    if (!startDate || !endDate) {
      return FormulaError.VALUE;
    }
    
    return diffDays(startDate, endDate);
  }
};

// EDATE関数の実装（月数後の日付）
export const EDATE: CustomFormula = {
  name: 'EDATE',
  pattern: /EDATE\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, dateRef, monthsRef] = matches;
    
    const dateValue = getCellValue(dateRef, context) ?? dateRef;
    const monthsValue = getCellValue(monthsRef, context) ?? monthsRef;
    
    const date = parseDateNew(dateValue);
    const months = Number(monthsValue);
    
    if (!date || isNaN(months)) {
      return FormulaError.VALUE;
    }
    
    // Excelのロジック：月末日の場合、結果も月末日にする
    const day = date.getDate();
    const lastDayOfMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
    const isEndOfMonth = day === lastDayOfMonth;
    
    // 月を加算
    let newDate = new Date(date.getFullYear(), date.getMonth() + months, day);
    
    // 元の日付が月末なら、新しい日付も月末にする
    if (isEndOfMonth) {
      const newLastDay = new Date(newDate.getFullYear(), newDate.getMonth() + 1, 0).getDate();
      newDate = new Date(newDate.getFullYear(), newDate.getMonth(), Math.min(day, newLastDay));
    }
    
    if (!isValidDate(newDate)) {
      return FormulaError.VALUE;
    }
    
    return dateToExcelSerial(newDate);
  }
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
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, dateRef, monthsRef] = matches;
    
    const dateValue = getCellValue(dateRef, context) ?? dateRef;
    const monthsValue = getCellValue(monthsRef, context) ?? monthsRef;
    
    const date = parseDateNew(dateValue);
    const months = Number(monthsValue);
    
    if (!date || isNaN(months)) {
      return FormulaError.VALUE;
    }
    
    // 指定した月数後の月末を取得
    // new Date(year, month + 1, 0) は指定月の最終日を返す
    const targetMonth = new Date(date.getFullYear(), date.getMonth() + months + 1, 0);
    
    if (!isValidDate(targetMonth)) {
      return FormulaError.VALUE;
    }
    
    return dateToExcelSerial(targetMonth);
  }
};

// TIME関数の実装（時刻を作成）
export const TIME: CustomFormula = {
  name: 'TIME',
  pattern: /TIME\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, hourRef, minuteRef, secondRef] = matches;
    
    const hour = Number(getCellValue(hourRef, context) ?? hourRef);
    const minute = Number(getCellValue(minuteRef, context) ?? minuteRef);
    const second = Number(getCellValue(secondRef, context) ?? secondRef);
    
    if (isNaN(hour) || isNaN(minute) || isNaN(second)) {
      return FormulaError.VALUE;
    }
    
    // 時刻を日の割合で返す（Excel的な時刻値）
    // オーバーフローを処理（24時間以上、60分以上、60秒以上）
    const totalSeconds = hour * 3600 + minute * 60 + second;
    const normalizedSeconds = totalSeconds % 86400; // 24時間でモジュロ演算
    return normalizedSeconds / 86400; // 1日 = 86400秒
  }
};

// HOUR関数の実装（時を抽出）
export const HOUR: CustomFormula = {
  name: 'HOUR',
  pattern: /HOUR\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, timeRef] = matches;
    
    const timeValue = getCellValue(timeRef, context) ?? timeRef;
    
    // 数値の場合（Excel的な時刻値）
    if (typeof timeValue === 'number') {
      const totalSeconds = Math.floor(timeValue * 86400);
      return Math.floor(totalSeconds / 3600) % 24;
    }
    
    // 文字列の場合、時刻として解析を試みる
    if (typeof timeValue === 'string') {
      const timeDecimal = parseTimeString(timeValue);
      if (timeDecimal !== null) {
        const totalSeconds = Math.floor(timeDecimal * 86400);
        return Math.floor(totalSeconds / 3600) % 24;
      }
    }
    
    // 日付時刻の場合
    const date = parseDateNew(timeValue);
    if (!date) {
      return FormulaError.VALUE;
    }
    
    return date.getHours();
  }
};

// MINUTE関数の実装（分を抽出）
export const MINUTE: CustomFormula = {
  name: 'MINUTE',
  pattern: /MINUTE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, timeRef] = matches;
    
    const timeValue = getCellValue(timeRef, context) ?? timeRef;
    
    // 数値の場合（Excel的な時刻値）
    if (typeof timeValue === 'number') {
      const totalSeconds = Math.floor(timeValue * 86400);
      return Math.floor((totalSeconds % 3600) / 60);
    }
    
    // 文字列の場合、時刻として解析を試みる
    if (typeof timeValue === 'string') {
      const timeDecimal = parseTimeString(timeValue);
      if (timeDecimal !== null) {
        const totalSeconds = Math.floor(timeDecimal * 86400);
        return Math.floor((totalSeconds % 3600) / 60);
      }
    }
    
    // 日付時刻の場合
    const date = parseDateNew(timeValue);
    if (!date) {
      return FormulaError.VALUE;
    }
    
    return date.getMinutes();
  }
};

// SECOND関数の実装（秒を抽出）
export const SECOND: CustomFormula = {
  name: 'SECOND',
  pattern: /SECOND\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, timeRef] = matches;
    
    const timeValue = getCellValue(timeRef, context) ?? timeRef;
    
    // 数値の場合（Excel的な時刻値）
    if (typeof timeValue === 'number') {
      const totalSeconds = Math.floor(timeValue * 86400);
      return totalSeconds % 60;
    }
    
    // 文字列の場合、時刻として解析を試みる
    if (typeof timeValue === 'string') {
      const timeDecimal = parseTimeString(timeValue);
      if (timeDecimal !== null) {
        const totalSeconds = Math.floor(timeDecimal * 86400);
        return totalSeconds % 60;
      }
    }
    
    // 日付時刻の場合
    const date = parseDateNew(timeValue);
    if (!date) {
      return FormulaError.VALUE;
    }
    
    return date.getSeconds();
  }
};

// WEEKNUM関数の実装（週番号を返す）
export const WEEKNUM: CustomFormula = {
  name: 'WEEKNUM',
  pattern: /WEEKNUM\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, dateRef, returnTypeRef] = matches;
    
    const dateValue = getCellValue(dateRef, context) ?? dateRef;
    const date = parseDateNew(dateValue);
    
    if (!date) {
      return FormulaError.VALUE;
    }
    
    const returnType = returnTypeRef ? Number(getCellValue(returnTypeRef, context) ?? returnTypeRef) : 1;
    
    try {
      return getWeekNumber(date, returnType);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DAYS360関数の実装（360日基準の日数）
export const DAYS360: CustomFormula = {
  name: 'DAYS360',
  pattern: /DAYS360\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, startDateRef, endDateRef, methodRef] = matches;
    
    const startValue = getCellValue(startDateRef, context) ?? startDateRef;
    const endValue = getCellValue(endDateRef, context) ?? endDateRef;
    
    const startDate = parseDateNew(startValue);
    const endDate = parseDateNew(endValue);
    
    if (!startDate || !endDate) {
      return FormulaError.VALUE;
    }
    
    const methodValue = methodRef ? (getCellValue(methodRef, context) ?? methodRef) : false;
    const isEuropean = methodValue === true || methodValue === 'TRUE' || methodValue === 1 || methodValue === '1' || String(methodValue).toUpperCase() === 'TRUE';
    
    let startDay = startDate.getDate();
    const startMonth = startDate.getMonth() + 1;
    const startYear = startDate.getFullYear();
    
    let endDay = endDate.getDate();
    const endMonth = endDate.getMonth() + 1;
    const endYear = endDate.getFullYear();
    
    if (isEuropean) {
      // ヨーロッパ式（30E/360） - Excel互換
      if (startDay === 31) startDay = 30;
      if (endDay === 31) endDay = 30;
    } else {
      // アメリカ式（30U/360） - Excel互換
      if (startDay === 31) startDay = 30;
      if (endDay === 31 && startDay === 30) endDay = 30;
    }
    
    const days = (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay);
    
    // Europeanメソッドの場合は1を加算（Excel仕様）
    return isEuropean ? days + 1 : days;
  }
};

// YEARFRAC関数の実装（年の割合を計算）- Excel完全互換版
export const YEARFRAC: CustomFormula = {
  name: 'YEARFRAC',
  pattern: /YEARFRAC\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, startDateRef, endDateRef, basisRef] = matches;
    
    const startValue = getCellValue(startDateRef, context) ?? startDateRef;
    const endValue = getCellValue(endDateRef, context) ?? endDateRef;
    
    const startDate = parseDateNew(startValue);
    const endDate = parseDateNew(endValue);
    
    if (!startDate || !endDate) {
      return FormulaError.VALUE;
    }
    
    const basis = basisRef ? Number(getCellValue(basisRef, context) ?? basisRef) : 0;
    
    // 日付の順序を正しくする
    const earlier = startDate <= endDate ? startDate : endDate;
    const later = startDate <= endDate ? endDate : startDate;
    
    switch (basis) {
      case 0: // 30/360 US (NASD)
        return yearfrac30_360US(earlier, later);
      case 1: // Actual/actual
        return yearfracActualActual(earlier, later);
      case 2: // Actual/360
        return yearfracActual360(earlier, later);
      case 3: // Actual/365
        return yearfracActual365(earlier, later);
      case 4: // 30/360 European
        return yearfrac30_360European(earlier, later);
      default:
        return FormulaError.NUM;
    }
  }
};

// 30/360 US (NASD) 方式 - Excel完全互換
function yearfrac30_360US(startDate: Date, endDate: Date): number {
  const startYear = startDate.getFullYear();
  const startMonth = startDate.getMonth() + 1;
  let startDay = startDate.getDate();
  
  const endYear = endDate.getFullYear();
  const endMonth = endDate.getMonth() + 1;
  let endDay = endDate.getDate();
  
  // US 30/360の特殊ルール（Excel互換）
  const lastDayOfStartMonth = daysInMonth(startYear, startMonth);
  const lastDayOfEndMonth = daysInMonth(endYear, endMonth);
  
  // Rule 1: 開始日が31日の場合は30日にする
  if (startDay === 31) {
    startDay = 30;
  }
  
  // Rule 2: 終了日が31日で開始日が30日以上の場合は30日にする
  if (endDay === 31 && startDay >= 30) {
    endDay = 30;
  }
  
  // Rule 3: 開始日が月末（2月29日など）の場合は30日にする
  if (startDay === lastDayOfStartMonth) {
    startDay = 30;
  }
  
  // Rule 4: 終了日が月末で開始日が30日の場合は30日にする
  if (endDay === lastDayOfEndMonth && startDay === 30) {
    endDay = 30;
  }
  
  const days = (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay);
  return days / 360;
}

// 30/360 European方式 - Excel完全互換
function yearfrac30_360European(startDate: Date, endDate: Date): number {
  const startYear = startDate.getFullYear();
  const startMonth = startDate.getMonth() + 1;
  let startDay = startDate.getDate();
  
  const endYear = endDate.getFullYear();
  const endMonth = endDate.getMonth() + 1;
  let endDay = endDate.getDate();
  
  // European 30/360のルール（Excel互換、basis 0と同じ）
  const lastDayOfStartMonth = daysInMonth(startYear, startMonth);
  const lastDayOfEndMonth = daysInMonth(endYear, endMonth);
  
  // Rule 1: 開始日が31日の場合は30日にする
  if (startDay === 31) {
    startDay = 30;
  }
  
  // Rule 2: 終了日が31日で開始日が30日以上の場合は30日にする
  if (endDay === 31 && startDay >= 30) {
    endDay = 30;
  }
  
  // Rule 3: 開始日が月末（2月29日など）の場合は30日にする
  if (startDay === lastDayOfStartMonth) {
    startDay = 30;
  }
  
  // Rule 4: 終了日が月末で開始日が30日の場合は30日にする
  if (endDay === lastDayOfEndMonth && startDay === 30) {
    endDay = 30;
  }
  
  const days = (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay);
  return days / 360;
}

// Actual/Actual方式 - Excel完全互換（ISMA Actual/Actual） 
function yearfracActualActual(startDate: Date, endDate: Date): number {
  const startYear = startDate.getFullYear();
  const endYear = endDate.getFullYear();
  
  // Excelでは終了日も含めて計算するため、+1日
  const actualDays = Math.floor((endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000)) + 1;
  
  if (startYear === endYear) {
    // 同年内の場合
    // Excelの実装では、うるう年の場合は366で割り、平年の場合は365で割る
    const yearDays = isLeapYear(startYear) ? 366 : 365;
    return actualDays / yearDays;
  }
  
  // 複数年にまたがる場合のExcelロジック（ISMA Actual/Actual）
  let result = 0;
  let currentDate = new Date(startDate);
  
  while (currentDate.getFullYear() <= endYear) {
    const currentYear = currentDate.getFullYear();
    // const yearEnd = new Date(currentYear + 1, 0, 1); // 次年の1月1日
    
    const periodStart = currentDate;
    const periodEnd = currentYear === endYear ? endDate : new Date(currentYear, 11, 31);
    
    if (periodStart <= periodEnd) {
      // 期間の日数（終了日を含む）
      const periodDays = Math.floor((periodEnd.getTime() - periodStart.getTime()) / (24 * 60 * 60 * 1000)) + 1;
      const yearDays = isLeapYear(currentYear) ? 366 : 365;
      
      result += periodDays / yearDays;
    }
    
    // 次の年の1月1日に移動
    currentDate = new Date(currentYear + 1, 0, 1);
  }
  
  return result;
}

// Actual/360方式 - Excel完全互換
function yearfracActual360(startDate: Date, endDate: Date): number {
  const actualDays = Math.floor((endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000));
  return actualDays / 360;
}

// Actual/365方式 - Excel完全互換
function yearfracActual365(startDate: Date, endDate: Date): number {
  const actualDays = Math.floor((endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000));
  return actualDays / 365;
}

// ヘルパー関数：月の日数を取得
function daysInMonth(year: number, month: number): number {
  return new Date(year, month, 0).getDate();
}

// ヘルパー関数：うるう年判定
function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
}

// DATEVALUE関数の実装（日付文字列を日付値に変換）
export const DATEVALUE: CustomFormula = {
  name: 'DATEVALUE',
  pattern: /DATEVALUE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, dateRef] = matches;
    
    let dateValue = getCellValue(dateRef, context) ?? dateRef;
    
    // 文字列の引用符を除去
    if (typeof dateValue === 'string' && dateValue.startsWith('"') && dateValue.endsWith('"')) {
      dateValue = dateValue.slice(1, -1);
    }
    
    const date = parseDateNew(dateValue);
    
    if (!date) {
      return FormulaError.VALUE;
    }
    
    return dateToExcelSerial(date);
  }
};

// TIMEVALUE関数の実装（時刻文字列を時刻値に変換）
export const TIMEVALUE: CustomFormula = {
  name: 'TIMEVALUE',
  pattern: /TIMEVALUE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, timeRef] = matches;
    
    let timeValue = getCellValue(timeRef, context) ?? timeRef;
    
    // 文字列の引用符を除去
    if (typeof timeValue === 'string' && timeValue.startsWith('"') && timeValue.endsWith('"')) {
      timeValue = timeValue.slice(1, -1);
    }
    
    if (typeof timeValue !== 'string') {
      return FormulaError.VALUE;
    }
    
    try {
      const time = parseTimeString(timeValue);
      
      if (time === null) {
        return FormulaError.VALUE;
      }
      
      return time;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ISOWEEKNUM関数の実装（ISO週番号を返す）
export const ISOWEEKNUM: CustomFormula = {
  name: 'ISOWEEKNUM',
  pattern: /ISOWEEKNUM\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, dateRef] = matches;
    
    const dateValue = getCellValue(dateRef, context) ?? dateRef;
    const date = parseDateNew(dateValue);
    
    if (!date) {
      return FormulaError.VALUE;
    }
    
    try {
      return getISOWeekNumber(date);
    } catch {
      return FormulaError.VALUE;
    }
  }
};



