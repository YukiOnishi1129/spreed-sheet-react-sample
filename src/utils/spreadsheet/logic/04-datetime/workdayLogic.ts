// 稼働日関連の日付関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';
import { dateToExcelSerial } from './dateUtils';

// 日付を正規化する関数
function normalizeDate(value: unknown): Date | null {
  if (value instanceof Date) {
    return value;
  }
  
  if (typeof value === 'string') {
    // 引用符を除去
    const cleanStr = value.replace(/^["']|["']$/g, '');
    const date = new Date(cleanStr);
    return isNaN(date.getTime()) ? null : date;
  }
  
  if (typeof value === 'number') {
    // Excelのシリアル値として扱う（1900年1月1日が1）
    // Excelの1900年うるう年バグを考慮
    let adjustedSerial = value;
    if (value > 60) {
      adjustedSerial = value - 1;
    }
    
    const excelEpoch = new Date(1899, 11, 31); // 1899年12月31日
    const date = new Date(excelEpoch.getTime() + adjustedSerial * 24 * 60 * 60 * 1000);
    return date;
  }
  
  return null;
}

// 週末かどうかを判定（デフォルトは土日） - Excel互換
function isWeekend(date: Date, weekendDays: string = '1'): boolean {
  const dayOfWeek = date.getDay(); // 0=日曜日, 1=月曜日, ..., 6=土曜日
  
  // 数値で渡された場合は文字列に変換
  const weekendStr = String(weekendDays);
  
  switch (weekendStr) {
    case '1': // 土日（デフォルト）
      return dayOfWeek === 0 || dayOfWeek === 6;
    case '2': // 日月
      return dayOfWeek === 0 || dayOfWeek === 1;
    case '3': // 月火
      return dayOfWeek === 1 || dayOfWeek === 2;
    case '4': // 火水
      return dayOfWeek === 2 || dayOfWeek === 3;
    case '5': // 水木
      return dayOfWeek === 3 || dayOfWeek === 4;
    case '6': // 木金
      return dayOfWeek === 4 || dayOfWeek === 5;
    case '7': // 金土
      return dayOfWeek === 5 || dayOfWeek === 6;
    case '11': // 日のみ
      return dayOfWeek === 0;
    case '12': // 月のみ
      return dayOfWeek === 1;
    case '13': // 火のみ
      return dayOfWeek === 2;
    case '14': // 水のみ
      return dayOfWeek === 3;
    case '15': // 木のみ
      return dayOfWeek === 4;
    case '16': // 金のみ
      return dayOfWeek === 5;
    case '17': // 土のみ
      return dayOfWeek === 6;
    default:
      // カスタム週末（7文字の文字列で各曜日を指定）
      if (weekendStr.length === 7) {
        return weekendStr[dayOfWeek] === '1';
      }
      // 無効な場合はデフォルト（土日）にフォールバック
      return dayOfWeek === 0 || dayOfWeek === 6;
  }
}

// 祝日リストを解析
function parseHolidays(holidaysRef: string, context: FormulaContext): Date[] {
  const holidays: Date[] = [];
  
  if (holidaysRef.includes(':')) {
    // セル範囲の場合
    const rangeMatch = holidaysRef.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (rangeMatch) {
      const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
      const startRow = parseInt(startRowStr) - 1;
      const endRow = parseInt(endRowStr) - 1;
      const startColIndex = startCol.charCodeAt(0) - 65;
      const endColIndex = endCol.charCodeAt(0) - 65;
      
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startColIndex; col <= endColIndex; col++) {
          if (context.data[row]?.[col] !== undefined) {
            const cellData = context.data[row][col];
            // セルデータから値を取得
            let value;
            if (typeof cellData === 'string' || typeof cellData === 'number' || cellData === null) {
              value = cellData;
            } else if (cellData && typeof cellData === 'object') {
              value = cellData.value ?? cellData.v ?? cellData._ ?? cellData;
            } else {
              value = cellData;
            }
            
            const date = normalizeDate(value);
            if (date) {
              holidays.push(date);
            }
          }
        }
      }
    }
  } else {
    // 単一セルの場合
    const value = getCellValue(holidaysRef, context);
    // セル参照エラーの場合はスキップ
    if (value !== FormulaError.REF && value !== null && value !== undefined) {
      const date = normalizeDate(value);
      if (date) {
        holidays.push(date);
      }
    }
  }
  
  return holidays;
}

// 日付が祝日かどうかを判定
function isHoliday(date: Date, holidays: Date[]): boolean {
  const dateStr = date.toDateString();
  return holidays.some(holiday => holiday.toDateString() === dateStr);
}

// NETWORKDAYS関数の実装（稼働日数を計算）
export const NETWORKDAYS: CustomFormula = {
  name: 'NETWORKDAYS',
  pattern: /NETWORKDAYS\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, startDateRef, endDateRef, holidaysRef] = matches;
    
    try {
      // 開始日を取得
      const startValue = getCellValue(startDateRef.trim(), context);
      // セル参照エラーの場合は、直接値として解釈を試みる
      const startDateInput = (startValue === FormulaError.REF || startValue === null || startValue === undefined) 
        ? startDateRef.trim() 
        : startValue;
      const startDate = normalizeDate(startDateInput);
      
      if (!startDate) {
        return FormulaError.VALUE;
      }
      
      // 終了日を取得
      const endValue = getCellValue(endDateRef.trim(), context);
      // セル参照エラーの場合は、直接値として解釈を試みる
      const endDateInput = (endValue === FormulaError.REF || endValue === null || endValue === undefined)
        ? endDateRef.trim()
        : endValue;
      const endDate = normalizeDate(endDateInput);
      
      if (!endDate) {
        return FormulaError.VALUE;
      }
      
      // 祝日リストを取得
      const holidays = holidaysRef ? parseHolidays(holidaysRef.trim(), context) : [];
      
      // 日付の順序を正規化
      const start = new Date(Math.min(startDate.getTime(), endDate.getTime()));
      const end = new Date(Math.max(startDate.getTime(), endDate.getTime()));
      const isReverse = startDate > endDate;
      
      // 稼働日数をカウント
      let workdays = 0;
      const current = new Date(start);
      
      // Excel's NETWORKDAYS includes both start and end dates
      while (current <= end) {
        if (!isWeekend(current) && !isHoliday(current, holidays)) {
          workdays++;
        }
        current.setDate(current.getDate() + 1);
      }
      
      return isReverse ? -workdays : workdays;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// NETWORKDAYS.INTL関数の実装（国際版稼働日数）
export const NETWORKDAYS_INTL: CustomFormula = {
  name: 'NETWORKDAYS.INTL',
  pattern: /NETWORKDAYS\.INTL\(([^,]+),\s*([^,]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, startDateRef, endDateRef, weekendRef, holidaysRef] = matches;
    
    try {
      // 開始日を取得
      const startValue = getCellValue(startDateRef.trim(), context);
      // セル参照エラーの場合は、直接値として解釈を試みる
      const startDateInput = (startValue === FormulaError.REF || startValue === null || startValue === undefined) 
        ? startDateRef.trim() 
        : startValue;
      const startDate = normalizeDate(startDateInput);
      
      if (!startDate) {
        return FormulaError.VALUE;
      }
      
      // 終了日を取得
      const endValue = getCellValue(endDateRef.trim(), context);
      // セル参照エラーの場合は、直接値として解釈を試みる
      const endDateInput = (endValue === FormulaError.REF || endValue === null || endValue === undefined)
        ? endDateRef.trim()
        : endValue;
      const endDate = normalizeDate(endDateInput);
      
      if (!endDate) {
        return FormulaError.VALUE;
      }
      
      // 週末の定義を取得（デフォルトは1=土日）
      let weekendDays = '1';
      if (weekendRef) {
        const weekendValue = getCellValue(weekendRef.trim(), context);
        const weekendInput = (weekendValue === FormulaError.REF || weekendValue === null || weekendValue === undefined)
          ? weekendRef.trim()
          : weekendValue;
        weekendDays = String(weekendInput).replace(/['"]/g, '');
      }
      
      // 祝日リストを取得
      const holidays = holidaysRef ? parseHolidays(holidaysRef.trim(), context) : [];
      
      // 日付の順序を正規化
      const start = new Date(Math.min(startDate.getTime(), endDate.getTime()));
      const end = new Date(Math.max(startDate.getTime(), endDate.getTime()));
      const isReverse = startDate > endDate;
      
      // 稼働日数をカウント
      let workdays = 0;
      const current = new Date(start);
      
      while (current <= end) {
        if (!isWeekend(current, weekendDays) && !isHoliday(current, holidays)) {
          workdays++;
        }
        current.setDate(current.getDate() + 1);
      }
      
      return isReverse ? -workdays : workdays;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// WORKDAY関数の実装（稼働日を計算）
export const WORKDAY: CustomFormula = {
  name: 'WORKDAY',
  pattern: /WORKDAY\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, startDateRef, daysRef, holidaysRef] = matches;
    
    try {
      // 開始日を取得
      const startValue = getCellValue(startDateRef.trim(), context);
      // セル参照エラーの場合は、直接値として解釈を試みる
      const startDateInput = (startValue === FormulaError.REF || startValue === null || startValue === undefined) 
        ? startDateRef.trim() 
        : startValue;
      const startDate = normalizeDate(startDateInput);
      
      if (!startDate) {
        return FormulaError.VALUE;
      }
      
      // 日数を取得
      const daysValue = getCellValue(daysRef.trim(), context);
      const daysInput = (daysValue === FormulaError.REF || daysValue === null || daysValue === undefined)
        ? daysRef.trim()
        : daysValue;
      const days = parseInt(String(daysInput));
      
      if (isNaN(days)) {
        return FormulaError.VALUE;
      }
      
      // 祝日リストを取得
      const holidays = holidaysRef ? parseHolidays(holidaysRef.trim(), context) : [];
      
      // 稼働日を計算
      let current = new Date(startDate);
      let remainingDays = Math.abs(days);
      const isForward = days >= 0;
      
      // 0日の場合は開始日を返す
      if (days === 0) {
        return dateToExcelSerial(startDate);
      }
      
      // ExcelのWORKDAYのロジックに合わせる
      while (remainingDays > 0) {
        if (isForward) {
          current.setDate(current.getDate() + 1);
        } else {
          current.setDate(current.getDate() - 1);
        }
        
        // 週末や祝日でない場合のみカウントを減らす
        if (!isWeekend(current) && !isHoliday(current, holidays)) {
          remainingDays--;
        }
      }
      
      return dateToExcelSerial(current);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// WORKDAY.INTL関数の実装（国際版稼働日）
export const WORKDAY_INTL: CustomFormula = {
  name: 'WORKDAY.INTL',
  pattern: /WORKDAY\.INTL\(([^,]+),\s*([^,]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, startDateRef, daysRef, weekendRef, holidaysRef] = matches;
    
    try {
      // 開始日を取得
      const startValue = getCellValue(startDateRef.trim(), context);
      // セル参照エラーの場合は、直接値として解釈を試みる
      const startDateInput = (startValue === FormulaError.REF || startValue === null || startValue === undefined) 
        ? startDateRef.trim() 
        : startValue;
      const startDate = normalizeDate(startDateInput);
      
      if (!startDate) {
        return FormulaError.VALUE;
      }
      
      // 日数を取得
      const daysValue = getCellValue(daysRef.trim(), context);
      const daysInput = (daysValue === FormulaError.REF || daysValue === null || daysValue === undefined)
        ? daysRef.trim()
        : daysValue;
      const days = parseInt(String(daysInput));
      
      if (isNaN(days)) {
        return FormulaError.VALUE;
      }
      
      // 週末の定義を取得（デフォルトは1=土日）
      let weekendDays = '1';
      if (weekendRef) {
        const weekendValue = getCellValue(weekendRef.trim(), context);
        const weekendInput = (weekendValue === FormulaError.REF || weekendValue === null || weekendValue === undefined)
          ? weekendRef.trim()
          : weekendValue;
        weekendDays = String(weekendInput).replace(/['"]/g, '');
      }
      
      // 祝日リストを取得
      const holidays = holidaysRef ? parseHolidays(holidaysRef.trim(), context) : [];
      
      // 稼働日を計算
      let current = new Date(startDate);
      let remainingDays = Math.abs(days);
      const isForward = days >= 0;
      
      // 0日の場合は開始日を返す
      if (days === 0) {
        return dateToExcelSerial(startDate);
      }
      
      // ExcelのWORKDAY.INTLのロジックに合わせる
      while (remainingDays > 0) {
        if (isForward) {
          current.setDate(current.getDate() + 1);
        } else {
          current.setDate(current.getDate() - 1);
        }
        
        // 週末や祝日でない場合のみカウントを減らす
        if (!isWeekend(current, weekendDays) && !isHoliday(current, holidays)) {
          remainingDays--;
        }
      }
      
      return dateToExcelSerial(current);
    } catch {
      return FormulaError.VALUE;
    }
  }
};