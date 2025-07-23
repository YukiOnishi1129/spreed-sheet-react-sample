// Date オブジェクトベースの日付ユーティリティ関数

export interface DateConfig {
  timezone?: string;
  locale?: string;
}

/**
 * グローバル設定
 */
let globalDateConfig: DateConfig = {
  timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
  locale: 'ja-JP'
};

/**
 * グローバル日付設定を更新
 */
export function setGlobalDateConfig(config: Partial<DateConfig>): void {
  globalDateConfig = { ...globalDateConfig, ...config };
}

/**
 * 現在のグローバル日付設定を取得
 */
export function getGlobalDateConfig(): DateConfig {
  return { ...globalDateConfig };
}

/**
 * タイムゾーン対応の日付作成
 */
export function createDateInTimezone(year: number, month: number, day: number, config?: DateConfig): Date {
  const cfg = { ...globalDateConfig, ...config };
  
  if (cfg.timezone) {
    // 指定されたタイムゾーンで日付を作成
    const isoString = `${year.toString().padStart(4, '0')}-${(month).toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}T00:00:00`;
    const date = new Date(isoString);
    
    // タイムゾーンオフセットを適用
    const formatter = new Intl.DateTimeFormat('en-CA', {
      timeZone: cfg.timezone,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    });
    
    const parts = formatter.formatToParts(date);
    const partsObj: Record<string, string> = parts.reduce((acc, part) => {
      acc[part.type] = part.value;
      return acc;
    }, {} as Record<string, string>);
    
    return new Date(
      parseInt(partsObj.year),
      parseInt(partsObj.month) - 1,
      parseInt(partsObj.day)
    );
  }
  
  return new Date(year, month - 1, day);
}

/**
 * ロケール対応の日付フォーマット
 */
export function formatDate(date: Date, format?: string, config?: DateConfig): string {
  const cfg = { ...globalDateConfig, ...config };
  
  const options: Intl.DateTimeFormatOptions = {
    timeZone: cfg.timezone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  };
  
  if (format) {
    // カスタムフォーマット対応
    const formatter = new Intl.DateTimeFormat(cfg.locale, options);
    const parts = formatter.formatToParts(date);
    const partsObj: Record<string, string> = parts.reduce((acc, part) => {
      acc[part.type] = part.value;
      return acc;
    }, {} as Record<string, string>);
    
    return format
      .replace(/YYYY/g, partsObj.year || '')
      .replace(/MM/g, partsObj.month || '')
      .replace(/DD/g, partsObj.day || '');
  }
  
  return new Intl.DateTimeFormat(cfg.locale, options).format(date);
}

/**
 * 週の開始日を考慮した曜日計算
 */
export function getWeekday(date: Date, startOfWeek: number = 1, config?: DateConfig): number {
  // startOfWeek: 1=月曜日, 0=日曜日
  const cfg = { ...globalDateConfig, ...config };
  
  let dayOfWeek = date.getDay(); // 0=日曜日, 1=月曜日, ..., 6=土曜日
  
  if (startOfWeek === 1) {
    // 月曜日を1とする場合
    dayOfWeek = dayOfWeek === 0 ? 7 : dayOfWeek;
  } else {
    // 日曜日を1とする場合
    dayOfWeek = dayOfWeek + 1;
  }
  
  return dayOfWeek;
}

/**
 * 日付文字列やExcelシリアル値をDateオブジェクトに変換
 */
export function parseDate(input: unknown): Date | null {
  if (!input) {
    return null;
  }
  
  // すでにDateオブジェクトの場合
  if (input instanceof Date) {
    return isValidDate(input) ? input : null;
  }
  
  // 数値の場合（Excelシリアル値）
  if (typeof input === 'number') {
    return excelSerialToDate(input);
  }
  
  // 文字列の場合
  if (typeof input === 'string') {
    return parseDateString(input);
  }
  
  return null;
}

/**
 * 文字列を日付に変換
 */
function parseDateString(dateStr: string): Date | null {
  
  // ISO形式やYYYY-MM-DD形式を試す
  let date = new Date(dateStr);
  if (isValidDate(date)) {
    return date;
  }
  
  // MM/DD/YYYY形式を変換
  const mmddyyyy = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mmddyyyy) {
    const [, month, day, year] = mmddyyyy;
    date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    if (isValidDate(date)) {
      return date;
    }
  }
  
  // DD/MM/YYYY形式を変換（必要に応じて）
  const ddmmyyyy = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmmyyyy) {
    const [, day, month, year] = ddmmyyyy;
    date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    if (isValidDate(date)) {
      return date;
    }
  }
  
  return null;
}

/**
 * Excelシリアル値をDateオブジェクトに変換
 */
export function excelSerialToDate(serial: number): Date {
  // Excel epoch: 1900/1/1 (実際は1899/12/30)
  const excelEpoch = new Date(1899, 11, 30); // 月は0ベース
  const result = new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
  return result;
}

/**
 * DateオブジェクトをExcelシリアル値に変換
 */
export function dateToExcelSerial(date: Date): number {
  const excelEpoch = new Date(1899, 11, 30);
  const diffTime = date.getTime() - excelEpoch.getTime();
  return Math.floor(diffTime / (24 * 60 * 60 * 1000));
}

/**
 * 有効なDateオブジェクトかチェック
 */
export function isValidDate(date: Date): boolean {
  return date instanceof Date && !isNaN(date.getTime());
}

/**
 * 2つの日付の差を日数で計算
 */
export function diffDays(startDate: Date, endDate: Date): number {
  const start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  const diffTime = end.getTime() - start.getTime();
  return Math.floor(diffTime / (24 * 60 * 60 * 1000));
}

/**
 * 2つの日付の差を月数で計算
 */
export function diffMonths(startDate: Date, endDate: Date): number {
  const yearDiff = endDate.getFullYear() - startDate.getFullYear();
  const monthDiff = endDate.getMonth() - startDate.getMonth();
  return yearDiff * 12 + monthDiff;
}

/**
 * 2つの日付の差を年数で計算
 */
export function diffYears(startDate: Date, endDate: Date): number {
  let years = endDate.getFullYear() - startDate.getFullYear();
  
  // 月日を考慮して調整
  if (endDate.getMonth() < startDate.getMonth() || 
      (endDate.getMonth() === startDate.getMonth() && endDate.getDate() < startDate.getDate())) {
    years--;
  }
  
  return years;
}

/**
 * 年を無視した月数の差（YM単位）
 */
export function diffMonthsYM(startDate: Date, endDate: Date): number {
  let monthDiff = endDate.getMonth() - startDate.getMonth();
  
  // 日を考慮して調整
  if (endDate.getDate() < startDate.getDate()) {
    monthDiff--;
  }
  
  // 負の場合は12を足す
  if (monthDiff < 0) {
    monthDiff += 12;
  }
  
  return monthDiff;
}

/**
 * 年を無視した日数の差（YD単位）
 */
export function diffDaysYD(startDate: Date, endDate: Date): number {
  // 終了日の年を開始日の年に合わせる
  const adjustedEndDate = new Date(
    startDate.getFullYear(),
    endDate.getMonth(),
    endDate.getDate()
  );
  
  let dayDiff = diffDays(startDate, adjustedEndDate);
  
  // 負の場合は翌年で計算
  if (dayDiff < 0) {
    const nextYearEndDate = new Date(
      startDate.getFullYear() + 1,
      endDate.getMonth(),
      endDate.getDate()
    );
    dayDiff = diffDays(startDate, nextYearEndDate);
  }
  
  return dayDiff;
}

/**
 * 月を無視した日数の差（MD単位）
 */
export function diffDaysMD(startDate: Date, endDate: Date): number {
  let dayDiff = endDate.getDate() - startDate.getDate();
  
  if (dayDiff < 0) {
    // 前月の日数を取得
    const prevMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 0);
    const daysInPrevMonth = prevMonth.getDate();
    dayDiff = daysInPrevMonth - startDate.getDate() + endDate.getDate();
  }
  
  return dayDiff;
}

/**
 * 現在の日時を取得（タイムゾーン対応）
 */
export function now(config?: DateConfig): Date {
  const cfg = { ...globalDateConfig, ...config };
  const now = new Date();
  
  if (cfg.timezone) {
    // タイムゾーン変換（簡易版）
    const options = {
      timeZone: cfg.timezone
    } as const;
    const formatter = new Intl.DateTimeFormat('en-CA', {
      ...options,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    });
    
    const parts = formatter.formatToParts(now);
    const partsObj: Record<string, string> = parts.reduce((acc, part) => {
      acc[part.type] = part.value;
      return acc;
    }, {} as Record<string, string>);
    
    return new Date(
      parseInt(partsObj.year),
      parseInt(partsObj.month) - 1,
      parseInt(partsObj.day),
      parseInt(partsObj.hour),
      parseInt(partsObj.minute),
      parseInt(partsObj.second)
    );
  }
  
  return now;
}

/**
 * 今日の日付を取得（時刻は00:00:00）
 */
export function today(config?: DateConfig): Date {
  const todayDate = now(config);
  return new Date(todayDate.getFullYear(), todayDate.getMonth(), todayDate.getDate());
}