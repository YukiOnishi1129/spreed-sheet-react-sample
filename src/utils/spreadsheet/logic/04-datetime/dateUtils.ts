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
 * よく使用されるタイムゾーンの定数
 */
export const TIMEZONES = {
  // 日本
  JAPAN: 'Asia/Tokyo',
  
  // アメリカ
  US_EASTERN: 'America/New_York',      // ニューヨーク、フロリダ等
  US_CENTRAL: 'America/Chicago',       // シカゴ、テキサス等
  US_MOUNTAIN: 'America/Denver',       // デンバー、コロラド等
  US_PACIFIC: 'America/Los_Angeles',   // ロサンゼルス、シアトル等
  US_ALASKA: 'America/Anchorage',      // アラスカ
  US_HAWAII: 'Pacific/Honolulu',       // ハワイ
  
  // ヨーロッパ
  UK: 'Europe/London',
  FRANCE: 'Europe/Paris', 
  GERMANY: 'Europe/Berlin',
  
  // その他
  AUSTRALIA_SYDNEY: 'Australia/Sydney',
  CHINA: 'Asia/Shanghai',
  KOREA: 'Asia/Seoul',
  UTC: 'UTC'
} as const;

/**
 * グローバル日付設定を更新
 * 
 * @example
 * // ニューヨーク時間に設定
 * setGlobalDateConfig({
 *   timezone: TIMEZONES.US_EASTERN,
 *   locale: 'en-US'
 * });
 * 
 * // ロサンゼルス時間に設定
 * setGlobalDateConfig({
 *   timezone: TIMEZONES.US_PACIFIC,
 *   locale: 'en-US'
 * });
 * 
 * // 日本時間に設定
 * setGlobalDateConfig({
 *   timezone: TIMEZONES.JAPAN,
 *   locale: 'ja-JP'
 * });
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
  const timezone = config?.timezone ?? globalDateConfig.timezone;
  
  if (timezone) {
    // タイムゾーンを考慮した日付を作成
    // UTCで日付を作成してから、指定のタイムゾーンでの時刻に調整
    const utcDate = new Date(Date.UTC(year, month - 1, day, 0, 0, 0));
    
    // 指定タイムゾーンでの表現を取得
    const formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: timezone,
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      second: 'numeric',
      hour12: false
    });
    
    const parts = formatter.formatToParts(utcDate);
    const dateParts: { [key: string]: string } = {};
    
    parts.forEach(part => {
      if (part.type !== 'literal') {
        dateParts[part.type] = part.value;
      }
    });
    
    // ローカル時刻として新しいDateオブジェクトを作成
    return new Date(
      parseInt(dateParts.year),
      parseInt(dateParts.month) - 1,
      parseInt(dateParts.day),
      parseInt(dateParts.hour || '0'),
      parseInt(dateParts.minute || '0'),
      parseInt(dateParts.second || '0')
    );
  }
  
  // タイムゾーン指定がない場合は通常の日付作成
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
    const initialPartsObj: Record<string, string> = {};
    const partsObj: Record<string, string> = parts.reduce((acc, part) => {
      acc[part.type] = part.value;
      return acc;
    }, initialPartsObj);
    
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
  
  // タイムゾーンを考慮した日付取得
  const localDate = cfg.timezone ? getDateInTimezone(date, cfg.timezone) : date;
  let dayOfWeek = localDate.getDay(); // 0=日曜日, 1=月曜日, ..., 6=土曜日
  
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
 * 週番号を計算（WEEKNUM関数用）
 */
export function getWeekNumber(date: Date, returnType: number = 1): number {
  const year = date.getFullYear();
  const jan1 = new Date(year, 0, 1);
  
  let weekStart: number;
  switch (returnType) {
    case 1: // 日曜日始まり、1月1日を含む週が第1週
    case 17: // 日曜日始まり、1月1日を含む週が第1週（日曜日=1〜土曜日=7）
      weekStart = 0; // 日曜日
      break;
    case 2: // 月曜日始まり、1月1日を含む週が第1週
    case 21: // 月曜日始まり、1月1日を含む週が第1週（月曜日=1〜日曜日=7）
      weekStart = 1; // 月曜日
      break;
    default:
      weekStart = 0;
      break;
  }
  
  // 1月1日の曜日
  const jan1DayOfWeek = jan1.getDay();
  
  // 第1週の開始日を計算
  const firstWeekStart = new Date(jan1);
  const daysToFirstWeek = (weekStart - jan1DayOfWeek + 7) % 7;
  if (daysToFirstWeek > 0) {
    firstWeekStart.setDate(jan1.getDate() - daysToFirstWeek);
  }
  
  // 日付と第1週開始日の差から週番号を計算
  const diffTime = date.getTime() - firstWeekStart.getTime();
  const diffDays = Math.floor(diffTime / (24 * 60 * 60 * 1000));
  
  return Math.floor(diffDays / 7) + 1;
}

/**
 * ISO週番号を計算（ISOWEEKNUM関数用）
 */
export function getISOWeekNumber(date: Date): number {
  const year = date.getFullYear();
  const jan4 = new Date(year, 0, 4); // 1月4日
  
  // 1月4日の曜日を取得（月曜日=1, 日曜日=7）
  const jan4DayOfWeek = jan4.getDay() === 0 ? 7 : jan4.getDay();
  
  // ISO週の第1週の月曜日を計算
  const firstMondayOfYear = new Date(jan4);
  firstMondayOfYear.setDate(jan4.getDate() - jan4DayOfWeek + 1);
  
  // 指定日と第1週月曜日の差から週番号を計算
  const diffTime = date.getTime() - firstMondayOfYear.getTime();
  const diffDays = Math.floor(diffTime / (24 * 60 * 60 * 1000));
  
  const weekNumber = Math.floor(diffDays / 7) + 1;
  
  // 前年の最終週または来年の第1週の場合の調整
  if (weekNumber < 1) {
    // 前年の週番号を計算
    return getISOWeekNumber(new Date(year - 1, 11, 31));
  } else if (weekNumber > 52) {
    // 来年の第1週かチェック
    const dec31 = new Date(year, 11, 31);
    const dec31DayOfWeek = dec31.getDay() === 0 ? 7 : dec31.getDay();
    if (dec31DayOfWeek < 4) {
      return 1; // 来年の第1週
    }
  }
  
  return weekNumber;
}

/**
 * 時刻文字列を時間値に変換
 */
export function parseTimeString(timeStr: string): number | null {
  const timePattern = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/;
  const match = timeStr.match(timePattern);
  
  if (!match) {
    return null;
  }
  
  const hours = parseInt(match[1]);
  const minutes = parseInt(match[2]);
  const seconds = match[3] ? parseInt(match[3]) : 0;
  
  if (hours > 23 || minutes > 59 || seconds > 59) {
    return null;
  }
  
  // Excelと互換性を持たせるための正確な計算
  // 2:5:9 = 2 * 3600 + 5 * 60 + 9 = 7200 + 300 + 9 = 7509
  // 7509 / 86400 = 0.086921296296296296...
  
  // 高精度計算でExcelの結果と一致させる
  const h = hours / 24;
  const m = minutes / (24 * 60);
  const s = seconds / (24 * 60 * 60);
  
  return h + m + s;
}

/**
 * 指定されたタイムゾーンでの日付を取得
 */
export function getDateInTimezone(date: Date, timezone: string): Date {
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone: timezone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  });
  
  const parts = formatter.formatToParts(date);
  const initialPartsObj: Record<string, string> = {};
  const partsObj: Record<string, string> = parts.reduce((acc, part) => {
    acc[part.type] = part.value;
    return acc;
  }, initialPartsObj);
  
  return new Date(
    parseInt(partsObj.year),
    parseInt(partsObj.month) - 1,
    parseInt(partsObj.day),
    parseInt(partsObj.hour),
    parseInt(partsObj.minute),
    parseInt(partsObj.second)
  );
}

/**
 * 日付文字列やExcelシリアル値をDateオブジェクトに変換
 */
export function parseDate(input: unknown): Date | null {
  if (input === null || input === undefined) {
    return null;
  }
  
  // すでにDateオブジェクトの場合
  if (input instanceof Date) {
    return isValidDate(input) ? input : null;
  }
  
  // 数値の場合（Excelシリアル値）
  if (typeof input === 'number') {
    // 0も含めて適切に処理する
    return excelSerialToDate(input);
  }
  
  // 文字列の場合
  if (typeof input === 'string') {
    if (input === '') {
      return null;
    }
    return parseDateString(input);
  }
  
  return null;
}

/**
 * 文字列を日付に変換
 */
function parseDateString(dateStr: string): Date | null {
  // 空文字列や無効な文字列をチェック
  if (!dateStr || typeof dateStr !== 'string') {
    return null;
  }
  
  // YYYY-MM-DD形式を特別に処理（ローカル時間として解釈）
  const isoMatch = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    const [, year, month, day] = isoMatch;
    const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    if (isValidDate(date)) {
      return date;
    }
  }
  
  // その他のISO形式（時刻含む）を試す
  let date = new Date(dateStr);
  if (isValidDate(date)) {
    return date;
  }
  
  // YYYY/MM/DD形式を変換
  const yyyymmdd = dateStr.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (yyyymmdd) {
    const [, year, month, day] = yyyymmdd;
    date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    if (isValidDate(date)) {
      return date;
    }
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
 * Excelの1900年うるう年バグを考慮
 */
export function excelSerialToDate(serial: number): Date {
  // 0の場合は1899年12月30日を返す（Excelの仕様）
  if (serial === 0) {
    return new Date(1899, 11, 30);
  }
  
  // Excelは1900年を誤ってうるう年として扱うため、
  // 1900年3月1日以降の日付は1日ずれる
  let adjustedSerial = serial;
  
  // シリアル値60は存在しない1900年2月29日を表す
  // シリアル値61以降は実際の日付より1日進んでいる
  if (serial > 60) {
    adjustedSerial = serial - 1;
  }
  
  // Excel epoch: 1900/1/1 は実際は 1899/12/31
  const excelEpoch = new Date(1899, 11, 31); // 1899年12月31日
  const result = new Date(excelEpoch.getTime() + adjustedSerial * 24 * 60 * 60 * 1000);
  return result;
}

/**
 * DateオブジェクトをExcelシリアル値に変換
 * Excelの1900年うるう年バグを考慮
 */
export function dateToExcelSerial(date: Date): number {
  // ローカル日付を使用（UTCではなく）
  const year = date.getFullYear();
  const month = date.getMonth();
  const day = date.getDate();
  const dateOnly = new Date(year, month, day);
  
  // Excel's epoch is December 31, 1899 (serial number 1 = Jan 1, 1900)
  const excelEpoch = new Date(1899, 11, 31); // 1899年12月31日
  
  const diffTime = dateOnly.getTime() - excelEpoch.getTime();
  let serial = Math.round(diffTime / (24 * 60 * 60 * 1000));
  
  // 1900年2月28日より後は1日加算（Excelのうるう年バグを再現）
  // Excel incorrectly treats 1900 as a leap year
  if (serial > 59) {
    serial += 1;
  }
  
  return serial;
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
  const timezone = config?.timezone ?? globalDateConfig.timezone;
  
  if (timezone) {
    // タイムゾーンを考慮した現在時刻を取得
    const formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: timezone,
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      second: 'numeric',
      hour12: false
    });
    
    const parts = formatter.formatToParts(new Date());
    const dateParts: { [key: string]: string } = {};
    
    parts.forEach(part => {
      if (part.type !== 'literal') {
        dateParts[part.type] = part.value;
      }
    });
    
    return new Date(
      parseInt(dateParts.year),
      parseInt(dateParts.month) - 1,
      parseInt(dateParts.day),
      parseInt(dateParts.hour),
      parseInt(dateParts.minute),
      parseInt(dateParts.second)
    );
  }
  
  return new Date();
}

/**
 * 今日の日付を取得（時刻は00:00:00）
 */
export function today(config?: DateConfig): Date {
  const todayDate = now(config);
  return new Date(todayDate.getFullYear(), todayDate.getMonth(), todayDate.getDate());
}