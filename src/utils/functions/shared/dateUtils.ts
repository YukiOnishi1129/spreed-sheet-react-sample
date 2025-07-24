// 日付関連のユーティリティ関数

// Excel日付のシリアル値を計算（1900年1月1日を1とする）
export const getExcelDateSerial = (date: Date): number => {
  const excelEpoch = new Date(1900, 0, 1);
  const diffTime = date.getTime() - excelEpoch.getTime();
  const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24)) + 1;
  
  // Excel has a bug where it considers 1900 as a leap year
  if (diffDays > 59) {
    return diffDays + 1;
  }
  return diffDays;
};

// Excelのシリアル値から日付を取得
export const getDateFromExcelSerial = (serial: number): Date => {
  // Excel's epoch is January 1, 1900, but it incorrectly treats 1900 as a leap year
  // So we need to adjust for this
  let adjustedSerial = serial;
  if (serial > 60) {
    adjustedSerial = serial - 1;
  }
  
  const excelEpoch = new Date(1900, 0, 1);
  const msPerDay = 24 * 60 * 60 * 1000;
  return new Date(excelEpoch.getTime() + (adjustedSerial - 1) * msPerDay);
};

// 日付文字列を解析
export const parseDate = (dateStr: string): Date | null => {
  // Remove quotes if present
  const cleanStr = dateStr.replace(/['"]/g, '').trim();
  
  // Try various date formats
  // const formats = [
  //   // ISO format
  //   /^\d{4}-\d{2}-\d{2}$/,
  //   // US format
  //   /^\d{1,2}\/\d{1,2}\/\d{4}$/,
  //   // European format
  //   /^\d{1,2}\.\d{1,2}\.\d{4}$/,
  //   // Japanese format
  //   /^\d{4}\/\d{1,2}\/\d{1,2}$/,
  // ];
  
  // Check if it's a number (Excel serial date)
  const num = parseFloat(cleanStr);
  if (!isNaN(num) && num > 0 && num < 2958466) { // Excel max date
    return getDateFromExcelSerial(num);
  }
  
  // Try parsing as date string
  const date = new Date(cleanStr);
  if (!isNaN(date.getTime())) {
    return date;
  }
  
  // Try specific formats
  if (/^\d{4}-\d{2}-\d{2}$/.test(cleanStr)) {
    const [year, month, day] = cleanStr.split('-').map(Number);
    return new Date(year, month - 1, day);
  }
  
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(cleanStr)) {
    const [month, day, year] = cleanStr.split('/').map(Number);
    return new Date(year, month - 1, day);
  }
  
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(cleanStr)) {
    const [day, month, year] = cleanStr.split('.').map(Number);
    return new Date(year, month - 1, day);
  }
  
  if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(cleanStr)) {
    const [year, month, day] = cleanStr.split('/').map(Number);
    return new Date(year, month - 1, day);
  }
  
  return null;
};

// 時刻文字列を解析して0-1の範囲の値を返す
export const parseTime = (timeStr: string): number | null => {
  const cleanStr = timeStr.replace(/['"]/g, '').trim();
  
  // Check if it's already a decimal time
  const num = parseFloat(cleanStr);
  if (!isNaN(num) && num >= 0 && num < 1) {
    return num;
  }
  
  // Parse HH:MM:SS format
  const timeMatch = cleanStr.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (timeMatch) {
    const [, hours, minutes, seconds] = timeMatch;
    const h = parseInt(hours);
    const m = parseInt(minutes);
    const s = seconds ? parseInt(seconds) : 0;
    
    if (h >= 0 && h < 24 && m >= 0 && m < 60 && s >= 0 && s < 60) {
      return (h * 3600 + m * 60 + s) / 86400;
    }
  }
  
  return null;
};