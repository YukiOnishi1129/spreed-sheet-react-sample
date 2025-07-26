// Formatting utilities for spreadsheet cells

// Convert Excel serial number to Date
export function excelSerialToDate(serial: number): Date {
  // Excel's epoch is December 30, 1899
  const excelEpoch = new Date(1899, 11, 30);
  return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
}

// Check if a value might be a date serial number
export function isDateSerial(value: unknown): boolean {
  if (typeof value !== 'number') return false;
  // Excel date serials are typically between 1 (1900-01-01) and 60000+ (2164+)
  // TODAY() in 2025 would be around 45600-46000
  // Allow decimals for date-time values
  return value >= 1 && value <= 100000;
}

// Format a date serial number as YYYY/MM/DD
export function formatDateSerial(serial: number, includeTime: boolean = false): string {
  const date = excelSerialToDate(serial);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  if (includeTime && serial % 1 !== 0) {
    // Has time component
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    return `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
  }
  
  return `${year}/${month}/${day}`;
}

// Check if a string is a date in ISO format (YYYY-MM-DD)
export function isISODateString(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$/.test(value);
}

// Parse ISO date string to Date object
export function parseISODate(dateStr: string): Date | null {
  const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  
  const [, year, month, day] = match;
  return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
}

// Format a value for display (handles dates and other values)
export function formatCellValue(value: unknown, formula?: string): string | number {
  // If it's a date function result, format as date
  if (typeof value === 'number' && formula) {
    const upperFormula = formula.toUpperCase();
    const dateFunctions = ['TODAY()', 'NOW()', 'DATE(', 'EDATE(', 'EOMONTH(', 'WORKDAY('];
    
    // Check if the formula contains date functions
    const isDateFormula = dateFunctions.some(func => upperFormula.includes(func));
    
    if (isDateFormula && isDateSerial(value)) {
      // Check if it's NOW() which includes time
      const hasTime = upperFormula.includes('NOW()') || value % 1 !== 0;
      return formatDateSerial(value, hasTime);
    }
  }
  
  // Handle raw date strings (like "2024-07-15" in demo data)
  if (typeof value === 'string' && isISODateString(value)) {
    const date = parseISODate(value);
    if (date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}/${month}/${day}`;
    }
  }
  
  // Return the value as-is for non-date values
  if (typeof value === 'string' || typeof value === 'number' || value === null || value === undefined) {
    return value as string | number;
  }
  
  return String(value);
}