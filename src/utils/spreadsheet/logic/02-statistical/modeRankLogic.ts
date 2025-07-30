// モードとランク関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// MODE.SNGL関数の実装（最頻値・単一）
export const MODE_SNGL: CustomFormula = {
  name: 'MODE.SNGL',
  pattern: /MODE\.SNGL\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, args] = matches;
    
    try {
      const values: number[] = [];
      // 複数の範囲を正しく分割する
      const argList = args.match(/([A-Z]+\d+(?::[A-Z]+\d+)?|[^,]+)/g) || [];
      
      for (const arg of argList) {
        const trimmedArg = arg.trim();
        
        if (trimmedArg.includes(':')) {
          // 範囲の場合
          const rangeMatch = trimmedArg.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
          if (!rangeMatch) continue;
          
          const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
          const startRow = parseInt(startRowStr) - 1;
          const endRow = parseInt(endRowStr) - 1;
          const startColIndex = startCol.charCodeAt(0) - 65;
          const endColIndex = endCol.charCodeAt(0) - 65;
          
          for (let row = startRow; row <= endRow; row++) {
            for (let col = startColIndex; col <= endColIndex; col++) {
              const cell = context.data[row]?.[col];
              const cellValue = cell?.value !== undefined ? cell.value : cell;
              if (cellValue !== null && cellValue !== '' && !isNaN(Number(cellValue))) {
                values.push(Number(cellValue));
              }
            }
          }
        } else {
          // 単一セルまたは値の場合
          const value = getCellValue(trimmedArg, context);
          if (value !== null && !isNaN(Number(value))) {
            values.push(Number(value));
          }
        }
      }
      
      if (values.length === 0) {
        return FormulaError.NUM;
      }
      
      // 頻度をカウント
      const frequency: Map<number, number> = new Map();
      for (const value of values) {
        frequency.set(value, (frequency.get(value) ?? 0) + 1);
      }
      
      // 最頻値を見つける
      let mode: number | null = null;
      let maxFreq = 0;
      
      for (const [value, freq] of frequency) {
        if (freq > maxFreq) {
          maxFreq = freq;
          mode = value;
        }
      }
      
      // 単一値の場合はその値を返す
      if (values.length === 1) {
        return values[0];
      }
      
      if (maxFreq === 1) {
        return FormulaError.NA; // 複数値ですべてがユニークな値の場合
      }
      
      return mode ?? FormulaError.NA;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// MODE.MULT関数の実装（最頻値・複数）
export const MODE_MULT: CustomFormula = {
  name: 'MODE.MULT',
  pattern: /MODE\.MULT\((.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, args] = matches;
    
    try {
      const values: number[] = [];
      // 複数の範囲を正しく分割する
      const argList = args.match(/([A-Z]+\d+(?::[A-Z]+\d+)?|[^,]+)/g) || [];
      
      for (const arg of argList) {
        const trimmedArg = arg.trim();
        
        if (trimmedArg.includes(':')) {
          // 範囲の場合
          const rangeMatch = trimmedArg.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
          if (!rangeMatch) continue;
          
          const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
          const startRow = parseInt(startRowStr) - 1;
          const endRow = parseInt(endRowStr) - 1;
          const startColIndex = startCol.charCodeAt(0) - 65;
          const endColIndex = endCol.charCodeAt(0) - 65;
          
          for (let row = startRow; row <= endRow; row++) {
            for (let col = startColIndex; col <= endColIndex; col++) {
              const cell = context.data[row]?.[col];
              const cellValue = cell?.value !== undefined ? cell.value : cell;
              if (cellValue !== null && cellValue !== '' && !isNaN(Number(cellValue))) {
                values.push(Number(cellValue));
              }
            }
          }
        } else {
          // 単一セルまたは値の場合
          const value = getCellValue(trimmedArg, context);
          if (value !== null && !isNaN(Number(value))) {
            values.push(Number(value));
          }
        }
      }
      
      if (values.length === 0) {
        return FormulaError.NUM;
      }
      
      // 頻度をカウント
      const frequency: Map<number, number> = new Map();
      for (const value of values) {
        frequency.set(value, (frequency.get(value) ?? 0) + 1);
      }
      
      // 最大頻度を見つける
      let maxFreq = 0;
      for (const freq of frequency.values()) {
        if (freq > maxFreq) {
          maxFreq = freq;
        }
      }
      
      // 単一値の場合はその値の配列を返す
      if (values.length === 1) {
        return [values[0]];
      }
      
      if (maxFreq === 1) {
        return FormulaError.NA; // 複数値ですべてがユニークな場合
      }
      
      // 最頻値をすべて収集
      const modes: number[] = [];
      for (const [value, freq] of frequency) {
        if (freq === maxFreq) {
          modes.push(value);
        }
      }
      
      // 配列として返す
      return modes;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// RANK.AVG関数の実装（順位・平均）
export const RANK_AVG: CustomFormula = {
  name: 'RANK.AVG',
  pattern: /RANK\.AVG\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valueRef, arrayRef, orderRef] = matches;
    
    try {
      const value = parseFloat(getCellValue(valueRef.trim(), context)?.toString() ?? valueRef.trim());
      const order = orderRef ? parseInt(getCellValue(orderRef.trim(), context)?.toString() ?? orderRef.trim()) : 0;
      
      if (isNaN(value)) {
        return FormulaError.VALUE;
      }
      
      const values: number[] = [];
      
      if (arrayRef.includes(':')) {
        // 範囲の場合
        const rangeMatch = arrayRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (!rangeMatch) {
          return FormulaError.REF;
        }
        
        const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            const cell = context.data[row]?.[col];
            const cellValue = cell?.value !== undefined ? cell.value : cell;
            if (cellValue !== null && cellValue !== '' && !isNaN(Number(cellValue))) {
              values.push(Number(cellValue));
            }
          }
        }
      }
      
      if (values.length === 0) {
        return FormulaError.NUM;
      }
      
      if (!values.includes(value)) {
        return FormulaError.NA;
      }
      
      // 順位を計算
      const sortedValues = order === 0 ? 
        [...values].sort((a, b) => b - a) : 
        [...values].sort((a, b) => a - b);
      
      // 同じ値のインデックスをすべて見つける
      const indices: number[] = [];
      sortedValues.forEach((v, i) => {
        if (v === value) {
          indices.push(i + 1);
        }
      });
      
      // 平均順位を返す
      return indices.reduce((sum, idx) => sum + idx, 0) / indices.length;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// RANK.EQ関数の実装（順位・同順位）
export const RANK_EQ: CustomFormula = {
  name: 'RANK.EQ',
  pattern: /RANK\.EQ\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valueRef, arrayRef, orderRef] = matches;
    
    try {
      const value = parseFloat(getCellValue(valueRef.trim(), context)?.toString() ?? valueRef.trim());
      const order = orderRef ? parseInt(getCellValue(orderRef.trim(), context)?.toString() ?? orderRef.trim()) : 0;
      
      if (isNaN(value)) {
        return FormulaError.VALUE;
      }
      
      const values: number[] = [];
      
      if (arrayRef.includes(':')) {
        // 範囲の場合
        const rangeMatch = arrayRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (!rangeMatch) {
          return FormulaError.REF;
        }
        
        const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
        const startRow = parseInt(startRowStr) - 1;
        const endRow = parseInt(endRowStr) - 1;
        const startColIndex = startCol.charCodeAt(0) - 65;
        const endColIndex = endCol.charCodeAt(0) - 65;
        
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            const cell = context.data[row]?.[col];
            const cellValue = cell?.value !== undefined ? cell.value : cell;
            if (cellValue !== null && cellValue !== '' && !isNaN(Number(cellValue))) {
              values.push(Number(cellValue));
            }
          }
        }
      }
      
      if (values.length === 0) {
        return FormulaError.NUM;
      }
      
      if (!values.includes(value)) {
        return FormulaError.NA;
      }
      
      // 順位を計算
      const sortedValues = order === 0 ? 
        [...values].sort((a, b) => b - a) : 
        [...values].sort((a, b) => a - b);
      
      // 最初に見つかったインデックスを返す
      return sortedValues.indexOf(value) + 1;
    } catch {
      return FormulaError.VALUE;
    }
  }
};