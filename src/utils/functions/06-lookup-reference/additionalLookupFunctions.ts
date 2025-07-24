// 追加のルックアップ＆参照関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// ROW関数の実装（行番号を返す）
export const ROW: CustomFormula = {
  name: 'ROW',
  pattern: /ROW\((?:([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, cellRef] = matches;
    
    try {
      if (!cellRef) {
        // 引数なしの場合は現在のセルの行番号を返す
        return context.row + 1;
      }
      
      // セル参照の場合
      const cellMatch = cellRef.trim().match(/^([A-Z]+)(\d+)$/);
      if (cellMatch) {
        return parseInt(cellMatch[2]);
      }
      
      // 範囲参照の場合は最初のセルの行番号
      const rangeMatch = cellRef.trim().match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if (rangeMatch) {
        return parseInt(rangeMatch[2]);
      }
      
      return FormulaError.VALUE;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ROWS関数の実装（行数を返す）
export const ROWS: CustomFormula = {
  name: 'ROWS',
  pattern: /ROWS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, _context: FormulaContext): FormulaResult => {
    const [, rangeRef] = matches;
    
    try {
      const ref = rangeRef.trim();
      
      // 単一セルの場合
      const cellMatch = ref.match(/^([A-Z]+)(\d+)$/);
      if (cellMatch) {
        return 1;
      }
      
      // 範囲の場合
      const rangeMatch = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if (rangeMatch) {
        const startRow = parseInt(rangeMatch[2]);
        const endRow = parseInt(rangeMatch[4]);
        return Math.abs(endRow - startRow) + 1;
      }
      
      return FormulaError.VALUE;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// COLUMN関数の実装（列番号を返す）
export const COLUMN: CustomFormula = {
  name: 'COLUMN',
  pattern: /COLUMN\((?:([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, cellRef] = matches;
    
    try {
      if (!cellRef) {
        // 引数なしの場合は現在のセルの列番号を返す
        return context.col + 1;
      }
      
      // セル参照の場合
      const cellMatch = cellRef.trim().match(/^([A-Z]+)(\d+)$/);
      if (cellMatch) {
        const colLetters = cellMatch[1];
        let colNum = 0;
        for (let i = 0; i < colLetters.length; i++) {
          colNum = colNum * 26 + (colLetters.charCodeAt(i) - 64);
        }
        return colNum;
      }
      
      // 範囲参照の場合は最初のセルの列番号
      const rangeMatch = cellRef.trim().match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if (rangeMatch) {
        const colLetters = rangeMatch[1];
        let colNum = 0;
        for (let i = 0; i < colLetters.length; i++) {
          colNum = colNum * 26 + (colLetters.charCodeAt(i) - 64);
        }
        return colNum;
      }
      
      return FormulaError.VALUE;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// COLUMNS関数の実装（列数を返す）
export const COLUMNS: CustomFormula = {
  name: 'COLUMNS',
  pattern: /COLUMNS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, _context: FormulaContext): FormulaResult => {
    const [, rangeRef] = matches;
    
    try {
      const ref = rangeRef.trim();
      
      // 単一セルの場合
      const cellMatch = ref.match(/^([A-Z]+)(\d+)$/);
      if (cellMatch) {
        return 1;
      }
      
      // 範囲の場合
      const rangeMatch = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if (rangeMatch) {
        const startCol = rangeMatch[1];
        const endCol = rangeMatch[3];
        
        let startColNum = 0;
        for (let i = 0; i < startCol.length; i++) {
          startColNum = startColNum * 26 + (startCol.charCodeAt(i) - 64);
        }
        
        let endColNum = 0;
        for (let i = 0; i < endCol.length; i++) {
          endColNum = endColNum * 26 + (endCol.charCodeAt(i) - 64);
        }
        
        return Math.abs(endColNum - startColNum) + 1;
      }
      
      return FormulaError.VALUE;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ADDRESS関数の実装（セルアドレスを作成）
export const ADDRESS: CustomFormula = {
  name: 'ADDRESS',
  pattern: /ADDRESS\(([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rowRef, colRef, absRef, a1Ref, sheetRef] = matches;
    
    try {
      const row = parseInt(getCellValue(rowRef.trim(), context)?.toString() ?? rowRef.trim());
      const col = parseInt(getCellValue(colRef.trim(), context)?.toString() ?? colRef.trim());
      const absNum = absRef ? parseInt(getCellValue(absRef.trim(), context)?.toString() ?? absRef.trim()) : 1;
      const useA1 = a1Ref ? getCellValue(a1Ref.trim(), context)?.toString().toLowerCase() !== 'false' : true;
      let sheetName = sheetRef ? getCellValue(sheetRef.trim(), context)?.toString() ?? '' : '';
      
      if (isNaN(row) || isNaN(col) || row < 1 || col < 1) {
        return FormulaError.VALUE;
      }
      
      // 列番号を文字に変換
      let colLetter = '';
      let colNum = col;
      while (colNum > 0) {
        colNum--;
        colLetter = String.fromCharCode(65 + (colNum % 26)) + colLetter;
        colNum = Math.floor(colNum / 26);
      }
      
      let address = '';
      
      if (useA1) {
        // A1形式
        switch (absNum) {
          case 1: // 絶対参照
            address = `$${colLetter}$${row}`;
            break;
          case 2: // 行は相対、列は絶対
            address = `$${colLetter}${row}`;
            break;
          case 3: // 行は絶対、列は相対
            address = `${colLetter}$${row}`;
            break;
          case 4: // 相対参照
            address = `${colLetter}${row}`;
            break;
          default:
            return FormulaError.VALUE;
        }
      } else {
        // R1C1形式
        switch (absNum) {
          case 1: // 絶対参照
            address = `R${row}C${col}`;
            break;
          case 2: // 行は相対、列は絶対
            address = `R[${row - context.row - 1}]C${col}`;
            break;
          case 3: // 行は絶対、列は相対
            address = `R${row}C[${col - context.col - 1}]`;
            break;
          case 4: // 相対参照
            address = `R[${row - context.row - 1}]C[${col - context.col - 1}]`;
            break;
          default:
            return FormulaError.VALUE;
        }
      }
      
      // シート名を処理
      if (sheetName) {
        if (sheetName.startsWith('"') && sheetName.endsWith('"')) {
          sheetName = sheetName.slice(1, -1);
        }
        // シート名にスペースや特殊文字が含まれる場合は引用符で囲む
        if (sheetName.match(/[\s!'"]/)) {
          sheetName = `'${sheetName.replace(/'/g, "''")}'`;
        }
        address = `${sheetName}!${address}`;
      }
      
      return address;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// AREAS関数の実装（領域数を返す）
export const AREAS: CustomFormula = {
  name: 'AREAS',
  pattern: /AREAS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, _context: FormulaContext): FormulaResult => {
    const [, reference] = matches;
    
    try {
      // 簡易実装：カンマで区切られた参照の数をカウント
      const refs = reference.split(',').map(ref => ref.trim()).filter(ref => ref.length > 0);
      
      // 各参照が有効かチェック
      for (const ref of refs) {
        if (!ref.match(/^[A-Z]+\d+(?::[A-Z]+\d+)?$/)) {
          return FormulaError.VALUE;
        }
      }
      
      return refs.length;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// FORMULATEXT関数の実装（数式を文字列で返す）
export const FORMULATEXT: CustomFormula = {
  name: 'FORMULATEXT',
  pattern: /FORMULATEXT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, cellRef] = matches;
    
    try {
      const ref = cellRef.trim();
      const cellMatch = ref.match(/^([A-Z]+)(\d+)$/);
      
      if (!cellMatch) {
        return FormulaError.VALUE;
      }
      
      const col = cellMatch[1].charCodeAt(0) - 65;
      const row = parseInt(cellMatch[2]) - 1;
      
      if (row < 0 || row >= context.data.length || col < 0 || col >= (context.data[0]?.length ?? 0)) {
        return FormulaError.REF;
      }
      
      const cell = context.data[row]?.[col];
      
      if (!cell || !cell.formula) {
        return FormulaError.NA;
      }
      
      return cell.formula;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// XMATCH関数の実装（拡張版MATCH）
export const XMATCH: CustomFormula = {
  name: 'XMATCH',
  pattern: /XMATCH\(([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, lookupRef, arrayRef, matchModeRef, searchModeRef] = matches;
    
    try {
      const lookupValue = getCellValue(lookupRef.trim(), context) ?? lookupRef.trim();
      const matchMode = matchModeRef ? parseInt(getCellValue(matchModeRef.trim(), context)?.toString() ?? matchModeRef.trim()) : 0;
      const searchMode = searchModeRef ? parseInt(getCellValue(searchModeRef.trim(), context)?.toString() ?? searchModeRef.trim()) : 1;
      
      // 配列を取得
      const values: unknown[] = [];
      
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
        
        // 1次元配列として扱う
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColIndex; col <= endColIndex; col++) {
            if (context.data[row]?.[col]) {
              values.push(context.data[row][col].value);
            }
          }
        }
      } else {
        // 単一セルの場合
        values.push(getCellValue(arrayRef.trim(), context));
      }
      
      if (values.length === 0) {
        return FormulaError.NA;
      }
      
      // 検索モード別の処理
      let searchArray = [...values];
      // let startIndex = 0;
      
      switch (searchMode) {
        case 1: // 最初から最後へ（デフォルト）
          break;
        case -1: // 最後から最初へ
          searchArray = searchArray.reverse();
          break;
        case 2: // バイナリ検索（昇順）
        case -2: // バイナリ検索（降順）
          // 簡易実装：通常の検索を行う
          break;
        default:
          return FormulaError.VALUE;
      }
      
      // マッチモード別の検索
      let matchIndex = -1;
      
      switch (matchMode) {
        case 0: // 完全一致（デフォルト）
          for (let i = 0; i < searchArray.length; i++) {
            if (String(searchArray[i]) === String(lookupValue)) {
              matchIndex = i;
              break;
            }
          }
          break;
          
        case -1: // 完全一致または次に小さい値
          for (let i = searchArray.length - 1; i >= 0; i--) {
            const val = Number(searchArray[i]);
            const lookup = Number(lookupValue);
            if (!isNaN(val) && !isNaN(lookup) && val <= lookup) {
              matchIndex = i;
              break;
            }
          }
          break;
          
        case 1: // 完全一致または次に大きい値
          for (let i = 0; i < searchArray.length; i++) {
            const val = Number(searchArray[i]);
            const lookup = Number(lookupValue);
            if (!isNaN(val) && !isNaN(lookup) && val >= lookup) {
              matchIndex = i;
              break;
            }
          }
          break;
          
        case 2: // ワイルドカード一致
          const pattern = String(lookupValue)
            .replace(/\*/g, '.*')
            .replace(/\?/g, '.')
            .replace(/~/g, '\\');
          const regex = new RegExp(`^${pattern}$`, 'i');
          
          for (let i = 0; i < searchArray.length; i++) {
            if (regex.test(String(searchArray[i]))) {
              matchIndex = i;
              break;
            }
          }
          break;
          
        default:
          return FormulaError.VALUE;
      }
      
      if (matchIndex === -1) {
        return FormulaError.NA;
      }
      
      // 検索モードが逆順の場合、インデックスを調整
      if (searchMode === -1) {
        matchIndex = values.length - matchIndex;
      } else {
        matchIndex = matchIndex + 1; // 1ベースのインデックス
      }
      
      return matchIndex;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// LAMBDA関数の実装（カスタム関数を作成）
export const LAMBDA: CustomFormula = {
  name: 'LAMBDA',
  pattern: /LAMBDA\((.+)\)/i,
  calculate: (_matches: RegExpMatchArray, _context: FormulaContext): FormulaResult => {
    // LAMBDAは通常、名前定義で使用され、直接セルで使用されることは少ない
    // ここでは簡易的にエラーを返す
    return FormulaError.VALUE;
  }
};

// HYPERLINK関数の実装（ハイパーリンクを作成）
export const HYPERLINK: CustomFormula = {
  name: 'HYPERLINK',
  pattern: /HYPERLINK\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, linkRef, friendlyRef] = matches;
    
    try {
      let link = getCellValue(linkRef.trim(), context)?.toString() ?? linkRef.trim();
      let friendlyName = friendlyRef ? 
        getCellValue(friendlyRef.trim(), context)?.toString() ?? friendlyRef.trim() : 
        link;
      
      // 引用符を除去
      if (link.startsWith('"') && link.endsWith('"')) {
        link = link.slice(1, -1);
      }
      if (friendlyName.startsWith('"') && friendlyName.endsWith('"')) {
        friendlyName = friendlyName.slice(1, -1);
      }
      
      // 表示テキストを返す（実際のリンク機能はUIで処理）
      return friendlyName;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CHOOSEROWS関数の実装（行を選択）
export const CHOOSEROWS: CustomFormula = {
  name: 'CHOOSEROWS',
  pattern: /CHOOSEROWS\(([^,]+),\s*(.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, rowNumsStr] = matches;
    
    try {
      // 配列を取得
      const rangeMatch = arrayRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (!rangeMatch) {
        return FormulaError.REF;
      }
      
      const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
      const startRow = parseInt(startRowStr) - 1;
      const endRow = parseInt(endRowStr) - 1;
      const startColIndex = startCol.charCodeAt(0) - 65;
      const endColIndex = endCol.charCodeAt(0) - 65;
      
      // 選択する行番号を解析
      const rowNums = rowNumsStr.split(',').map(num => {
        const n = parseInt(getCellValue(num.trim(), context)?.toString() ?? num.trim());
        return isNaN(n) ? 0 : n;
      });
      
      // 結果配列を構築
      const result: unknown[][] = [];
      
      for (const rowNum of rowNums) {
        if (rowNum < 1 || rowNum > (endRow - startRow + 1)) {
          return FormulaError.VALUE;
        }
        
        const sourceRow = startRow + rowNum - 1;
        const row: unknown[] = [];
        
        for (let col = startColIndex; col <= endColIndex; col++) {
          if (context.data[sourceRow]?.[col]) {
            row.push(context.data[sourceRow][col].value);
          } else {
            row.push(null);
          }
        }
        
        result.push(row);
      }
      
      return result as (string | number)[][];
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CHOOSECOLS関数の実装（列を選択）
export const CHOOSECOLS: CustomFormula = {
  name: 'CHOOSECOLS',
  pattern: /CHOOSECOLS\(([^,]+),\s*(.+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, arrayRef, colNumsStr] = matches;
    
    try {
      // 配列を取得
      const rangeMatch = arrayRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (!rangeMatch) {
        return FormulaError.REF;
      }
      
      const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
      const startRow = parseInt(startRowStr) - 1;
      const endRow = parseInt(endRowStr) - 1;
      const startColIndex = startCol.charCodeAt(0) - 65;
      const endColIndex = endCol.charCodeAt(0) - 65;
      
      // 選択する列番号を解析
      const colNums = colNumsStr.split(',').map(num => {
        const n = parseInt(getCellValue(num.trim(), context)?.toString() ?? num.trim());
        return isNaN(n) ? 0 : n;
      });
      
      // 結果配列を構築
      const result: unknown[][] = [];
      
      for (let row = startRow; row <= endRow; row++) {
        const resultRow: unknown[] = [];
        
        for (const colNum of colNums) {
          if (colNum < 1 || colNum > (endColIndex - startColIndex + 1)) {
            return FormulaError.VALUE;
          }
          
          const sourceCol = startColIndex + colNum - 1;
          
          if (context.data[row]?.[sourceCol]) {
            resultRow.push(context.data[row][sourceCol].value);
          } else {
            resultRow.push(null);
          }
        }
        
        result.push(resultRow);
      }
      
      return result as (string | number)[][];
    } catch {
      return FormulaError.VALUE;
    }
  }
};