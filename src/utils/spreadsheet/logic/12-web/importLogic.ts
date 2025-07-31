// Web インポート関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// IMPORTDATA関数の実装（CSVやTSVをインポート）
export const IMPORTDATA: CustomFormula = {
  name: 'IMPORTDATA',
  pattern: /IMPORTDATA\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, urlRef] = matches;
    
    try {
      let url = getCellValue(urlRef.trim(), context)?.toString() ?? urlRef.trim();
      
      // 引用符を除去
      if (url.startsWith('"') && url.endsWith('"')) {
        url = url.slice(1, -1);
      }
      
      if (!url || url === '') {
        return FormulaError.VALUE;
      }
      
      // 実際のWebアクセスは実装せず、エラーメッセージを返す
      // 本番実装では fetch API を使用
      return '#N/A - Web import functions require external data access';
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMPORTFEED関数の実装（RSSやAtomフィードを取得）
export const IMPORTFEED: CustomFormula = {
  name: 'IMPORTFEED',
  pattern: /IMPORTFEED\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, urlRef] = matches;
    
    try {
      let url = getCellValue(urlRef?.trim(), context)?.toString() ?? urlRef?.trim() ?? '';
      
      // 引用符を除去
      if (url.startsWith('"') && url.endsWith('"')) {
        url = url.slice(1, -1);
      }
      
      if (!url || url === '') {
        return FormulaError.VALUE;
      }
      
      // 実際のWebアクセスは実装せず、エラーメッセージを返す
      return '#N/A - Web import functions require external data access';
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMPORTHTML関数の実装（HTMLテーブルやリストを取得）
export const IMPORTHTML: CustomFormula = {
  name: 'IMPORTHTML',
  pattern: /IMPORTHTML\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, urlRef, queryRef, indexRef] = matches;
    
    try {
      let url = getCellValue(urlRef.trim(), context)?.toString() ?? urlRef.trim();
      let query = getCellValue(queryRef.trim(), context)?.toString() ?? queryRef.trim();
      const indexValue = getCellValue(indexRef.trim(), context)?.toString() ?? indexRef.trim();
      
      // 引用符を除去
      if (url.startsWith('"') && url.endsWith('"')) {
        url = url.slice(1, -1);
      }
      if (query.startsWith('"') && query.endsWith('"')) {
        query = query.slice(1, -1);
      }
      
      const index = parseInt(indexValue);
      
      if (!url || url === '' || !query) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(index) || index < 0) {
        return FormulaError.NUM;
      }
      
      const queryType = query.toLowerCase();
      if (queryType !== 'table' && queryType !== 'list') {
        return FormulaError.VALUE;
      }
      
      // 実際のWebアクセスは実装せず、エラーメッセージを返す
      return '#N/A - Web import functions require external data access';
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMPORTXML関数の実装（XMLデータを取得）
export const IMPORTXML: CustomFormula = {
  name: 'IMPORTXML',
  pattern: /IMPORTXML\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, urlRef, xpathRef] = matches;
    
    try {
      let url = getCellValue(urlRef.trim(), context)?.toString() ?? urlRef.trim();
      let xpath = getCellValue(xpathRef.trim(), context)?.toString() ?? xpathRef.trim();
      
      // 引用符を除去
      if (url.startsWith('"') && url.endsWith('"')) {
        url = url.slice(1, -1);
      }
      if (xpath.startsWith('"') && xpath.endsWith('"')) {
        xpath = xpath.slice(1, -1);
      }
      
      if (!url || url === '' || !xpath || xpath === '') {
        return FormulaError.VALUE;
      }
      
      // 実際のWebアクセスは実装せず、エラーメッセージを返す
      return '#N/A - Web import functions require external data access';
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMPORTRANGE関数の実装（他のスプレッドシートから取得）
export const IMPORTRANGE: CustomFormula = {
  name: 'IMPORTRANGE',
  pattern: /IMPORTRANGE\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, spreadsheetUrlRef, rangeRef] = matches;
    
    try {
      let spreadsheetUrl = getCellValue(spreadsheetUrlRef.trim(), context)?.toString() ?? spreadsheetUrlRef.trim();
      let range = getCellValue(rangeRef.trim(), context)?.toString() ?? rangeRef.trim();
      
      // 引用符を除去
      if (spreadsheetUrl.startsWith('"') && spreadsheetUrl.endsWith('"')) {
        spreadsheetUrl = spreadsheetUrl.slice(1, -1);
      }
      if (range.startsWith('"') && range.endsWith('"')) {
        range = range.slice(1, -1);
      }
      
      if (!spreadsheetUrl || spreadsheetUrl === '' || !range || range === '') {
        return FormulaError.VALUE;
      }
      
      // 実際のWebアクセスは実装せず、エラーメッセージを返す
      return '#N/A - Web import functions require external data access';
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// IMAGE関数の実装（画像を挿入）
export const IMAGE: CustomFormula = {
  name: 'IMAGE',
  pattern: /IMAGE\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, urlRef, modeRef, heightRef, widthRef] = matches;
    
    try {
      const url = getCellValue(urlRef.trim(), context)?.toString() ?? urlRef.trim();
      const mode = modeRef ? parseInt(getCellValue(modeRef.trim(), context)?.toString() ?? modeRef.trim()) : 1;
      const height = heightRef ? parseInt(getCellValue(heightRef.trim(), context)?.toString() ?? heightRef.trim()) : null;
      const width = widthRef ? parseInt(getCellValue(widthRef.trim(), context)?.toString() ?? widthRef.trim()) : null;
      
      if (!url || url === '') {
        return FormulaError.VALUE;
      }
      
      if (mode < 1 || mode > 4) {
        return FormulaError.NUM;
      }
      
      // モード3または4の場合、高さと幅が必要
      if ((mode === 3 || mode === 4) && (!height || !width || height <= 0 || width <= 0)) {
        return FormulaError.NUM;
      }
      
      // 画像URLを返す（実際のレンダリングはUI側で処理）
      // 今回の実装では単純にURLを返す
      return url;
    } catch {
      return FormulaError.VALUE;
    }
  }
};