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
      const url = getCellValue(urlRef.trim(), context)?.toString() ?? urlRef.trim();
      
      if (!url || url === '') {
        return FormulaError.VALUE;
      }
      
      // 実際のWebアクセスは実装せず、ダミーデータを返す
      // 本番実装では fetch API を使用
      const dummyData = [
        ['Product', 'Price', 'Stock'],
        ['Apple', '1.99', '100'],
        ['Banana', '0.99', '150'],
        ['Orange', '2.49', '80'],
      ];
      
      // 配列データを返す（実際にはスプレッドシートに展開される）
      return dummyData;
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
    const [, urlRef, queryRef, headersRef, numItemsRef] = matches;
    
    try {
      const url = getCellValue(urlRef.trim(), context)?.toString() ?? urlRef.trim();
      const query = queryRef ? getCellValue(queryRef.trim(), context)?.toString() ?? queryRef.trim() : 'items';
      const includeHeaders = headersRef ? getCellValue(headersRef.trim(), context)?.toString().toLowerCase() === 'true' : true;
      const numItems = numItemsRef ? parseInt(getCellValue(numItemsRef.trim(), context)?.toString() ?? numItemsRef.trim()) : 10;
      
      if (!url || url === '') {
        return FormulaError.VALUE;
      }
      
      if (isNaN(numItems) || numItems < 1) {
        return FormulaError.NUM;
      }
      
      // ダミーのRSSフィードデータ
      const feedData = [];
      
      if (includeHeaders) {
        feedData.push(['Title', 'Link', 'Date', 'Description']);
      }
      
      for (let i = 1; i <= Math.min(numItems, 5); i++) {
        feedData.push([
          `Article ${i}`,
          `https://example.com/article${i}`,
          new Date(Date.now() - i * 86400000).toISOString(),
          `This is the description for article ${i}`
        ]);
      }
      
      return feedData;
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
      const url = getCellValue(urlRef.trim(), context)?.toString() ?? urlRef.trim();
      const query = getCellValue(queryRef.trim(), context)?.toString() ?? queryRef.trim();
      const index = parseInt(getCellValue(indexRef.trim(), context)?.toString() ?? indexRef.trim());
      
      if (!url || url === '' || !query) {
        return FormulaError.VALUE;
      }
      
      if (isNaN(index) || index < 1) {
        return FormulaError.NUM;
      }
      
      const queryType = query.toLowerCase();
      if (queryType !== 'table' && queryType !== 'list') {
        return FormulaError.VALUE;
      }
      
      // ダミーのHTMLテーブルデータ
      if (queryType === 'table') {
        return [
          ['Header 1', 'Header 2', 'Header 3'],
          ['Row 1 Col 1', 'Row 1 Col 2', 'Row 1 Col 3'],
          ['Row 2 Col 1', 'Row 2 Col 2', 'Row 2 Col 3'],
          ['Row 3 Col 1', 'Row 3 Col 2', 'Row 3 Col 3'],
        ];
      } else {
        // リストデータ
        return [
          ['Item 1'],
          ['Item 2'],
          ['Item 3'],
          ['Item 4'],
          ['Item 5'],
        ];
      }
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
      const url = getCellValue(urlRef.trim(), context)?.toString() ?? urlRef.trim();
      const xpath = getCellValue(xpathRef.trim(), context)?.toString() ?? xpathRef.trim();
      
      if (!url || url === '' || !xpath || xpath === '') {
        return FormulaError.VALUE;
      }
      
      // ダミーのXMLデータ
      // 実際の実装では、XPath式に基づいてXMLを解析
      const xpathLower = xpath.toLowerCase();
      
      if (xpathLower.includes('title')) {
        return [
          ['Page Title 1'],
          ['Page Title 2'],
          ['Page Title 3'],
        ];
      } else if (xpathLower.includes('price')) {
        return [
          ['19.99'],
          ['29.99'],
          ['39.99'],
        ];
      } else {
        return [
          ['XML Data 1'],
          ['XML Data 2'],
          ['XML Data 3'],
        ];
      }
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
      const spreadsheetUrl = getCellValue(spreadsheetUrlRef.trim(), context)?.toString() ?? spreadsheetUrlRef.trim();
      const range = getCellValue(rangeRef.trim(), context)?.toString() ?? rangeRef.trim();
      
      if (!spreadsheetUrl || spreadsheetUrl === '' || !range || range === '') {
        return FormulaError.VALUE;
      }
      
      // 範囲の解析（例: "Sheet1!A1:C3"）
      const rangeMatch = range.match(/^(?:(.+)!)?([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
      if (!rangeMatch) {
        return FormulaError.REF;
      }
      
      const [, , startCol, startRowStr, endCol, endRowStr] = rangeMatch;
      const startRow = parseInt(startRowStr);
      const endRow = parseInt(endRowStr);
      const startColIndex = startCol.charCodeAt(0) - 65;
      const endColIndex = endCol.charCodeAt(0) - 65;
      
      // ダミーデータを生成
      const data = [];
      for (let row = 0; row <= endRow - startRow; row++) {
        const rowData = [];
        for (let col = 0; col <= endColIndex - startColIndex; col++) {
          rowData.push(`External ${String.fromCharCode(65 + startColIndex + col)}${startRow + row}`);
        }
        data.push(rowData);
      }
      
      return data;
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
      
      // 画像オブジェクトを返す（実際のレンダリングはUI側で処理）
      return {
        type: 'image',
        url: url,
        mode: mode,
        height: height,
        width: width
      };
    } catch {
      return FormulaError.VALUE;
    }
  }
};