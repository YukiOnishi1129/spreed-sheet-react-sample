// Google Sheets追加関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// SORTN関数の実装（上位N件をソート）
export const SORTN: CustomFormula = {
  name: 'SORTN',
  pattern: /SORTN\(([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, rangeRef, nRef, displayTiesRef, sortColumnRef, isAscendingRef] = matches;
    
    try {
      // データ配列を取得
      const data: (number | string | boolean | null)[][] = [];
      const rangeMatch = rangeRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      
      if (!rangeMatch) {
        return FormulaError.REF;
      }
      
      const [, startCol, startRowStr, endCol, endRowStr] = rangeMatch;
      const startRow = parseInt(startRowStr) - 1;
      const endRow = parseInt(endRowStr) - 1;
      const startColIndex = startCol.charCodeAt(0) - 65;
      const endColIndex = endCol.charCodeAt(0) - 65;
      
      for (let row = startRow; row <= endRow; row++) {
        const rowData: (number | string | boolean | null)[] = [];
        for (let col = startColIndex; col <= endColIndex; col++) {
          const cell = context.data[row]?.[col];
          rowData.push(cell ? cell.value as (number | string | boolean | null) : null);
        }
        data.push(rowData);
      }
      
      const n = parseInt(getCellValue(nRef.trim(), context)?.toString() ?? nRef.trim());
      const displayTies = displayTiesRef ? parseInt(getCellValue(displayTiesRef.trim(), context)?.toString() ?? displayTiesRef.trim()) : 0;
      const sortColumn = sortColumnRef ? parseInt(getCellValue(sortColumnRef.trim(), context)?.toString() ?? sortColumnRef.trim()) - 1 : 0;
      const isAscending = isAscendingRef ? getCellValue(isAscendingRef.trim(), context)?.toString().toLowerCase() === 'true' : false;
      
      if (isNaN(n) || n <= 0) {
        return FormulaError.VALUE;
      }
      
      // ソート
      const sortedData = [...data].sort((a, b) => {
        const valA = a[sortColumn];
        const valB = b[sortColumn];
        
        if (valA === null || valA === undefined) return 1;
        if (valB === null || valB === undefined) return -1;
        
        let comparison = 0;
        if (typeof valA === 'number' && typeof valB === 'number') {
          comparison = valA - valB;
        } else {
          comparison = String(valA).localeCompare(String(valB));
        }
        
        return isAscending ? comparison : -comparison;
      });
      
      // 上位N件を取得
      const result = sortedData.slice(0, n);
      
      // タイの処理
      if (displayTies === 1 && n < sortedData.length) {
        // 最後の値と同じ値を持つ行を追加
        const lastValue = result[result.length - 1][sortColumn];
        for (let i = n; i < sortedData.length; i++) {
          if (sortedData[i][sortColumn] === lastValue) {
            result.push(sortedData[i]);
          } else {
            break;
          }
        }
      }
      
      return result;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// SPARKLINE関数の実装（スパークライン）
export const SPARKLINE: CustomFormula = {
  name: 'SPARKLINE',
  pattern: /SPARKLINE\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, dataRef, optionsRef] = matches;
    
    try {
      // データを取得
      const values: number[] = [];
      
      if (dataRef.includes(':')) {
        const rangeMatch = dataRef.trim().match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
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
            if (cell && !isNaN(Number(cell.value))) {
              values.push(Number(cell.value));
            }
          }
        }
      } else {
        const value = getCellValue(dataRef.trim(), context);
        if (value !== null && !isNaN(Number(value))) {
          values.push(Number(value));
        }
      }
      
      if (values.length === 0) {
        return '';
      }
      
      // オプションの解析（簡易実装）
      let chartType = 'line';
      if (optionsRef?.includes('charttype')) {
        if (optionsRef.includes('bar')) chartType = 'bar';
        else if (optionsRef.includes('column')) chartType = 'column';
        else if (optionsRef.includes('winloss')) chartType = 'winloss';
      }
      
      // 簡易的なテキストベースのスパークライン
      const min = Math.min(...values);
      const max = Math.max(...values);
      const range = max - min || 1;
      
      if (chartType === 'bar' || chartType === 'column') {
        // バーチャート風
        const bars = ['▁', '▂', '▃', '▄', '▅', '▆', '▇', '█'];
        return values.map(v => {
          const normalized = (v - min) / range;
          const index = Math.min(Math.floor(normalized * bars.length), bars.length - 1);
          return bars[Math.max(0, index)];
        }).join('');
      } else if (chartType === 'winloss') {
        // Win/Loss
        return values.map(v => v > 0 ? '▲' : v < 0 ? '▼' : '─').join('');
      } else {
        // ラインチャート風
        const sparkChars = ['⎯', '⎼', '─', '╌', '┄', '┈', '⋯'];
        return values.map(v => {
          const normalized = (v - min) / range;
          const index = Math.min(Math.floor(normalized * sparkChars.length), sparkChars.length - 1);
          return sparkChars[Math.max(0, index)];
        }).join('');
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// GOOGLETRANSLATE関数の実装（翻訳）
export const GOOGLETRANSLATE: CustomFormula = {
  name: 'GOOGLETRANSLATE',
  pattern: /GOOGLETRANSLATE\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, textRef, sourceLangRef, targetLangRef] = matches;
    
    try {
      const text = getCellValue(textRef.trim(), context)?.toString() ?? textRef.trim();
      const sourceLang = getCellValue(sourceLangRef.trim(), context)?.toString() ?? sourceLangRef.trim();
      const targetLang = getCellValue(targetLangRef.trim(), context)?.toString() ?? targetLangRef.trim();
      
      // 引用符を除去
      const cleanText = text.replace(/^["']|["']$/g, '');
      const cleanSourceLang = sourceLang.replace(/^["']|["']$/g, '').toLowerCase();
      const cleanTargetLang = targetLang.replace(/^["']|["']$/g, '').toLowerCase();
      
      // 実際の翻訳APIは使用せず、シミュレーション
      // 言語コードの検証
      const validLangs = ['en', 'ja', 'es', 'fr', 'de', 'it', 'pt', 'ru', 'zh', 'ko', 'auto'];
      
      if (!validLangs.includes(cleanSourceLang) || !validLangs.includes(cleanTargetLang)) {
        return FormulaError.VALUE;
      }
      
      // シンプルな翻訳シミュレーション
      if (cleanSourceLang === cleanTargetLang && cleanSourceLang !== 'auto') {
        return cleanText;
      }
      
      // デモ用の簡単な翻訳例
      const translations: { [key: string]: { [lang: string]: string } } = {
        'hello': { 'ja': 'こんにちは', 'es': 'hola', 'fr': 'bonjour', 'de': 'hallo' },
        'goodbye': { 'ja': 'さようなら', 'es': 'adiós', 'fr': 'au revoir', 'de': 'auf wiedersehen' },
        'thank you': { 'ja': 'ありがとう', 'es': 'gracias', 'fr': 'merci', 'de': 'danke' },
      };
      
      const lowerText = cleanText.toLowerCase();
      if (translations[lowerText]?.[cleanTargetLang]) {
        return translations[lowerText][cleanTargetLang];
      }
      
      // 翻訳できない場合は、[Translated]プレフィックスを付けて返す
      return `[${cleanTargetLang.toUpperCase()}] ${cleanText}`;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// DETECTLANGUAGE関数の実装（言語検出）
export const DETECTLANGUAGE: CustomFormula = {
  name: 'DETECTLANGUAGE',
  pattern: /DETECTLANGUAGE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, textRef] = matches;
    
    try {
      const text = getCellValue(textRef.trim(), context)?.toString() ?? textRef.trim();
      const cleanText = text.replace(/^["']|["']$/g, '');
      
      // 簡易的な言語検出
      // 実際のAPIは使用せず、文字種別で判定
      
      // 日本語の検出
      if (/[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/.test(cleanText)) {
        return 'ja';
      }
      
      // 韓国語の検出
      if (/[\uAC00-\uD7AF]/.test(cleanText)) {
        return 'ko';
      }
      
      // 中国語の検出（簡体字・繁体字）
      if (/[\u4E00-\u9FFF]/.test(cleanText) && !/[\u3040-\u309F\u30A0-\u30FF]/.test(cleanText)) {
        return 'zh';
      }
      
      // アラビア語の検出
      if (/[\u0600-\u06FF]/.test(cleanText)) {
        return 'ar';
      }
      
      // キリル文字（ロシア語等）の検出
      if (/[\u0400-\u04FF]/.test(cleanText)) {
        return 'ru';
      }
      
      // ギリシャ文字の検出
      if (/[\u0370-\u03FF]/.test(cleanText)) {
        return 'el';
      }
      
      // ヘブライ語の検出
      if (/[\u0590-\u05FF]/.test(cleanText)) {
        return 'he';
      }
      
      // デフォルトは英語
      return 'en';
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// GOOGLEFINANCE関数の実装（金融情報）
export const GOOGLEFINANCE: CustomFormula = {
  name: 'GOOGLEFINANCE',
  pattern: /GOOGLEFINANCE\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, tickerRef, attributeRef, startDateRef, endDateRef] = matches;
    
    try {
      const ticker = (getCellValue(tickerRef.trim(), context)?.toString() ?? tickerRef.trim()).replace(/^["']|["']$/g, '');
      const attribute = attributeRef ? 
        (getCellValue(attributeRef.trim(), context)?.toString() ?? attributeRef.trim()).replace(/^["']|["']$/g, '').toLowerCase() : 
        'price';
      
      // 実際のAPIは使用せず、ダミーデータを返す
      const dummyData: { [key: string]: { [attr: string]: number } } = {
        'AAPL': { price: 178.50, volume: 75234000, pe: 28.5, marketcap: 2800000000000 },
        'GOOGL': { price: 142.30, volume: 25678000, pe: 25.2, marketcap: 1800000000000 },
        'MSFT': { price: 380.20, volume: 22345000, pe: 32.1, marketcap: 2900000000000 },
        'AMZN': { price: 155.40, volume: 45678000, pe: 58.3, marketcap: 1600000000000 },
        'TSLA': { price: 240.80, volume: 98765000, pe: 71.2, marketcap: 760000000000 },
      };
      
      const upperTicker = ticker.toUpperCase();
      
      if (!dummyData[upperTicker]) {
        return FormulaError.NA;
      }
      
      const validAttributes = ['price', 'volume', 'pe', 'marketcap', 'high', 'low', 'open', 'close'];
      
      if (!validAttributes.includes(attribute)) {
        return FormulaError.VALUE;
      }
      
      // 日付範囲が指定されている場合は配列を返す（簡易実装）
      if (startDateRef && endDateRef) {
        // ヘッダー行と数日分のダミーデータを返す
        const result: (number | string | boolean | null)[][] = [['Date', attribute.charAt(0).toUpperCase() + attribute.slice(1)]];
        const baseValue = dummyData[upperTicker][attribute] || dummyData[upperTicker].price;
        
        for (let i = 0; i < 5; i++) {
          const date = new Date();
          date.setDate(date.getDate() - i);
          const variation = 1 + (Math.random() - 0.5) * 0.1; // ±5%の変動
          result.push([
            date.toISOString().split('T')[0],
            Math.round(baseValue * variation * 100) / 100
          ]);
        }
        
        return result;
      }
      
      // 単一の値を返す
      return dummyData[upperTicker][attribute] || dummyData[upperTicker].price;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TO_DATE関数の実装（日付変換）
export const TO_DATE: CustomFormula = {
  name: 'TO_DATE',
  pattern: /TO_DATE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valueRef] = matches;
    
    try {
      const value = getCellValue(valueRef.trim(), context);
      
      if (value === null || value === undefined) {
        return FormulaError.VALUE;
      }
      
      // 数値の場合はExcelシリアル値として扱う
      if (typeof value === 'number') {
        const excelEpoch = new Date(1900, 0, 1);
        const msPerDay = 24 * 60 * 60 * 1000;
        
        // Excelの1900年のうるう年バグを考慮
        let adjustedValue = value;
        if (value > 60) {
          adjustedValue = value - 1;
        }
        
        const date = new Date(excelEpoch.getTime() + (adjustedValue - 1) * msPerDay);
        return date.toLocaleDateString();
      }
      
      // 文字列の場合は日付として解析
      const date = new Date(String(value));
      if (isNaN(date.getTime())) {
        return FormulaError.VALUE;
      }
      
      return date.toLocaleDateString();
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TO_PERCENT関数の実装（パーセント変換）
export const TO_PERCENT: CustomFormula = {
  name: 'TO_PERCENT',
  pattern: /TO_PERCENT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valueRef] = matches;
    
    try {
      const value = getCellValue(valueRef.trim(), context);
      const num = Number(value);
      
      if (isNaN(num)) {
        return FormulaError.VALUE;
      }
      
      return `${(num * 100).toFixed(2)}%`;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TO_DOLLARS関数の実装（ドル形式変換）
export const TO_DOLLARS: CustomFormula = {
  name: 'TO_DOLLARS',
  pattern: /TO_DOLLARS\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valueRef] = matches;
    
    try {
      const value = getCellValue(valueRef.trim(), context);
      const num = Number(value);
      
      if (isNaN(num)) {
        return FormulaError.VALUE;
      }
      
      return `$${num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',')}`;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// TO_TEXT関数の実装（テキスト変換）
export const TO_TEXT: CustomFormula = {
  name: 'TO_TEXT',
  pattern: /TO_TEXT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valueRef] = matches;
    
    try {
      const value = getCellValue(valueRef.trim(), context);
      
      if (value === null || value === undefined) {
        return '';
      }
      
      return String(value);
    } catch {
      return FormulaError.VALUE;
    }
  }
};