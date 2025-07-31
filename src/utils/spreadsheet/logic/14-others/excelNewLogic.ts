// Excel専用の新関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';
import { parseDate } from '../shared/dateUtils';

// ISOMITTED関数の実装（引数が省略されたか判定）- Excel専用
export const ISOMITTED: CustomFormula = {
  name: 'ISOMITTED',
  pattern: /ISOMITTED\(([^)]*)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, argumentRef] = matches;
    
    // 引数が空文字列または未定義の場合は省略されたと判定
    if (!argumentRef || argumentRef.trim() === '') {
      return true;
    }
    
    const value = getCellValue(argumentRef.trim(), context);
    
    // nullまたはundefinedの場合は省略されたと判定
    return value === null || value === undefined || value === '';
  }
};

// STOCKHISTORY関数の実装（株価履歴を取得）- Excel専用
export const STOCKHISTORY: CustomFormula = {
  name: 'STOCKHISTORY',
  pattern: /STOCKHISTORY\(([^,]+),\s*([^,]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, , startDateRef, endDateRef, intervalRef, , propertiesRef] = matches;
    
    try {
      // const stock = getCellValue(stockRef.trim(), context)?.toString() ?? stockRef.trim();
      const startDate = parseDate(getCellValue(startDateRef.trim(), context)?.toString() ?? startDateRef.trim());
      const endDate = endDateRef ? parseDate(getCellValue(endDateRef.trim(), context)?.toString() ?? endDateRef.trim()) : new Date();
      const interval = intervalRef ? parseInt(getCellValue(intervalRef.trim(), context)?.toString() ?? intervalRef.trim()) : 0;
      
      if (!startDate) {
        return FormulaError.VALUE;
      }
      
      // ダミーの株価データを生成
      const data = [];
      const headers = ['Date', 'Close', 'Open', 'High', 'Low', 'Volume'];
      
      // ヘッダーを追加
      if (!propertiesRef || propertiesRef.trim() === '0') {
        data.push(headers);
      }
      
      // 日次データを生成（intervalが0の場合）
      if (interval === 0) {
        const currentDate = new Date(startDate);
        let basePrice = 100;
        
        while (currentDate <= (endDate ?? new Date())) {
          // ランダムな株価変動を生成
          const change = (Math.random() - 0.5) * 5;
          basePrice += change;
          const open = basePrice + (Math.random() - 0.5) * 2;
          const close = basePrice;
          const high = Math.max(open, close) + Math.random() * 2;
          const low = Math.min(open, close) - Math.random() * 2;
          const volume = Math.floor(Math.random() * 1000000) + 100000;
          
          data.push([
            currentDate.toLocaleDateString(),
            close.toFixed(2),
            open.toFixed(2),
            high.toFixed(2),
            low.toFixed(2),
            volume.toString()
          ]);
          
          // 次の日へ
          currentDate.setDate(currentDate.getDate() + 1);
          
          // 土日をスキップ
          if (currentDate.getDay() === 0) {
            currentDate.setDate(currentDate.getDate() + 1);
          } else if (currentDate.getDay() === 6) {
            currentDate.setDate(currentDate.getDate() + 2);
          }
        }
      } else {
        // 週次または月次データ（簡易実装）
        const periods = Math.floor(((endDate ?? new Date()).getTime() - startDate.getTime()) / (interval * 24 * 60 * 60 * 1000));
        let basePrice = 100;
        
        for (let i = 0; i < Math.min(periods, 100); i++) {
          const periodDate = new Date(startDate.getTime() + i * interval * 24 * 60 * 60 * 1000);
          const change = (Math.random() - 0.5) * 10;
          basePrice += change;
          
          data.push([
            periodDate.toLocaleDateString(),
            basePrice.toFixed(2),
            (basePrice + (Math.random() - 0.5) * 2).toFixed(2),
            (basePrice + Math.random() * 3).toFixed(2),
            (basePrice - Math.random() * 3).toFixed(2),
            Math.floor(Math.random() * 2000000).toString()
          ]);
        }
      }
      
      return data;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// GPT関数の実装（ChatGPTを呼び出し）- AI関連関数
export const GPT: CustomFormula = {
  name: 'GPT',
  pattern: /GPT\(([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, promptRef, temperatureRef, maxTokensRef] = matches;
    
    try {
      const prompt = getCellValue(promptRef.trim(), context)?.toString() ?? promptRef.trim();
      const temperature = temperatureRef ? parseFloat(getCellValue(temperatureRef.trim(), context)?.toString() ?? temperatureRef.trim()) : 0.7;
      const maxTokens = maxTokensRef ? parseInt(getCellValue(maxTokensRef.trim(), context)?.toString() ?? maxTokensRef.trim()) : 150;
      
      if (!prompt || prompt === '') {
        return FormulaError.VALUE;
      }
      
      if (temperature < 0 || temperature > 2) {
        return FormulaError.NUM;
      }
      
      if (maxTokens < 1 || maxTokens > 4000) {
        return FormulaError.NUM;
      }
      
      // 実際のAPI呼び出しは実装せず、ダミーレスポンスを返す
      // プロンプトに基づいて異なるレスポンスを生成
      const promptLower = prompt.toLowerCase();
      
      if (promptLower.includes('translate')) {
        return 'Translation: ' + prompt.replace(/translate|翻訳/gi, '').trim();
      } else if (promptLower.includes('summarize') || promptLower.includes('要約')) {
        return 'Summary: This is a brief summary of the provided text.';
      } else if (promptLower.includes('explain') || promptLower.includes('説明')) {
        return 'Explanation: This is a detailed explanation of the concept.';
      } else if (promptLower.includes('calculate') || promptLower.includes('計算')) {
        return 'Calculation result: 42';
      } else if (promptLower.includes('list') || promptLower.includes('リスト')) {
        return '1. First item\n2. Second item\n3. Third item';
      } else {
        return `AI Response to "${prompt.substring(0, 50)}...": This is a simulated response from the AI model.`;
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// 注: LAMBDA, MAKEARRAY, BYROW, BYCOL, REDUCE, SCAN, MAP関数は
// 06-lookup-reference/lambdaArrayLogic.tsで実装されています

// 注: FILTER, SORT, UNIQUE関数は
// 06-lookup-reference/lookupLogic.tsで実装されています