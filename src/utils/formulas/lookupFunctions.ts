// 検索関数の実装

import type { CustomFormula, FormulaContext } from './types';
import { FormulaError } from './types';
import { getCellValue, getCellRangeValues } from './utils';

// VLOOKUP関数の実装
export const VLOOKUP: CustomFormula = {
  name: 'VLOOKUP',
  pattern: /VLOOKUP\(([^,]+),\s*([^,]+),\s*(\d+),\s*(TRUE|FALSE|0|1)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// HLOOKUP関数の実装
export const HLOOKUP: CustomFormula = {
  name: 'HLOOKUP',
  pattern: /HLOOKUP\(([^,]+),\s*([^,]+),\s*(\d+),\s*(TRUE|FALSE|0|1)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// INDEX関数の実装
export const INDEX: CustomFormula = {
  name: 'INDEX',
  pattern: /INDEX\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// MATCH関数の実装
export const MATCH: CustomFormula = {
  name: 'MATCH',
  pattern: /MATCH\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// LOOKUP関数の実装
export const LOOKUP: CustomFormula = {
  name: 'LOOKUP',
  pattern: /LOOKUP\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  isSupported: true, // HyperFormulaでサポート
  calculate: () => null // HyperFormulaが処理
};

// XLOOKUP関数の実装（手動実装が必要）
export const XLOOKUP: CustomFormula = {
  name: 'XLOOKUP',
  pattern: /XLOOKUP\(([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  isSupported: false, // HyperFormulaでサポートされていない
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, lookupValue, lookupArray, returnArray, ifNotFound] = matches;
    
    console.log('XLOOKUP計算:', { lookupValue, lookupArray, returnArray });
    
    // セル参照から値を取得
    let searchValue: unknown;
    if (lookupValue.match(/^[A-Z]+\d+$/)) {
      searchValue = getCellValue(lookupValue, context);
    } else {
      searchValue = lookupValue.replace(/^"|"$/g, '');
    }
    
    // 検索配列の値を取得
    const lookupValues = getCellRangeValues(lookupArray, context);
    // 返り値配列の値を取得
    const returnValues = getCellRangeValues(returnArray, context);
    
    if (lookupValues.length !== returnValues.length) {
      return FormulaError.REF;
    }
    
    // 完全一致で検索（デフォルト）
    for (let i = 0; i < lookupValues.length; i++) {
      if (String(lookupValues[i]).toLowerCase() === String(searchValue).toLowerCase()) {
        return returnValues[i];
      }
    }
    
    // 見つからない場合
    if (ifNotFound) {
      return ifNotFound.replace(/^"|"$/g, '');
    }
    
    return FormulaError.NA;
  }
};