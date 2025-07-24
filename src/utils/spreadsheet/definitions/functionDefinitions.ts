// Excel/Google Sheets関数の詳細定義
// spreadsheet-functions-reference.mdに基づく完全な関数リスト

import { type FunctionDefinition } from '../types/types';

// カテゴリ別の色定義をインポート
export { COLOR_SCHEMES } from '../types/colorSchemes';

// 各ファイルから関数定義をインポート
import { MATH_FUNCTIONS } from './mathDefinitions';
import { STATISTICAL_FUNCTIONS } from './statisticalDefinitions';
import { TEXT_FUNCTIONS } from './textDefinitions';
import { DATETIME_FUNCTIONS } from './datetimeDefinitions';
import { LOGICAL_FUNCTIONS } from './logicalDefinitions';
import { LOOKUP_FUNCTIONS } from './lookupDefinitions';
import { INFORMATION_FUNCTIONS } from './informationDefinitions';
import { OTHER_FUNCTIONS } from './otherDefinitions';
import { DATABASE_FUNCTIONS } from './databaseDefinitions';
import { FINANCIAL_FUNCTIONS } from './financialDefinitions';
import { ENGINEERING_FUNCTIONS } from './engineeringDefinitions';
import { CUBE_FUNCTIONS } from './cubeDefinitions';
import { WEB_FUNCTIONS } from './webDefinitions';
import { GOOGLESHEETS_FUNCTIONS } from './googleSheetsDefinitions';

// 全ての関数を統合
export const FUNCTION_DEFINITIONS: Record<string, FunctionDefinition> = {
  ...MATH_FUNCTIONS,
  ...STATISTICAL_FUNCTIONS,
  ...TEXT_FUNCTIONS,
  ...DATETIME_FUNCTIONS,
  ...LOGICAL_FUNCTIONS,
  ...LOOKUP_FUNCTIONS,
  ...INFORMATION_FUNCTIONS,
  ...OTHER_FUNCTIONS,
  ...DATABASE_FUNCTIONS,
  ...FINANCIAL_FUNCTIONS,
  ...ENGINEERING_FUNCTIONS,
  ...CUBE_FUNCTIONS,
  ...WEB_FUNCTIONS,
  ...GOOGLESHEETS_FUNCTIONS
};

/**
 * 関数名から使用されている個別の関数を抽出
 */
export function extractFunctionsFromName(functionName: string): string[] {
  const functions: string[] = [];
  
  // &で区切られた関数名を抽出
  const parts = functionName.split('&').map(part => part.trim());
  
  for (const part of parts) {
    // 各パートから関数名を抽出（英語の関数名のみ）
    const matches = part.match(/[A-Z][A-Z0-9.]*/g);
    if (matches) {
      functions.push(...matches);
    }
  }
  
  // 重複を除去して返す
  return [...new Set(functions)];
}

/**
 * 使用関数の詳細説明を生成（見やすく整形）
 */
export function generateFunctionDetails(functions: string[]): string {
  let details = '';
  
  for (const func of functions) {
    const definition = FUNCTION_DEFINITIONS[func];
    if (!definition) continue;
    
    // 関数名とシンタックスを表示
    details += `${definition.syntax}\n`;
    
    // 各引数を改行して表示
    for (const param of definition.params) {
      const optional = param.optional ? '（省略可能）' : '';
      details += `- ${param.name}: ${param.desc}${optional}\n`;
    }
    
    // 関数間にスペースを追加
    details += '\n';
  }
  
  return details.trim();
}

/**
 * 関数タイプを取得する関数
 */
export function getFunctionType(functionName: string): string {
  const def = FUNCTION_DEFINITIONS[functionName];
  return def ? def.category : 'other';
}

/**
 * 関数の定義を取得
 */
export function getFunctionDefinition(functionName: string): FunctionDefinition | undefined {
  return FUNCTION_DEFINITIONS[functionName.toUpperCase()];
}

/**
 * カテゴリから関数リストを取得
 */
export function getFunctionsByCategory(category: string): FunctionDefinition[] {
  return Object.values(FUNCTION_DEFINITIONS).filter(def => def.category === category);
}

/**
 * すべての関数名を取得
 */
export function getAllFunctionNames(): string[] {
  return Object.keys(FUNCTION_DEFINITIONS);
}

/**
 * 数式から使用されている関数を抽出
 */
export function extractFunctionsFromFormula(formula: string): string[] {
  const functions: string[] = [];
  const functionPattern = /([A-Z_]+(?:\.[A-Z]+)?)\s*\(/g;
  let match;
  
  while ((match = functionPattern.exec(formula)) !== null) {
    const funcName = match[1];
    if (FUNCTION_DEFINITIONS[funcName]) {
      functions.push(funcName);
    }
  }
  
  return [...new Set(functions)]; // 重複を除去
}