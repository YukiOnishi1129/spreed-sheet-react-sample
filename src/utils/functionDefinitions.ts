// Excel/Google Sheets関数の詳細定義
// spreadsheet-functions-reference.mdに基づく完全な関数リスト

export interface FunctionParam {
  name: string;
  desc: string;
  optional?: boolean;
}

export interface FunctionDefinition {
  name: string;
  syntax: string;
  params: FunctionParam[];
  description: string;
  category: string;
  examples?: string[];
  colorScheme?: {
    bg: string;
    fc: string;
  };
}

// カテゴリ別の色定義
export const COLOR_SCHEMES = {
  math: { bg: "#FFE0B2", fc: "#D84315" },         // 数学・三角関数 - オレンジ
  statistical: { bg: "#FFE0B2", fc: "#D84315" },  // 統計関数 - オレンジ
  text: { bg: "#FCE4EC", fc: "#C2185B" },         // 文字列関数 - ピンク
  datetime: { bg: "#F3E5F5", fc: "#7B1FA2" },     // 日付・時刻関数 - 紫
  logical: { bg: "#E8F5E8", fc: "#2E7D32" },      // 論理関数 - 緑
  lookup: { bg: "#E3F2FD", fc: "#1976D2" },       // 検索・参照関数 - 青
  information: { bg: "#F5F5F5", fc: "#616161" },   // 情報関数 - グレー
  database: { bg: "#FFF8E1", fc: "#F57C00" },     // データベース関数 - 薄い黄
  financial: { bg: "#FFFDE7", fc: "#F57F17" },    // 財務関数 - 黄色
  engineering: { bg: "#E0F2F1", fc: "#00796B" },  // エンジニアリング関数 - 青緑
  cube: { bg: "#F5F5F5", fc: "#616161" },         // キューブ関数 - グレー
  web: { bg: "#E8EAF6", fc: "#3F51B5" },          // Web関数 - インディゴ
  googlesheets: { bg: "#E8F5E9", fc: "#43A047" }, // Google Sheets専用 - 緑
  other: { bg: "#F5F5F5", fc: "#616161" }         // その他 - グレー
};

// 各ファイルから関数定義をインポート
import { MATH_FUNCTIONS } from './functions/mathFunctions';
import { STATISTICAL_FUNCTIONS } from './functions/statisticalFunctions';
import { TEXT_FUNCTIONS } from './functions/textFunctions';
import { DATETIME_FUNCTIONS } from './functions/datetimeFunctions';
import { LOGICAL_FUNCTIONS } from './functions/logicalFunctions';
import { LOOKUP_FUNCTIONS } from './functions/lookupFunctions';
import { INFORMATION_FUNCTIONS } from './functions/informationFunctions';
import { OTHER_FUNCTIONS } from './functions/otherFunctions';

// 全ての関数を統合
export const FUNCTION_DEFINITIONS: Record<string, FunctionDefinition> = {
  ...MATH_FUNCTIONS,
  ...STATISTICAL_FUNCTIONS,
  ...TEXT_FUNCTIONS,
  ...DATETIME_FUNCTIONS,
  ...LOGICAL_FUNCTIONS,
  ...LOOKUP_FUNCTIONS,
  ...INFORMATION_FUNCTIONS,
  ...OTHER_FUNCTIONS
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