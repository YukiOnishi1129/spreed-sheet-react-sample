// 回帰分析関連の統計関数

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellRangeValues } from '../shared/utils';

// 回帰分析の計算結果を保持する型
interface RegressionResult {
  slope: number;
  intercept: number;
  rSquared: number;
  standardError: number;
  n: number;
  sumX: number;
  sumY: number;
  sumXY: number;
  sumX2: number;
  sumY2: number;
}

// 回帰分析を実行する共通関数
function calculateRegression(yValues: number[], xValues: number[]): RegressionResult | null {
  if (yValues.length !== xValues.length || yValues.length < 2) {
    return null;
  }
  
  // すべてのX値が同じかチェック
  const firstX = xValues[0];
  const allSameX = xValues.every(x => x === firstX);
  if (allSameX) {
    return null; // Xの分散が0
  }
  
  const n = yValues.length;
  let sumX = 0;
  let sumY = 0;
  let sumXY = 0;
  let sumX2 = 0;
  let sumY2 = 0;
  
  // 各種合計を計算
  for (let i = 0; i < n; i++) {
    sumX += xValues[i];
    sumY += yValues[i];
    sumXY += xValues[i] * yValues[i];
    sumX2 += xValues[i] * xValues[i];
    sumY2 += yValues[i] * yValues[i];
  }
  
  // 傾きを計算
  const denominator = n * sumX2 - sumX * sumX;
  if (Math.abs(denominator) < 1e-10) {
    return null; // 分母が0に近い場合
  }
  
  const slope = (n * sumXY - sumX * sumY) / denominator;
  
  // 切片を計算
  const intercept = (sumY - slope * sumX) / n;
  
  // 決定係数（R²）を計算
  const meanY = sumY / n;
  let ssTotal = 0;
  let ssResidual = 0;
  
  for (let i = 0; i < n; i++) {
    const predicted = slope * xValues[i] + intercept;
    ssTotal += Math.pow(yValues[i] - meanY, 2);
    ssResidual += Math.pow(yValues[i] - predicted, 2);
  }
  
  const rSquared = ssTotal > 0 ? 1 - (ssResidual / ssTotal) : 0;
  
  // 標準誤差を計算
  const standardError = n > 2 ? Math.sqrt(ssResidual / (n - 2)) : 0;
  
  return {
    slope,
    intercept,
    rSquared,
    standardError,
    n,
    sumX,
    sumY,
    sumXY,
    sumX2,
    sumY2
  };
}

// データを数値配列に変換
function parseDataRange(rangeRef: string, context: FormulaContext): number[] {
  const values = getCellRangeValues(rangeRef.trim(), context);
  const numbers: number[] = [];
  
  for (const value of values) {
    const num = typeof value === 'string' ? parseFloat(value) : Number(value);
    if (!isNaN(num) && isFinite(num)) {
      numbers.push(num);
    }
  }
  
  return numbers;
}

// SLOPE関数の実装（回帰直線の傾き）
export const SLOPE: CustomFormula = {
  name: 'SLOPE',
  pattern: /SLOPE\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, yRangeRef, xRangeRef] = matches;
    
    try {
      // 生データを取得
      const yRawValues = getCellRangeValues(yRangeRef.trim(), context);
      const xRawValues = getCellRangeValues(xRangeRef.trim(), context);
      
      // ペアで数値を抽出
      const yValues: number[] = [];
      const xValues: number[] = [];
      
      const minLength = Math.min(yRawValues.length, xRawValues.length);
      for (let i = 0; i < minLength; i++) {
        const yNum = typeof yRawValues[i] === 'string' ? parseFloat(yRawValues[i]) : Number(yRawValues[i]);
        const xNum = typeof xRawValues[i] === 'string' ? parseFloat(xRawValues[i]) : Number(xRawValues[i]);
        
        if (!isNaN(yNum) && isFinite(yNum) && !isNaN(xNum) && isFinite(xNum)) {
          yValues.push(yNum);
          xValues.push(xNum);
        }
      }
      
      if (yValues.length < 2) {
        return FormulaError.DIV0;
      }
      
      const result = calculateRegression(yValues, xValues);
      
      if (!result) {
        return FormulaError.DIV0;
      }
      
      return result.slope;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// INTERCEPT関数の実装（回帰直線の切片）
export const INTERCEPT: CustomFormula = {
  name: 'INTERCEPT',
  pattern: /INTERCEPT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, yRangeRef, xRangeRef] = matches;
    
    try {
      const yValues = parseDataRange(yRangeRef, context);
      const xValues = parseDataRange(xRangeRef, context);
      
      if (yValues.length === 0 || xValues.length === 0) {
        return FormulaError.VALUE;
      }
      
      if (yValues.length !== xValues.length) {
        return FormulaError.NA;
      }
      
      const result = calculateRegression(yValues, xValues);
      
      if (!result) {
        return FormulaError.DIV0;
      }
      
      return result.intercept;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// RSQ関数の実装（決定係数）
export const RSQ: CustomFormula = {
  name: 'RSQ',
  pattern: /RSQ\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, yRangeRef, xRangeRef] = matches;
    
    try {
      const yValues = parseDataRange(yRangeRef, context);
      const xValues = parseDataRange(xRangeRef, context);
      
      if (yValues.length === 0 || xValues.length === 0) {
        return FormulaError.VALUE;
      }
      
      if (yValues.length !== xValues.length) {
        return FormulaError.NA;
      }
      
      const result = calculateRegression(yValues, xValues);
      
      if (!result) {
        return FormulaError.DIV0;
      }
      
      return result.rSquared;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// STEYX関数の実装（回帰の標準誤差）
export const STEYX: CustomFormula = {
  name: 'STEYX',
  pattern: /STEYX\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, yRangeRef, xRangeRef] = matches;
    
    try {
      const yValues = parseDataRange(yRangeRef, context);
      const xValues = parseDataRange(xRangeRef, context);
      
      if (yValues.length === 0 || xValues.length === 0) {
        return FormulaError.VALUE;
      }
      
      if (yValues.length !== xValues.length) {
        return FormulaError.NA;
      }
      
      if (yValues.length < 3) {
        return FormulaError.DIV0; // 最低3つのデータポイントが必要
      }
      
      const result = calculateRegression(yValues, xValues);
      
      if (!result) {
        return FormulaError.DIV0;
      }
      
      return result.standardError;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// FORECAST関数の実装（線形回帰による予測）
export const FORECAST: CustomFormula = {
  name: 'FORECAST',
  pattern: /FORECAST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xValueRef, yRangeRef, xRangeRef] = matches;
    
    try {
      // 予測したいX値を取得
      const xValueStr = xValueRef.trim();
      let xValue: number;
      
      if (xValueStr.match(/^[A-Z]+\d+$/)) {
        const cellValue = getCellRangeValues(xValueStr, context)[0];
        xValue = Number(cellValue);
      } else {
        xValue = parseFloat(xValueStr);
      }
      
      if (isNaN(xValue)) {
        return FormulaError.VALUE;
      }
      
      const yValues = parseDataRange(yRangeRef, context);
      const xValues = parseDataRange(xRangeRef, context);
      
      if (yValues.length === 0 || xValues.length === 0) {
        return FormulaError.VALUE;
      }
      
      if (yValues.length !== xValues.length) {
        return FormulaError.NA;
      }
      
      const result = calculateRegression(yValues, xValues);
      
      if (!result) {
        return FormulaError.DIV0;
      }
      
      // 予測値を計算
      return result.slope * xValue + result.intercept;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// LINEST関数の実装（線形回帰の統計情報）
export const LINEST: CustomFormula = {
  name: 'LINEST',
  pattern: /LINEST\(([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, yRangeRef, xRangeRef, constRef, statsRef] = matches;
    
    try {
      const yValues = parseDataRange(yRangeRef, context);
      const xValues = parseDataRange(xRangeRef, context);
      
      if (yValues.length === 0 || xValues.length === 0) {
        return FormulaError.VALUE;
      }
      
      if (yValues.length !== xValues.length) {
        return FormulaError.NA;
      }
      
      const includeConst = constRef?.trim().toLowerCase() !== 'false';
      const includeStats = statsRef?.trim().toLowerCase() === 'true';
      
      const result = calculateRegression(yValues, xValues);
      
      if (!result) {
        return FormulaError.DIV0;
      }
      
      if (includeStats) {
        // 統計情報を含む配列を返す
        const n = result.n;
        const df = n - 2; // 自由度
        
        // 傾きの標準誤差
        const slopeStdError = result.standardError / Math.sqrt(result.sumX2 - result.sumX * result.sumX / n);
        
        // 切片の標準誤差
        const interceptStdError = result.standardError * Math.sqrt(result.sumX2 / (n * (result.sumX2 - result.sumX * result.sumX / n)));
        
        // F統計量
        const fStat = result.rSquared * df / (1 - result.rSquared);
        
        return [
          [result.slope, includeConst ? result.intercept : 0],
          [slopeStdError, includeConst ? interceptStdError : 0],
          [result.rSquared, result.standardError],
          [fStat, df],
          [result.rSquared * result.sumY2 - result.sumY * result.sumY / n, (1 - result.rSquared) * (result.sumY2 - result.sumY * result.sumY / n)]
        ];
      } else {
        // 傾きと切片のみを返す
        return [[result.slope, includeConst ? result.intercept : 0]];
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};