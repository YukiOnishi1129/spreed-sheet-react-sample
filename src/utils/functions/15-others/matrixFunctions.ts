// 行列関数の実装

import type { CustomFormula, FormulaContext } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue, getCellRangeValues } from '../shared/utils';

// 2次元配列として行列を解析するヘルパー関数
function parseMatrix(rangeRef: string, context: FormulaContext): number[][] {
  const values = getCellRangeValues(rangeRef, context);
  const matrix: number[][] = [];
  
  // 行数と列数を推定（簡単な実装）
  // 実際の実装では、rangeRefから正確な行列サイズを取得する必要がある
  const totalValues = values.length;
  const rows = Math.ceil(Math.sqrt(totalValues));
  const cols = Math.ceil(totalValues / rows);
  
  for (let i = 0; i < rows; i++) {
    const row: number[] = [];
    for (let j = 0; j < cols; j++) {
      const index = i * cols + j;
      if (index < values.length) {
        const num = Number(values[index]);
        if (isNaN(num)) {
          throw new Error('Matrix contains non-numeric values');
        }
        row.push(num);
      } else {
        row.push(0);
      }
    }
    matrix.push(row);
  }
  
  return matrix;
}

// 行列の行列式を計算するヘルパー関数
function determinant(matrix: number[][]): number {
  const n = matrix.length;
  
  if (n !== matrix[0].length) {
    throw new Error('Matrix must be square');
  }
  
  if (n === 1) return matrix[0][0];
  if (n === 2) return matrix[0][0] * matrix[1][1] - matrix[0][1] * matrix[1][0];
  
  let det = 0;
  for (let j = 0; j < n; j++) {
    const subMatrix = matrix.slice(1).map(row => row.filter((_, index) => index !== j));
    det += matrix[0][j] * Math.pow(-1, j) * determinant(subMatrix);
  }
  
  return det;
}

// 行列の逆行列を計算するヘルパー関数
function inverse(matrix: number[][]): number[][] {
  const n = matrix.length;
  
  if (n !== matrix[0].length) {
    throw new Error('Matrix must be square');
  }
  
  const det = determinant(matrix);
  if (Math.abs(det) < Number.EPSILON) {
    throw new Error('Matrix is singular (determinant is zero)');
  }
  
  if (n === 1) return [[1 / matrix[0][0]]];
  
  // 余因子行列を作成
  const cofactorMatrix: number[][] = [];
  for (let i = 0; i < n; i++) {
    cofactorMatrix[i] = [];
    for (let j = 0; j < n; j++) {
      const subMatrix = matrix
        .filter((_, rowIndex) => rowIndex !== i)
        .map(row => row.filter((_, colIndex) => colIndex !== j));
      
      const cofactor = Math.pow(-1, i + j) * determinant(subMatrix);
      cofactorMatrix[i][j] = cofactor;
    }
  }
  
  // 転置して逆行列を計算
  const inverseMatrix: number[][] = [];
  for (let i = 0; i < n; i++) {
    inverseMatrix[i] = [];
    for (let j = 0; j < n; j++) {
      inverseMatrix[i][j] = cofactorMatrix[j][i] / det;
    }
  }
  
  return inverseMatrix;
}

// 行列の積を計算するヘルパー関数
function multiply(matrix1: number[][], matrix2: number[][]): number[][] {
  const rows1 = matrix1.length;
  const cols1 = matrix1[0].length;
  const rows2 = matrix2.length;
  const cols2 = matrix2[0].length;
  
  if (cols1 !== rows2) {
    throw new Error('Matrix dimensions do not match for multiplication');
  }
  
  const result: number[][] = [];
  for (let i = 0; i < rows1; i++) {
    result[i] = [];
    for (let j = 0; j < cols2; j++) {
      let sum = 0;
      for (let k = 0; k < cols1; k++) {
        sum += matrix1[i][k] * matrix2[k][j];
      }
      result[i][j] = sum;
    }
  }
  
  return result;
}

// MDETERM関数（行列式を計算）
export const MDETERM: CustomFormula = {
  name: 'MDETERM',
  pattern: /MDETERM\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRef] = matches;
    
    try {
      const matrix = parseMatrix(arrayRef, context);
      return determinant(matrix);
    } catch (error) {
      if (error instanceof Error) {
        if (error.message.includes('non-numeric')) {
          return FormulaError.VALUE;
        }
        if (error.message.includes('square')) {
          return FormulaError.VALUE;
        }
      }
      return FormulaError.NUM;
    }
  }
};

// MINVERSE関数（逆行列を計算）
export const MINVERSE: CustomFormula = {
  name: 'MINVERSE',
  pattern: /MINVERSE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, arrayRef] = matches;
    
    try {
      const matrix = parseMatrix(arrayRef, context);
      const inverseMatrix = inverse(matrix);
      
      // 結果が1x1の場合は単一の値を返す
      if (inverseMatrix.length === 1 && inverseMatrix[0].length === 1) {
        return inverseMatrix[0][0];
      }
      
      return inverseMatrix;
    } catch (error) {
      if (error instanceof Error) {
        if (error.message.includes('non-numeric')) {
          return FormulaError.VALUE;
        }
        if (error.message.includes('square') || error.message.includes('singular')) {
          return FormulaError.NUM;
        }
      }
      return FormulaError.NUM;
    }
  }
};

// MMULT関数（行列の積を計算）
export const MMULT: CustomFormula = {
  name: 'MMULT',
  pattern: /MMULT\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, array1Ref, array2Ref] = matches;
    
    try {
      const matrix1 = parseMatrix(array1Ref, context);
      const matrix2 = parseMatrix(array2Ref, context);
      const result = multiply(matrix1, matrix2);
      
      // 結果が1x1の場合は単一の値を返す
      if (result.length === 1 && result[0].length === 1) {
        return result[0][0];
      }
      
      return result;
    } catch (error) {
      if (error instanceof Error) {
        if (error.message.includes('non-numeric')) {
          return FormulaError.VALUE;
        }
        if (error.message.includes('dimensions')) {
          return FormulaError.VALUE;
        }
      }
      return FormulaError.NUM;
    }
  }
};

// MUNIT関数（単位行列を作成）
export const MUNIT: CustomFormula = {
  name: 'MUNIT',
  pattern: /MUNIT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext) => {
    const [, sizeRef] = matches;
    
    const size = Number(getCellValue(sizeRef, context) ?? sizeRef);
    
    if (isNaN(size) || !Number.isInteger(size) || size <= 0) {
      return FormulaError.VALUE;
    }
    
    if (size > 1000) {
      return FormulaError.NUM;
    }
    
    const matrix: number[][] = [];
    for (let i = 0; i < size; i++) {
      matrix[i] = [];
      for (let j = 0; j < size; j++) {
        matrix[i][j] = i === j ? 1 : 0;
      }
    }
    
    // 1x1の場合は単一の値を返す
    if (size === 1) {
      return 1;
    }
    
    return matrix;
  }
};