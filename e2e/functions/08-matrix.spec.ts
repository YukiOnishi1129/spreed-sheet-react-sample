import { generateTestsForCategory } from '../generate-function-tests';
import { matrixTests } from '../../src/data/individualTests/08-matrix';

// 行列関数のテスト
generateTestsForCategory('08. 行列', matrixTests);