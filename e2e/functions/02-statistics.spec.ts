import { generateTestsForCategory } from '../generate-function-tests';
import { statisticsTests } from '../../src/data/individualTests/02-statistics';

// 統計関数のテスト
generateTestsForCategory('02. 統計', statisticsTests);