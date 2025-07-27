import { generateTestsForCategory } from '../generate-function-tests';
import { financialTests } from '../../src/data/individualTests/07-financial';

// 財務関数のテスト
generateTestsForCategory('07. 財務', financialTests);