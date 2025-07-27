import { generateTestsForCategory } from '../generate-function-tests';
import { mathTrigonometryTests } from '../../src/data/individualTests/01-math-trigonometry';

// 数学・三角関数のテスト
generateTestsForCategory('01. 数学', mathTrigonometryTests);
generateTestsForCategory('01. 数学・三角', mathTrigonometryTests);