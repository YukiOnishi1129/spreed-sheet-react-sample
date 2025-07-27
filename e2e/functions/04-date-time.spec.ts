import { generateTestsForCategory } from '../generate-function-tests';
import { dateTimeTests } from '../../src/data/individualTests/04-date-time';

// 日付関数のテスト
generateTestsForCategory('04. 日付', dateTimeTests);