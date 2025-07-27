import { generateTestsForCategory } from '../generate-function-tests';
import { databaseTests } from '../../src/data/individualTests/10-database';

// データベース関数のテスト
generateTestsForCategory('10. データベース', databaseTests);