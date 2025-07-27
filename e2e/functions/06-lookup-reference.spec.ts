import { generateTestsForCategory } from '../generate-function-tests';
import { lookupReferenceTests } from '../../src/data/individualTests/06-lookup-reference';

// 検索・参照関数のテスト
generateTestsForCategory('06. 検索・参照', lookupReferenceTests);