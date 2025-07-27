import { generateTestsForCategory } from '../generate-function-tests';
import { textTests } from '../../src/data/individualTests/03-text';

// テキスト関数のテスト
generateTestsForCategory('03. 文字列', textTests);