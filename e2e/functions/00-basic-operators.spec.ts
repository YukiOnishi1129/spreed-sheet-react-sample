import { generateTestsForCategory } from '../generate-function-tests';
import { basicOperatorsTests } from '../../src/data/individualTests/00-basic-operators';

// 基本演算子のテスト
generateTestsForCategory('00. 基本演算子', basicOperatorsTests);