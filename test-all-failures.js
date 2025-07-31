// 全ての失敗を集計するスクリプト
import { allIndividualFunctionTests } from './src/data/individualFunctionTests.js';

const functionsWithExpectedValues = allIndividualFunctionTests.filter(
  func => func.expectedValues && Object.keys(func.expectedValues).length > 0
);

// カテゴリ別に集計
const categories = {};

functionsWithExpectedValues.forEach(func => {
  if (!categories[func.category]) {
    categories[func.category] = [];
  }
  categories[func.category].push(func.name);
});

console.log('=== 期待値を持つ関数の一覧 ===');
Object.entries(categories).forEach(([category, functions]) => {
  console.log(`\n${category}: ${functions.length}個`);
  console.log('関数:', functions.join(', '));
});