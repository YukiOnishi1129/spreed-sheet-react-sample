import { allIndividualFunctionTests } from './src/data/individualFunctionTests.js';

// カテゴリ別に集計
const categories = {};

allIndividualFunctionTests.forEach(func => {
  if (!categories[func.category]) {
    categories[func.category] = {
      total: 0,
      withExpectedValues: 0,
      functions: []
    };
  }
  
  categories[func.category].total++;
  
  if (func.expectedValues && Object.keys(func.expectedValues).length > 0) {
    categories[func.category].withExpectedValues++;
    categories[func.category].functions.push(func.name);
  }
});

console.log('=== カテゴリ別の関数一覧 ===\n');

Object.entries(categories).sort().forEach(([category, data]) => {
  console.log(`${category}:`);
  console.log(`  総関数数: ${data.total}`);
  console.log(`  期待値あり: ${data.withExpectedValues}`);
  if (data.withExpectedValues > 0) {
    console.log(`  関数リスト: ${data.functions.join(', ')}`);
  }
  console.log('');
});

const totalFunctions = Object.values(categories).reduce((sum, cat) => sum + cat.total, 0);
const totalWithExpected = Object.values(categories).reduce((sum, cat) => sum + cat.withExpectedValues, 0);

console.log(`\n総計: ${totalFunctions}個の関数中、${totalWithExpected}個に期待値あり`);