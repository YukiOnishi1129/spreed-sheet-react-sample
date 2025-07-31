// 特定の関数のテスト結果を詳細に確認するスクリプト
const fs = require('fs');
const path = require('path');

// 失敗数が多い関数の詳細を確認
const targetFunctions = [
  'EXPAND',
  'VSTACK', 
  'CHOOSECOLS',
  'WRAPROWS',
  'WRAPCOLS',
  'SEQUENCE',
  'SERIESSUM',
  'TEXT',
  'EDATE',
  'HLOOKUP'
];

// 各関数のテストデータを読み込む
const testDataDir = './src/data/individualTests';
const testFiles = fs.readdirSync(testDataDir);

const functionDetails = {};

testFiles.forEach(file => {
  if (file.endsWith('.ts')) {
    const content = fs.readFileSync(path.join(testDataDir, file), 'utf8');
    
    targetFunctions.forEach(funcName => {
      // 関数のテストデータを探す
      const regex = new RegExp(`name:\\s*['"]${funcName}['"][^}]*expectedValues:\\s*{([^}]+)}`, 'gs');
      const match = regex.exec(content);
      
      if (match) {
        // expectedValuesの内容を抽出
        const expectedValuesStr = match[1];
        functionDetails[funcName] = {
          file: file,
          expectedValues: expectedValuesStr.trim()
        };
      }
    });
  }
});

console.log('=== 失敗している関数の期待値詳細 ===\n');

targetFunctions.forEach(funcName => {
  if (functionDetails[funcName]) {
    console.log(`${funcName} (${functionDetails[funcName].file}):`);
    console.log('期待値:');
    console.log(functionDetails[funcName].expectedValues);
    console.log('\n---\n');
  } else {
    console.log(`${funcName}: テストデータが見つかりません\n`);
  }
});