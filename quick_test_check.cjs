#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

// テストデータを読み込んで、expectedValuesがある関数をリスト
const testDir = path.join(__dirname, 'src/data/individualTests');
const files = fs.readdirSync(testDir).filter(f => f.endsWith('.ts'));

let totalFunctions = 0;
let functionsWithExpected = 0;
const problemFunctions = [];

files.forEach(file => {
  const content = fs.readFileSync(path.join(testDir, file), 'utf8');
  
  // name: 'FUNCTION_NAME' のパターンを探す
  const functions = content.match(/name:\s*['"]([^'"]+)['"]/g) || [];
  const expectedMatches = content.match(/expectedValues:\s*\{[^}]+\}/g) || [];
  
  totalFunctions += functions.length;
  functionsWithExpected += expectedMatches.length;
  
  // 各関数ブロックを解析
  const blocks = content.split(/(?=\s*{\s*name:)/);
  blocks.forEach(block => {
    const nameMatch = block.match(/name:\s*['"]([^'"]+)['"]/);
    if (nameMatch) {
      const funcName = nameMatch[1];
      const hasExpected = block.includes('expectedValues:');
      
      if (!hasExpected) {
        problemFunctions.push(`${funcName} (${file})`);
      } else {
        // expectedValuesがあるが、値に問題がある可能性をチェック
        const expectedMatch = block.match(/expectedValues:\s*\{([^}]+)\}/);
        if (expectedMatch) {
          const expectedContent = expectedMatch[1];
          
          // 動的な値や問題のあるパターンをチェック
          if (funcName === 'TODAY' || funcName === 'NOW' || funcName === 'RAND' || funcName === 'RANDBETWEEN') {
            console.log(`警告: ${funcName} は動的関数です`);
          }
          
          // #VALUE! などのエラー値をチェック
          if (expectedContent.includes('#VALUE!')) {
            console.log(`エラー: ${funcName} のexpectedValuesに#VALUE!が含まれています`);
          }
        }
      }
    }
  });
});

console.log('\n=== サマリー ===');
console.log(`総関数数: ${totalFunctions}`);
console.log(`expectedValues有り: ${functionsWithExpected}`);
console.log(`expectedValues無し: ${problemFunctions.length}`);

if (problemFunctions.length > 0) {
  console.log('\n=== expectedValuesが無い関数 ===');
  problemFunctions.forEach(func => console.log(`- ${func}`));
}