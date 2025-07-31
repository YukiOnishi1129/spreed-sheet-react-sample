// ブラウザコンソールで実行するコード
// http://localhost:5173/test で実行してください

(() => {
  // すべての失敗を検出
  const failedElements = document.querySelectorAll('.bg-red-50');
  const failures = [];
  
  failedElements.forEach(el => {
    const cellAddress = el.querySelector('.font-mono')?.textContent;
    const expected = el.textContent.match(/期待値: ([^\n]+)/)?.[1];
    const actual = el.textContent.match(/実際値: ([^\n]+)/)?.[1];
    
    // 親要素から関数名を取得
    let functionName = 'Unknown';
    let parent = el.parentElement;
    while (parent) {
      const header = parent.querySelector('h3');
      if (header && header.textContent.includes('関数:')) {
        functionName = header.textContent.replace('関数:', '').trim();
        break;
      }
      parent = parent.parentElement;
    }
    
    failures.push({
      function: functionName,
      cell: cellAddress,
      expected: expected,
      actual: actual
    });
  });
  
  // カテゴリ別に集計
  const categories = {};
  document.querySelectorAll('option').forEach(opt => {
    if (opt.value && opt.textContent.match(/^\d+\./)) {
      const categoryName = opt.textContent;
      const categorySection = document.querySelector(`h2:has(+ * option[value="${opt.value}"])`);
      if (categorySection) {
        const failureCount = document.querySelectorAll(`.bg-red-50`).length;
        categories[categoryName] = failureCount;
      }
    }
  });
  
  console.log('=== テスト失敗サマリー ===');
  console.log(`総失敗数: ${failures.length}`);
  console.log('\n=== 失敗詳細 ===');
  failures.slice(0, 10).forEach(f => {
    console.log(`${f.function} - ${f.cell}: 期待値=${f.expected}, 実際値=${f.actual}`);
  });
  
  if (failures.length > 10) {
    console.log(`... 他 ${failures.length - 10} 件の失敗`);
  }
  
  // 実装されていない関数のリスト
  const notImplemented = failures.filter(f => f.actual === '#VALUE!');
  console.log(`\n=== 未実装の関数 (${notImplemented.length}件) ===`);
  const uniqueFunctions = [...new Set(notImplemented.map(f => f.function))];
  uniqueFunctions.forEach(func => console.log(`- ${func}`));
  
  return {
    totalFailures: failures.length,
    notImplemented: notImplemented.length,
    failures: failures
  };
})();