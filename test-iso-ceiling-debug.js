// ISO.CEILING関数のデバッグ
console.log('Testing ISO.CEILING(-4.3, 1):');

// Math.ceilの動作確認
console.log('Math.ceil(-4.3) =', Math.ceil(-4.3)); // -4
console.log('Math.ceil(-4.3 / 1) =', Math.ceil(-4.3 / 1)); // -4
console.log('Math.ceil(-4.3 / 1) * 1 =', Math.ceil(-4.3 / 1) * 1); // -4

// 正の数の場合
console.log('\nTesting ISO.CEILING(4.3, 1):');
console.log('Math.ceil(4.3) =', Math.ceil(4.3)); // 5
console.log('Math.ceil(4.3 / 1) =', Math.ceil(4.3 / 1)); // 5
console.log('Math.ceil(4.3 / 1) * 1 =', Math.ceil(4.3 / 1) * 1); // 5

// significance = -1の場合
console.log('\nTesting ISO.CEILING(-4.3, -1):');
console.log('Math.ceil(-4.3 / -1) =', Math.ceil(-4.3 / -1)); // 5
console.log('Math.ceil(-4.3 / -1) * -1 =', Math.ceil(-4.3 / -1) * -1); // -5