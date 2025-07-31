// TEXT関数のフォーマットテスト
const value = 1234.56;
const format = '#,##0.00';

// Node.jsのtoLocaleStringでテスト
console.log('toLocaleString("en-US"):', value.toLocaleString('en-US'));
console.log('toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }):', 
  value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));

// 手動でフォーマット
function formatNumber(value, format) {
  if (format === '#,##0.00') {
    return value.toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
      useGrouping: true
    });
  }
  return value.toString();
}

console.log('formatNumber結果:', formatNumber(value, format));