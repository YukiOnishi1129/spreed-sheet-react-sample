// EDATE関数のテスト
// 入力: '2024/1/15', 3
// 期待値: 44946

// Excelのシリアル値計算
// 1900年1月1日を1として計算（Excelの1900年日付システム）

const startDate = new Date(2024, 0, 15); // 2024年1月15日
console.log('開始日:', startDate.toISOString().split('T')[0]);

// 3ヶ月後
const resultDate = new Date(2024, 0 + 3, 15); // 2024年4月15日
console.log('3ヶ月後:', resultDate.toISOString().split('T')[0]);

// Excelシリアル値の計算
function dateToExcelSerial(date) {
  // Excelの基準日: 1900年1月1日（ただし、1900年は閏年として扱われる）
  const excelEpoch = new Date(1900, 0, 1);
  const daysDiff = Math.floor((date - excelEpoch) / (1000 * 60 * 60 * 24));
  
  // 1900年2月29日問題の補正
  // Excelは1900年を誤って閏年として扱うため、1900年3月1日以降は+1日
  const correction = date >= new Date(1900, 2, 1) ? 2 : 1;
  
  return daysDiff + correction;
}

const serial1 = dateToExcelSerial(startDate);
const serial2 = dateToExcelSerial(resultDate);

console.log('\n開始日のシリアル値:', serial1);
console.log('3ヶ月後のシリアル値:', serial2);
console.log('期待値:', 44946);

// 期待値44946が何の日付か確認
function excelSerialToDate(serial) {
  // Excelの補正を考慮
  const actualSerial = serial >= 61 ? serial - 2 : serial - 1;
  const excelEpoch = new Date(1900, 0, 1);
  const resultDate = new Date(excelEpoch.getTime() + actualSerial * 24 * 60 * 60 * 1000);
  return resultDate;
}

const expectedDate = excelSerialToDate(44946);
console.log('\n期待値44946の日付:', expectedDate.toISOString().split('T')[0]);

// 実際値45397の日付も確認
const actualDate = excelSerialToDate(45397);
console.log('実際値45397の日付:', actualDate.toISOString().split('T')[0]);