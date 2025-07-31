// Test failure analysis script
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Read test results
const testResultsPath = path.join(__dirname, 'test-results.json');
if (!fs.existsSync(testResultsPath)) {
  console.log('test-results.json not found');
  process.exit(1);
}

const testResults = JSON.parse(fs.readFileSync(testResultsPath, 'utf-8'));

// Functions reported as failing in the test
const failingFunctions = [
  'SERIESSUM', 'SKEW', 'GEOMEAN', 'HARMEAN', 'TRIMMEAN', 'GAMMALN', 'GAUSS', 'STEYX',
  'TEXT', 'NUMBERVALUE', 'SEARCHB', 'EDATE', 'EOMONTH', 'DATEVALUE', 'TIMEVALUE', 'ISOWEEKNUM',
  'HLOOKUP', 'INDIRECT', 'HYPERLINK', 'FORMULATEXT', 'GETPIVOTDATA', 'RATE', 'NPER', 'IRR',
  'XNPV', 'XIRR', 'IPMT', 'PPMT', 'MIRR', 'SLN', 'ACCRINT', 'DURATION', 'MDURATION', 'PRICE',
  'COUPDAYS', 'COUPNCD', 'AMORDEGRC', 'CUMIPMT', 'CUMPRINC', 'ODDFPRICE', 'ODDLPRICE', 'ODDLYIELD',
  'TBILLPRICE', 'PRICEDISC', 'RECEIVED', 'INTRATE', 'PRICEMAT', 'YIELDMAT', 'MINVERSE',
  'ISTEXT', 'ISNUMBER', 'TYPE', 'SHEET', 'SHEETS', 'CELL', 'INFO', 'DSUM', 'DAVERAGE', 'DCOUNT',
  'DCOUNTA', 'DPRODUCT', 'DGET', 'IMABS', 'IMSUM', 'IMDIV', 'IMPOWER', 'IMLOG10', 'IMLOG2',
  'PHONETIC', 'IMSQRT', 'IMEXP', 'IMLN', 'IMSIN', 'IMCOS', 'IMTAN', 'BESSELY', 'BITAND', 'BITXOR',
  'TRANSPOSE', 'SEQUENCE', 'LAMBDA', 'HSTACK', 'VSTACK', 'BYROW', 'BYCOL', 'MAKEARRAY', 'MAP',
  'REDUCE', 'SCAN', 'TAKE', 'DROP', 'EXPAND', 'TOCOL', 'TOROW', 'CHOOSEROWS', 'CHOOSECOLS',
  'WRAPROWS', 'WRAPCOLS', 'CUBEVALUE', 'CUBESETCOUNT', 'REGEXEXTRACT', 'REGEXMATCH', 'REGEXREPLACE',
  'SORTN', 'WEBSERVICE', 'SPARKLINE', 'IMPORTDATA', 'IMPORTFEED', 'IMPORTHTML', 'IMPORTXML',
  'IMPORTRANGE', 'IMAGE', 'GOOGLEFINANCE', 'GOOGLETRANSLATE', 'DETECTLANGUAGE', 'TO_DATE',
  'TO_PERCENT', 'TO_TEXT', 'ISOMITTED', 'STOCKHISTORY'
];

console.log('\n=== Detailed Test Failure Analysis ===\n');

// Analyze each failing function
failingFunctions.forEach(funcName => {
  const result = testResults[funcName];
  if (result) {
    console.log(`\n${funcName}:`);
    console.log(`  Status: ${result.status}`);
    console.log(`  Has Expected Values: ${result.hasExpectedValues ? 'Yes' : 'No'}`);
    console.log(`  Pass Count: ${result.passCount}`);
    console.log(`  Fail Count: ${result.failCount}`);
    
    if (result.failCount > 0 && result.failures) {
      console.log('  Failures:');
      result.failures.forEach((failure, index) => {
        if (index < 3) { // Show first 3 failures
          console.log(`    ${failure.cell}: expected ${JSON.stringify(failure.expected)}, got ${JSON.stringify(failure.actual)}`);
        }
      });
      if (result.failures.length > 3) {
        console.log(`    ... and ${result.failures.length - 3} more failures`);
      }
    }
  } else {
    console.log(`\n${funcName}: No test result data found`);
  }
});

// Summary
console.log('\n=== Summary ===');
const implemented = failingFunctions.filter(f => testResults[f] && testResults[f].status === 'implemented').length;
const unimplemented = failingFunctions.filter(f => testResults[f] && testResults[f].status === 'unimplemented').length;
const notFound = failingFunctions.filter(f => !testResults[f]).length;

console.log(`Total failing functions: ${failingFunctions.length}`);
console.log(`  Implemented but failing: ${implemented}`);
console.log(`  Unimplemented: ${unimplemented}`);
console.log(`  Not found in test results: ${notFound}`);