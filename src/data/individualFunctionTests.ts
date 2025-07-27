import type { IndividualFunctionTest } from '../types/spreadsheet';

// Import all test categories
import { basicOperatorsTests } from './individualTests/00-basic-operators';
import { mathTrigonometryTests } from './individualTests/01-math-trigonometry';
import { statisticsTests } from './individualTests/02-statistics';
import { textTests } from './individualTests/03-text';
import { dateTimeTests } from './individualTests/04-date-time';
import { logicalTests } from './individualTests/05-logical';
import { lookupReferenceTests } from './individualTests/06-lookup-reference';
import { financialTests } from './individualTests/07-financial';
import { matrixTests } from './individualTests/08-matrix';
import { informationTests } from './individualTests/09-information';
import { databaseTests } from './individualTests/10-database';
import { engineeringTests } from './individualTests/11-engineering';
import { dynamicArraysTests } from './individualTests/12-dynamic-arrays';
import { cubeTests } from './individualTests/13-cube';
import { webImportTests } from './individualTests/14-web-import';
import { othersTests } from './individualTests/15-others';

// Combine all tests
export const allIndividualFunctionTests: IndividualFunctionTest[] = [
  ...basicOperatorsTests,
  ...mathTrigonometryTests,
  ...statisticsTests,
  ...textTests,
  ...dateTimeTests,
  ...logicalTests,
  ...lookupReferenceTests,
  ...financialTests,
  ...matrixTests,
  ...informationTests,
  ...databaseTests,
  ...engineeringTests,
  ...dynamicArraysTests,
  ...cubeTests,
  ...webImportTests,
  ...othersTests,
];

// Helper function to get test by name
export function getIndividualTestByName(name: string): IndividualFunctionTest | undefined {
  return allIndividualFunctionTests.find(test => test.name === name);
}
