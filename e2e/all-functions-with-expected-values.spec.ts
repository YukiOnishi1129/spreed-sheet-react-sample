import { generateTestsForCategory } from './generate-function-tests';

// 各カテゴリのテストデータをインポート
import { basicOperatorsTests } from '../src/data/individualTests/00-basic-operators';
import { mathTrigonometryTests } from '../src/data/individualTests/01-math-trigonometry';
import { statisticsTests } from '../src/data/individualTests/02-statistics';
import { textTests } from '../src/data/individualTests/03-text';
import { dateTimeTests } from '../src/data/individualTests/04-date-time';
import { logicalTests } from '../src/data/individualTests/05-logical';
import { lookupReferenceTests } from '../src/data/individualTests/06-lookup-reference';
import { financialTests } from '../src/data/individualTests/07-financial';
import { matrixTests } from '../src/data/individualTests/08-matrix';
import { informationTests } from '../src/data/individualTests/09-information';
import { databaseTests } from '../src/data/individualTests/10-database';
import { engineeringTests } from '../src/data/individualTests/11-engineering';
import { dynamicArraysTests } from '../src/data/individualTests/12-dynamic-arrays';
import { cubeTests } from '../src/data/individualTests/13-cube';
import { webImportTests } from '../src/data/individualTests/14-web-import';
import { othersTests } from '../src/data/individualTests/15-others';

// 各カテゴリのテストを生成
generateTestsForCategory('00. 基本演算子', basicOperatorsTests);
generateTestsForCategory('01. 数学', mathTrigonometryTests);
generateTestsForCategory('01. 数学・三角', mathTrigonometryTests);
generateTestsForCategory('02. 統計', statisticsTests);
generateTestsForCategory('03. テキスト', textTests);
generateTestsForCategory('04. 日付', dateTimeTests);
generateTestsForCategory('05. 論理', logicalTests);
generateTestsForCategory('06. 検索・参照', lookupReferenceTests);
generateTestsForCategory('07. 財務', financialTests);
generateTestsForCategory('08. 行列', matrixTests);
generateTestsForCategory('09. 情報', informationTests);
generateTestsForCategory('10. データベース', databaseTests);
generateTestsForCategory('11. エンジニアリング', engineeringTests);
generateTestsForCategory('12. 動的配列', dynamicArraysTests);
generateTestsForCategory('13. CUBE', cubeTests);
generateTestsForCategory('14. Web/インポート', webImportTests);
generateTestsForCategory('15. その他', othersTests);