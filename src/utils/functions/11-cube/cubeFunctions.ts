// キューブ（OLAP）関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// CUBEVALUE関数の実装（キューブから集計値を返す）
export const CUBEVALUE: CustomFormula = {
  name: 'CUBEVALUE',
  pattern: /CUBEVALUE\(([^,]+)(?:,\s*(.+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, connectionRef, membersStr] = matches;
    
    try {
      const connection = getCellValue(connectionRef.trim(), context)?.toString() ?? connectionRef.trim();
      
      // 実際のOLAP接続は実装せず、ダミーデータを返す
      if (!connection || connection === '') {
        return FormulaError.NAME;
      }
      
      // メンバー式の解析（簡易実装）
      const members = membersStr ? membersStr.split(',').map(m => m.trim()) : [];
      
      // ダミーのキューブデータ
      const cubeData: { [key: string]: number } = {
        '[Sales]': 1000000,
        '[Sales].[2023]': 500000,
        '[Sales].[2024]': 500000,
        '[Sales].[2023].[Q1]': 120000,
        '[Sales].[2023].[Q2]': 130000,
        '[Sales].[2023].[Q3]': 125000,
        '[Sales].[2023].[Q4]': 125000,
        '[Profit]': 200000,
        '[Profit].[2023]': 100000,
        '[Profit].[2024]': 100000,
        '[Customers]': 50000,
        '[Products]': 1500,
      };
      
      // メンバーの組み合わせからキーを生成
      let key = members.join('.');
      if (!key) key = '[Sales]'; // デフォルト
      
      // キューブデータから値を取得
      if (cubeData[key] !== undefined) {
        return cubeData[key];
      }
      
      // 複数メンバーの場合は乗算（簡易実装）
      let result = 1;
      for (const member of members) {
        const cleanMember = member.replace(/^["']|["']$/g, '');
        if (cubeData[cleanMember] !== undefined) {
          result *= cubeData[cleanMember];
        }
      }
      
      return result > 1 ? result / 1000 : FormulaError.NA;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CUBEMEMBER関数の実装（キューブのメンバーを返す）
export const CUBEMEMBER: CustomFormula = {
  name: 'CUBEMEMBER',
  pattern: /CUBEMEMBER\(([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, connectionRef, memberExprRef, captionRef] = matches;
    
    try {
      const connection = getCellValue(connectionRef.trim(), context)?.toString() ?? connectionRef.trim();
      const memberExpr = getCellValue(memberExprRef.trim(), context)?.toString() ?? memberExprRef.trim();
      const caption = captionRef ? getCellValue(captionRef.trim(), context)?.toString() ?? captionRef.trim() : null;
      
      if (!connection || connection === '') {
        return FormulaError.NAME;
      }
      
      // メンバー式をクリーンアップ
      const cleanExpr = memberExpr.replace(/^["']|["']$/g, '');
      
      // キャプションが指定されている場合はそれを返す
      if (caption) {
        return caption.replace(/^["']|["']$/g, '');
      }
      
      // ダミーのメンバー階層
      const memberHierarchy: { [key: string]: string } = {
        '[Time]': 'Time',
        '[Time].[2023]': '2023',
        '[Time].[2024]': '2024',
        '[Time].[2023].[Q1]': 'Q1 2023',
        '[Time].[2023].[Q2]': 'Q2 2023',
        '[Time].[2023].[Q3]': 'Q3 2023',
        '[Time].[2023].[Q4]': 'Q4 2023',
        '[Geography]': 'Geography',
        '[Geography].[USA]': 'USA',
        '[Geography].[Europe]': 'Europe',
        '[Geography].[Asia]': 'Asia',
        '[Product]': 'Product',
        '[Product].[Electronics]': 'Electronics',
        '[Product].[Clothing]': 'Clothing',
        '[Product].[Food]': 'Food',
        '[Measures].[Sales]': 'Sales',
        '[Measures].[Profit]': 'Profit',
        '[Measures].[Units]': 'Units Sold',
      };
      
      // メンバーの表示名を返す
      return memberHierarchy[cleanExpr] || cleanExpr;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CUBESET関数の実装（キューブのメンバーセットを定義）
export const CUBESET: CustomFormula = {
  name: 'CUBESET',
  pattern: /CUBESET\(([^,]+),\s*([^,)]+)(?:,\s*([^,)]+))?(?:,\s*([^,)]+))?(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, connectionRef, , captionRef] = matches;
    
    try {
      const connection = getCellValue(connectionRef.trim(), context)?.toString() ?? connectionRef.trim();
      // const setExpr = getCellValue(setExprRef.trim(), context)?.toString() ?? setExprRef.trim();
      const caption = captionRef ? getCellValue(captionRef.trim(), context)?.toString() ?? captionRef.trim() : 'Set';
      
      if (!connection || connection === '') {
        return FormulaError.NAME;
      }
      
      // キャプションを返す（実際のセット定義は内部的に保持される想定）
      return caption.replace(/^["']|["']$/g, '');
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CUBESETCOUNT関数の実装（セット内のアイテム数を返す）
export const CUBESETCOUNT: CustomFormula = {
  name: 'CUBESETCOUNT',
  pattern: /CUBESETCOUNT\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, setRef] = matches;
    
    try {
      const set = getCellValue(setRef.trim(), context)?.toString() ?? setRef.trim();
      
      // ダミーのセットカウント
      const setCounts: { [key: string]: number } = {
        'Time Members': 7,
        'Geography Members': 4,
        'Product Members': 4,
        'All Members': 15,
        'Set': 5, // デフォルト
      };
      
      return setCounts[set] || 1;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CUBERANKEDMEMBER関数の実装（n番目のメンバーを返す）
export const CUBERANKEDMEMBER: CustomFormula = {
  name: 'CUBERANKEDMEMBER',
  pattern: /CUBERANKEDMEMBER\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, connectionRef, , rankRef] = matches;
    
    try {
      const connection = getCellValue(connectionRef.trim(), context)?.toString() ?? connectionRef.trim();
      // const setExpr = getCellValue(setExprRef.trim(), context)?.toString() ?? setExprRef.trim();
      const rank = parseInt(getCellValue(rankRef.trim(), context)?.toString() ?? rankRef.trim());
      
      if (!connection || connection === '') {
        return FormulaError.NAME;
      }
      
      if (isNaN(rank) || rank < 1) {
        return FormulaError.VALUE;
      }
      
      // ダミーのランク付きメンバー
      const rankedMembers: string[] = [
        'USA - $500,000',
        'Europe - $300,000',
        'Asia - $200,000',
        'Canada - $150,000',
        'Australia - $100,000',
      ];
      
      if (rank > rankedMembers.length) {
        return FormulaError.NA;
      }
      
      return rankedMembers[rank - 1];
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CUBEMEMBERPROPERTY関数の実装（メンバーのプロパティを返す）
export const CUBEMEMBERPROPERTY: CustomFormula = {
  name: 'CUBEMEMBERPROPERTY',
  pattern: /CUBEMEMBERPROPERTY\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, connectionRef, memberExprRef, propertyRef] = matches;
    
    try {
      const connection = getCellValue(connectionRef.trim(), context)?.toString() ?? connectionRef.trim();
      const memberExpr = getCellValue(memberExprRef.trim(), context)?.toString() ?? memberExprRef.trim();
      const property = getCellValue(propertyRef.trim(), context)?.toString() ?? propertyRef.trim();
      
      if (!connection || connection === '') {
        return FormulaError.NAME;
      }
      
      const cleanProperty = property.replace(/^["']|["']$/g, '').toUpperCase();
      
      // ダミーのプロパティデータ
      const memberProperties: { [key: string]: { [prop: string]: string } } = {
        '[Geography].[USA]': {
          'MEMBER_CAPTION': 'United States',
          'MEMBER_NAME': 'USA',
          'MEMBER_UNIQUE_NAME': '[Geography].[USA]',
          'LEVEL_NUMBER': '1',
          'PARENT_UNIQUE_NAME': '[Geography]',
        },
        '[Time].[2023].[Q1]': {
          'MEMBER_CAPTION': 'Q1 2023',
          'MEMBER_NAME': 'Q1',
          'MEMBER_UNIQUE_NAME': '[Time].[2023].[Q1]',
          'LEVEL_NUMBER': '2',
          'PARENT_UNIQUE_NAME': '[Time].[2023]',
        },
      };
      
      const cleanMember = memberExpr.replace(/^["']|["']$/g, '');
      
      if (memberProperties[cleanMember] && memberProperties[cleanMember][cleanProperty]) {
        return memberProperties[cleanMember][cleanProperty];
      }
      
      // デフォルトプロパティ
      switch (cleanProperty) {
        case 'MEMBER_CAPTION':
          return cleanMember.split('.').pop()?.replace(/[\[\]]/g, '') || cleanMember;
        case 'MEMBER_NAME':
          return cleanMember.split('.').pop()?.replace(/[\[\]]/g, '') || cleanMember;
        case 'MEMBER_UNIQUE_NAME':
          return cleanMember;
        case 'LEVEL_NUMBER':
          return (cleanMember.split('.').length - 1).toString();
        default:
          return FormulaError.NA;
      }
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// CUBEKPIMEMBER関数の実装（KPIプロパティを返す）
export const CUBEKPIMEMBER: CustomFormula = {
  name: 'CUBEKPIMEMBER',
  pattern: /CUBEKPIMEMBER\(([^,]+),\s*([^,]+),\s*([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, connectionRef, kpiNameRef, kpiPropertyRef] = matches;
    
    try {
      const connection = getCellValue(connectionRef.trim(), context)?.toString() ?? connectionRef.trim();
      const kpiName = getCellValue(kpiNameRef.trim(), context)?.toString() ?? kpiNameRef.trim();
      const kpiProperty = parseInt(getCellValue(kpiPropertyRef.trim(), context)?.toString() ?? kpiPropertyRef.trim());
      
      if (!connection || connection === '') {
        return FormulaError.NAME;
      }
      
      const cleanKpiName = kpiName.replace(/^["']|["']$/g, '');
      
      // KPIプロパティタイプ
      // 1 = KPI値, 2 = KPI目標, 3 = KPIステータス, 4 = KPIトレンド
      // 5 = KPI重み, 6 = KPI現在のタイムメンバー
      
      const kpiData: { [key: string]: { [prop: number]: string | number } } = {
        'Sales KPI': {
          1: 1000000,  // 値
          2: 1200000,  // 目標
          3: 0.83,     // ステータス (値/目標)
          4: 1,        // トレンド (1=上昇)
          5: 1.0,      // 重み
          6: '[Time].[2024].[Q1]', // タイムメンバー
        },
        'Customer Satisfaction': {
          1: 4.2,      // 値
          2: 4.5,      // 目標
          3: 0.93,     // ステータス
          4: 0,        // トレンド (0=横ばい)
          5: 0.8,      // 重み
          6: '[Time].[2024].[Q1]',
        },
      };
      
      if (kpiData[cleanKpiName] && kpiData[cleanKpiName][kpiProperty] !== undefined) {
        return kpiData[cleanKpiName][kpiProperty];
      }
      
      return FormulaError.NA;
    } catch {
      return FormulaError.VALUE;
    }
  }
};