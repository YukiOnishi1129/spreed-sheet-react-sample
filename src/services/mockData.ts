import type { ExcelFunctionResponse } from '../types/spreadsheet';

// デモ用のモック関数集
export const mockFunctions: { [key: string]: ExcelFunctionResponse } = {
  'sum': {
    function_name: 'SUM',
    description: '範囲内の数値の合計を計算します',
    syntax: 'SUM(range1, [range2], ...)',
    category: '数学関数',
    spreadsheet_data: [
      [
        { v: "月", ct: { t: "s" }, bg: "#E3F2FD" },
        { v: "売上", ct: { t: "s" }, bg: "#E3F2FD" },
        { v: "費用", ct: { t: "s" }, bg: "#E3F2FD" },
        { v: "利益", ct: { t: "s" }, bg: "#E3F2FD" },
        null, null, null, null
      ],
      [
        { v: "1月", ct: { t: "s" } },
        { v: 100000, ct: { t: "n" } },
        { v: 30000, ct: { t: "n" } },
        { v: 70000, f: "=B2-C2", bg: "#E8F5E8", fc: "#2E7D32" },
        null, null, null, null
      ],
      [
        { v: "2月", ct: { t: "s" } },
        { v: 120000, ct: { t: "n" } },
        { v: 35000, ct: { t: "n" } },
        { v: 85000, f: "=B3-C3", bg: "#E8F5E8", fc: "#2E7D32" },
        null, null, null, null
      ],
      [
        { v: "3月", ct: { t: "s" } },
        { v: 150000, ct: { t: "n" } },
        { v: 40000, ct: { t: "n" } },
        { v: 110000, f: "=B4-C4", bg: "#E8F5E8", fc: "#2E7D32" },
        null, null, null, null
      ],
      [
        { v: "合計", ct: { t: "s" }, bg: "#E8F5E8" },
        { v: 370000, f: "=SUM(B2:B4)", bg: "#FFE0B2", fc: "#D84315" },
        { v: 105000, f: "=SUM(C2:C4)", bg: "#FFE0B2", fc: "#D84315" },
        { v: 265000, f: "=SUM(D2:D4)", bg: "#FFE0B2", fc: "#D84315" },
        null, null, null, null
      ],
      [null, null, null, null, null, null, null, null],
      [null, null, null, null, null, null, null, null],
      [null, null, null, null, null, null, null, null]
    ],
    examples: ['=SUM(A1:A10)', '=SUM(A1,B1,C1)', '=SUM(A1:A5,C1:C5)']
  },
  'vlookup': {
    function_name: 'VLOOKUP',
    description: 'テーブルの左端の列で値を検索し、同じ行の指定した列から値を返します',
    syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    category: '検索関数',
    spreadsheet_data: [
      [
        { v: "商品コード", ct: { t: "s" }, bg: "#E1F5FE" },
        { v: "商品名", ct: { t: "s" }, bg: "#E1F5FE" },
        { v: "価格", ct: { t: "s" }, bg: "#E1F5FE" },
        { v: "検索コード", ct: { t: "s" }, bg: "#FFF8E1" },
        { v: "検索結果", ct: { t: "s" }, bg: "#FFF8E1" },
        null, null, null
      ],
      [
        { v: "P001", ct: { t: "s" }, bg: "#F0F4C3" },
        { v: "ノートPC", ct: { t: "s" }, bg: "#F0F4C3" },
        { v: 80000, ct: { t: "n" }, bg: "#F0F4C3" },
        { v: "P002", ct: { t: "s" }, bg: "#FFECB3" },
        { v: null, f: "=VLOOKUP(D2,A2:C4,2,0)", bg: "#FFE0B2", fc: "#D84315" },
        null, null, null
      ],
      [
        { v: "P002", ct: { t: "s" }, bg: "#F0F4C3" },
        { v: "タブレット", ct: { t: "s" }, bg: "#F0F4C3" },
        { v: 45000, ct: { t: "n" }, bg: "#F0F4C3" },
        { v: "P003", ct: { t: "s" }, bg: "#FFECB3" },
        { v: null, f: "=VLOOKUP(D3,A2:C4,2,0)", bg: "#FFE0B2", fc: "#D84315" },
        null, null, null
      ],
      [
        { v: "P003", ct: { t: "s" }, bg: "#F0F4C3" },
        { v: "スマートフォン", ct: { t: "s" }, bg: "#F0F4C3" },
        { v: 65000, ct: { t: "n" }, bg: "#F0F4C3" },
        null, null, null, null, null
      ],
      [null, null, null, null, null, null, null, null],
      [null, null, null, null, null, null, null, null],
      [null, null, null, null, null, null, null, null],
      [null, null, null, null, null, null, null, null]
    ],
    examples: ['=VLOOKUP("P001",A2:C4,2,0)', '=VLOOKUP(D2,A2:C4,3,0)']
  },
  'if': {
    function_name: 'IF',
    description: '条件に基づいて異なる値を返します',
    syntax: 'IF(logical_test, value_if_true, value_if_false)',
    category: '論理関数',
    spreadsheet_data: [
      [
        { v: "学生名", ct: { t: "s" }, bg: "#E8EAF6" },
        { v: "点数", ct: { t: "s" }, bg: "#E8EAF6" },
        { v: "判定", ct: { t: "s" }, bg: "#E1F5FE" },
        { v: "ランク", ct: { t: "s" }, bg: "#F3E5F5" },
        null, null, null, null
      ],
      [
        { v: "田中", ct: { t: "s" } },
        { v: 85, ct: { t: "n" } },
        { v: "合格", f: '=IF(B2>=60,"合格","不合格")', bg: "#4CAF50", fc: "#FFFFFF",  },
        { v: "A", f: '=IF(B2>=90,"S",IF(B2>=80,"A",IF(B2>=70,"B","C")))', bg: "#9C27B0", fc: "#FFFFFF",  },
        null, null, null, null
      ],
      [
        { v: "佐藤", ct: { t: "s" } },
        { v: 45, ct: { t: "n" } },
        { v: "不合格", f: '=IF(B3>=60,"合格","不合格")', bg: "#4CAF50", fc: "#FFFFFF",  },
        { v: "C", f: '=IF(B3>=90,"S",IF(B3>=80,"A",IF(B3>=70,"B","C")))', bg: "#9C27B0", fc: "#FFFFFF",  },
        null, null, null, null
      ],
      [
        { v: "鈴木", ct: { t: "s" } },
        { v: 92, ct: { t: "n" } },
        { v: "合格", f: '=IF(B4>=60,"合格","不合格")', bg: "#4CAF50", fc: "#FFFFFF",  },
        { v: "S", f: '=IF(B4>=90,"S",IF(B4>=80,"A",IF(B4>=70,"B","C")))', bg: "#9C27B0", fc: "#FFFFFF",  },
        null, null, null, null
      ],
      [null, null, null, null, null, null, null, null],
      [null, null, null, null, null, null, null, null],
      [null, null, null, null, null, null, null, null],
      [null, null, null, null, null, null, null, null]
    ],
    examples: ['=IF(A1>10,"大","小")', '=IF(B1="",0,B1*2)']
  }
};

// デモ用：クエリに基づいて適切な関数を返すモック関数
export const getMockFunction = (query: string): ExcelFunctionResponse => {
  const lowerQuery = query.toLowerCase();
  
  if (lowerQuery.includes('合計') || lowerQuery.includes('sum')) {
    return mockFunctions.sum;
  } else if (lowerQuery.includes('検索') || lowerQuery.includes('vlookup') || lowerQuery.includes('lookup')) {
    return mockFunctions.vlookup;
  } else if (lowerQuery.includes('条件') || lowerQuery.includes('if') || lowerQuery.includes('判定')) {
    return mockFunctions.if;
  } else {
    // ランダムに関数を選択
    const functions = Object.values(mockFunctions);
    return functions[Math.floor(Math.random() * functions.length)];
  }
};