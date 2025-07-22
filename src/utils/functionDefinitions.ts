// Excel関数の定義辞書

export interface FunctionParam {
  name: string;
  desc: string;
  optional?: boolean;
}

export interface FunctionDefinition {
  name: string;
  syntax: string;
  params: FunctionParam[];
  description: string;
}

export const FUNCTION_DEFINITIONS: Record<string, FunctionDefinition> = {
  // 検索関数
  HLOOKUP: {
    name: 'HLOOKUP',
    syntax: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])',
    params: [
      { name: 'lookup_value', desc: '検索したい値が入っているセル' },
      { name: 'table_array', desc: '検索対象のデータ範囲' },
      { name: 'row_index_num', desc: '取得したいデータの行番号' },
      { name: 'range_lookup', desc: '完全一致は0またはFALSE（省略可能）', optional: true }
    ],
    description: '横方向（行）でデータを検索し、指定した行から値を返します'
  },
  
  VLOOKUP: {
    name: 'VLOOKUP',
    syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    params: [
      { name: 'lookup_value', desc: '検索したい値が入っているセル' },
      { name: 'table_array', desc: '検索対象のデータ範囲' },
      { name: 'col_index_num', desc: '取得したいデータの列番号' },
      { name: 'range_lookup', desc: '完全一致は0またはFALSE（省略可能）', optional: true }
    ],
    description: '縦方向（列）でデータを検索し、指定した列から値を返します'
  },

  // エラーハンドリング関数
  IFNA: {
    name: 'IFNA',
    syntax: 'IFNA(value, value_if_na)',
    params: [
      { name: 'value', desc: '評価する値や数式' },
      { name: 'value_if_na', desc: '#N/Aエラー時に表示する代替値' }
    ],
    description: '#N/Aエラーの場合に指定した値を表示します'
  },

  IFERROR: {
    name: 'IFERROR',
    syntax: 'IFERROR(value, value_if_error)',
    params: [
      { name: 'value', desc: '評価する値や数式' },
      { name: 'value_if_error', desc: 'エラー時に表示する代替値' }
    ],
    description: 'すべてのエラーに対して指定した値を表示します'
  },

  // 論理関数
  IF: {
    name: 'IF',
    syntax: 'IF(logical_test, value_if_true, value_if_false)',
    params: [
      { name: 'logical_test', desc: '判定する条件' },
      { name: 'value_if_true', desc: '条件がTRUEの場合の値' },
      { name: 'value_if_false', desc: '条件がFALSEの場合の値' }
    ],
    description: '条件に応じて異なる値を返します'
  },

  AND: {
    name: 'AND',
    syntax: 'AND(logical1, logical2, ...)',
    params: [
      { name: 'logical1', desc: '1つ目の論理値または条件' },
      { name: 'logical2', desc: '2つ目の論理値または条件' },
      { name: '...', desc: '追加の論理値（最大255個）' }
    ],
    description: 'すべての条件がTRUEの場合にTRUEを返します'
  },

  OR: {
    name: 'OR',
    syntax: 'OR(logical1, logical2, ...)',
    params: [
      { name: 'logical1', desc: '1つ目の論理値または条件' },
      { name: 'logical2', desc: '2つ目の論理値または条件' },
      { name: '...', desc: '追加の論理値（最大255個）' }
    ],
    description: 'いずれかの条件がTRUEの場合にTRUEを返します'
  },

  // 数学関数
  SUM: {
    name: 'SUM',
    syntax: 'SUM(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '合計する最初の数値または範囲' },
      { name: 'number2', desc: '合計する追加の数値または範囲（省略可能）', optional: true }
    ],
    description: '数値の合計を計算します'
  },

  AVERAGE: {
    name: 'AVERAGE',
    syntax: 'AVERAGE(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '平均を求める最初の数値または範囲' },
      { name: 'number2', desc: '平均を求める追加の数値または範囲（省略可能）', optional: true }
    ],
    description: '数値の平均を計算します'
  },

  COUNT: {
    name: 'COUNT',
    syntax: 'COUNT(value1, [value2], ...)',
    params: [
      { name: 'value1', desc: '数値をカウントする最初の値または範囲' },
      { name: 'value2', desc: '数値をカウントする追加の値または範囲（省略可能）', optional: true }
    ],
    description: '数値が入力されているセルの個数を数えます'
  },

  MAX: {
    name: 'MAX',
    syntax: 'MAX(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '最大値を求める最初の数値または範囲' },
      { name: 'number2', desc: '最大値を求める追加の数値または範囲（省略可能）', optional: true }
    ],
    description: '最大値を返します'
  },

  MIN: {
    name: 'MIN',
    syntax: 'MIN(number1, [number2], ...)',
    params: [
      { name: 'number1', desc: '最小値を求める最初の数値または範囲' },
      { name: 'number2', desc: '最小値を求める追加の数値または範囲（省略可能）', optional: true }
    ],
    description: '最小値を返します'
  },

  // テキスト関数
  CONCATENATE: {
    name: 'CONCATENATE',
    syntax: 'CONCATENATE(text1, text2, ...)',
    params: [
      { name: 'text1', desc: '結合する最初のテキスト' },
      { name: 'text2', desc: '結合する2番目のテキスト' },
      { name: '...', desc: '結合する追加のテキスト（最大255個）' }
    ],
    description: '複数のテキストを結合します'
  },

  TEXT: {
    name: 'TEXT',
    syntax: 'TEXT(value, format_text)',
    params: [
      { name: 'value', desc: '書式を適用する数値' },
      { name: 'format_text', desc: '書式パターン（例："0,000"、"yyyy/mm/dd"）' }
    ],
    description: '数値を指定した書式の文字列に変換します'
  },

  REPT: {
    name: 'REPT',
    syntax: 'REPT(text, number_times)',
    params: [
      { name: 'text', desc: '繰り返すテキスト' },
      { name: 'number_times', desc: '繰り返し回数' }
    ],
    description: 'テキストを指定した回数繰り返します'
  },

  // 日付関数
  TODAY: {
    name: 'TODAY',
    syntax: 'TODAY()',
    params: [],
    description: '今日の日付を返します'
  },

  EDATE: {
    name: 'EDATE',
    syntax: 'EDATE(start_date, months)',
    params: [
      { name: 'start_date', desc: '開始日' },
      { name: 'months', desc: '追加する月数（負の値も可能）' }
    ],
    description: '指定した月数後の日付を返します'
  },

  YEAR: {
    name: 'YEAR',
    syntax: 'YEAR(serial_number)',
    params: [
      { name: 'serial_number', desc: '年を取得したい日付' }
    ],
    description: '日付から年を取得します'
  },

  MONTH: {
    name: 'MONTH',
    syntax: 'MONTH(serial_number)',
    params: [
      { name: 'serial_number', desc: '月を取得したい日付' }
    ],
    description: '日付から月を取得します'
  },

  DAY: {
    name: 'DAY',
    syntax: 'DAY(serial_number)',
    params: [
      { name: 'serial_number', desc: '日を取得したい日付' }
    ],
    description: '日付から日を取得します'
  }
};

/**
 * 関数名から使用されている個別の関数を抽出
 */
export function extractFunctionsFromName(functionName: string): string[] {
  const functions: string[] = [];
  
  // &で区切られた関数名を抽出
  const parts = functionName.split('&').map(part => part.trim());
  
  for (const part of parts) {
    // 各パートから関数名を抽出（英語の関数名のみ）
    const matches = part.match(/[A-Z][A-Z0-9]*/g);
    if (matches) {
      functions.push(...matches);
    }
  }
  
  // 重複を除去して返す
  return [...new Set(functions)];
}

/**
 * 使用関数の詳細説明を生成（見やすく整形）
 */
export function generateFunctionDetails(functions: string[]): string {
  let details = '';
  
  for (const func of functions) {
    const definition = FUNCTION_DEFINITIONS[func];
    if (!definition) continue;
    
    // 関数名とシンタックスを表示
    details += `${definition.syntax}\n`;
    
    // 各引数を改行して表示
    for (const param of definition.params) {
      const optional = param.optional ? '（省略可能）' : '';
      details += `- ${param.name}: ${param.desc}${optional}\n`;
    }
    
    // 関数間にスペースを追加
    details += '\n';
  }
  
  return details.trim();
}