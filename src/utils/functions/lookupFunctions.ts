// 検索・参照関数の定義
import { type FunctionDefinition } from '../types';
import { COLOR_SCHEMES } from '../colorSchemes';

export const LOOKUP_FUNCTIONS: Record<string, FunctionDefinition> = {
  VLOOKUP: {
    name: 'VLOOKUP',
    syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    params: [
      { name: 'lookup_value', desc: '検索したい値が入っているセル' },
      { name: 'table_array', desc: '検索対象のデータ範囲' },
      { name: 'col_index_num', desc: '取得したいデータの列番号' },
      { name: 'range_lookup', desc: '完全一致は0またはFALSE（省略可能）', optional: true }
    ],
    description: '縦方向（列）でデータを検索し、指定した列から値を返します',
    category: 'lookup',
    examples: ['=VLOOKUP(A1,B:D,3,0)', '=VLOOKUP("商品A",A1:C10,2,FALSE)'],
    colorScheme: COLOR_SCHEMES.lookup
  }
};