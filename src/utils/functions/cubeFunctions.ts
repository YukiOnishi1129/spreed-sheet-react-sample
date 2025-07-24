// キューブ関数の定義
import { type FunctionDefinition } from '../types';
import { COLOR_SCHEMES } from '../colorSchemes';

export const CUBE_FUNCTIONS: Record<string, FunctionDefinition> = {
  CUBEKPIMEMBER: {
    name: 'CUBEKPIMEMBER',
    syntax: 'CUBEKPIMEMBER(connection, kpi_name, kpi_property, [caption])',
    params: [
      { name: 'connection', desc: 'キューブへの接続名' },
      { name: 'kpi_name', desc: 'KPI名' },
      { name: 'kpi_property', desc: 'KPIプロパティ' },
      { name: 'caption', desc: 'キャプション', optional: true }
    ],
    description: 'キューブからKPIプロパティを返します',
    category: 'cube',
    examples: ['=CUBEKPIMEMBER("Adventure Works","Internet Revenue","Goal")'],
    colorScheme: COLOR_SCHEMES.cube
  },

  CUBEMEMBER: {
    name: 'CUBEMEMBER',
    syntax: 'CUBEMEMBER(connection, member_expression, [caption])',
    params: [
      { name: 'connection', desc: 'キューブへの接続名' },
      { name: 'member_expression', desc: 'メンバー式' },
      { name: 'caption', desc: 'キャプション', optional: true }
    ],
    description: 'キューブからメンバーを返します',
    category: 'cube',
    examples: ['=CUBEMEMBER("Adventure Works","[Geography].[Country].&[Australia]")'],
    colorScheme: COLOR_SCHEMES.cube
  },

  CUBEMEMBERPROPERTY: {
    name: 'CUBEMEMBERPROPERTY',
    syntax: 'CUBEMEMBERPROPERTY(connection, member_expression, property)',
    params: [
      { name: 'connection', desc: 'キューブへの接続名' },
      { name: 'member_expression', desc: 'メンバー式' },
      { name: 'property', desc: 'プロパティ名' }
    ],
    description: 'キューブからメンバープロパティを返します',
    category: 'cube',
    examples: ['=CUBEMEMBERPROPERTY("Adventure Works","[Geography].[Country].&[Australia]","Caption")'],
    colorScheme: COLOR_SCHEMES.cube
  },

  CUBERANKEDMEMBER: {
    name: 'CUBERANKEDMEMBER',
    syntax: 'CUBERANKEDMEMBER(connection, set_expression, rank, [caption])',
    params: [
      { name: 'connection', desc: 'キューブへの接続名' },
      { name: 'set_expression', desc: 'セット式' },
      { name: 'rank', desc: 'ランク' },
      { name: 'caption', desc: 'キャプション', optional: true }
    ],
    description: 'セットのn番目のメンバーを返します',
    category: 'cube',
    examples: ['=CUBERANKEDMEMBER("Adventure Works",CUBESET("Adventure Works","[Product].[Product Categories].[Category].Members"),3)'],
    colorScheme: COLOR_SCHEMES.cube
  },

  CUBESET: {
    name: 'CUBESET',
    syntax: 'CUBESET(connection, set_expression, [caption], [sort_order], [sort_by])',
    params: [
      { name: 'connection', desc: 'キューブへの接続名' },
      { name: 'set_expression', desc: 'セット式' },
      { name: 'caption', desc: 'キャプション', optional: true },
      { name: 'sort_order', desc: '並べ替え順序', optional: true },
      { name: 'sort_by', desc: '並べ替え基準', optional: true }
    ],
    description: 'キューブにセットを定義します',
    category: 'cube',
    examples: ['=CUBESET("Adventure Works","[Product].[Product Categories].[Category].Members")'],
    colorScheme: COLOR_SCHEMES.cube
  },

  CUBESETCOUNT: {
    name: 'CUBESETCOUNT',
    syntax: 'CUBESETCOUNT(set)',
    params: [
      { name: 'set', desc: 'CUBESETで定義されたセット' }
    ],
    description: 'セット内のアイテム数を返します',
    category: 'cube',
    examples: ['=CUBESETCOUNT(CUBESET("Adventure Works","[Product].[Product Categories].[Category].Members"))'],
    colorScheme: COLOR_SCHEMES.cube
  },

  CUBEVALUE: {
    name: 'CUBEVALUE',
    syntax: 'CUBEVALUE(connection, [member_expression1], [member_expression2], ...)',
    params: [
      { name: 'connection', desc: 'キューブへの接続名' },
      { name: 'member_expression1', desc: '1つ目のメンバー式', optional: true },
      { name: 'member_expression2', desc: '2つ目のメンバー式', optional: true }
    ],
    description: 'キューブから集計値を返します',
    category: 'cube',
    examples: ['=CUBEVALUE("Adventure Works","[Measures].[Internet Sales Amount]","[Date].[Calendar Year].&[2004]")'],
    colorScheme: COLOR_SCHEMES.cube
  }
};