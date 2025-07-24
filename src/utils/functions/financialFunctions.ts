// 財務関数の定義
import { type FunctionDefinition } from '../types';
import { COLOR_SCHEMES } from '../colorSchemes';

export const FINANCIAL_FUNCTIONS: Record<string, FunctionDefinition> = {
  // 基本的な財務関数
  PV: {
    name: 'PV',
    syntax: 'PV(rate, nper, pmt, [fv], [type])',
    params: [
      { name: 'rate', desc: '各期の利率' },
      { name: 'nper', desc: '支払い回数' },
      { name: 'pmt', desc: '各期の支払額' },
      { name: 'fv', desc: '将来価値', optional: true },
      { name: 'type', desc: '支払期日（0=期末、1=期初）', optional: true }
    ],
    description: '現在価値を計算します',
    category: 'financial',
    examples: ['=PV(0.08/12,12*20,-1000)', '=PV(0.05,10,-500,10000)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  FV: {
    name: 'FV',
    syntax: 'FV(rate, nper, pmt, [pv], [type])',
    params: [
      { name: 'rate', desc: '各期の利率' },
      { name: 'nper', desc: '支払い回数' },
      { name: 'pmt', desc: '各期の支払額' },
      { name: 'pv', desc: '現在価値', optional: true },
      { name: 'type', desc: '支払期日（0=期末、1=期初）', optional: true }
    ],
    description: '将来価値を計算します',
    category: 'financial',
    examples: ['=FV(0.08/12,12*20,-1000)', '=FV(0.05,10,-500,-10000)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  PMT: {
    name: 'PMT',
    syntax: 'PMT(rate, nper, pv, [fv], [type])',
    params: [
      { name: 'rate', desc: '各期の利率' },
      { name: 'nper', desc: '支払い回数' },
      { name: 'pv', desc: '現在価値' },
      { name: 'fv', desc: '将来価値', optional: true },
      { name: 'type', desc: '支払期日（0=期末、1=期初）', optional: true }
    ],
    description: '定期支払額を計算します',
    category: 'financial',
    examples: ['=PMT(0.08/12,12*20,100000)', '=PMT(0.05,10,10000)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  PPMT: {
    name: 'PPMT',
    syntax: 'PPMT(rate, per, nper, pv, [fv], [type])',
    params: [
      { name: 'rate', desc: '各期の利率' },
      { name: 'per', desc: '対象期' },
      { name: 'nper', desc: '支払い回数' },
      { name: 'pv', desc: '現在価値' },
      { name: 'fv', desc: '将来価値', optional: true },
      { name: 'type', desc: '支払期日（0=期末、1=期初）', optional: true }
    ],
    description: '元金返済額を計算します',
    category: 'financial',
    examples: ['=PPMT(0.08/12,1,12*20,100000)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  IPMT: {
    name: 'IPMT',
    syntax: 'IPMT(rate, per, nper, pv, [fv], [type])',
    params: [
      { name: 'rate', desc: '各期の利率' },
      { name: 'per', desc: '対象期' },
      { name: 'nper', desc: '支払い回数' },
      { name: 'pv', desc: '現在価値' },
      { name: 'fv', desc: '将来価値', optional: true },
      { name: 'type', desc: '支払期日（0=期末、1=期初）', optional: true }
    ],
    description: '利息支払額を計算します',
    category: 'financial',
    examples: ['=IPMT(0.08/12,1,12*20,100000)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  RATE: {
    name: 'RATE',
    syntax: 'RATE(nper, pmt, pv, [fv], [type], [guess])',
    params: [
      { name: 'nper', desc: '支払い回数' },
      { name: 'pmt', desc: '各期の支払額' },
      { name: 'pv', desc: '現在価値' },
      { name: 'fv', desc: '将来価値', optional: true },
      { name: 'type', desc: '支払期日（0=期末、1=期初）', optional: true },
      { name: 'guess', desc: '推定値', optional: true }
    ],
    description: '利率を計算します',
    category: 'financial',
    examples: ['=RATE(48,-200,8000)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  NPER: {
    name: 'NPER',
    syntax: 'NPER(rate, pmt, pv, [fv], [type])',
    params: [
      { name: 'rate', desc: '各期の利率' },
      { name: 'pmt', desc: '各期の支払額' },
      { name: 'pv', desc: '現在価値' },
      { name: 'fv', desc: '将来価値', optional: true },
      { name: 'type', desc: '支払期日（0=期末、1=期初）', optional: true }
    ],
    description: '支払回数を計算します',
    category: 'financial',
    examples: ['=NPER(0.08/12,-1000,100000)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  NPV: {
    name: 'NPV',
    syntax: 'NPV(rate, value1, [value2], ...)',
    params: [
      { name: 'rate', desc: '割引率' },
      { name: 'value1', desc: '1期目のキャッシュフロー' },
      { name: 'value2', desc: '2期目のキャッシュフロー', optional: true }
    ],
    description: '正味現在価値を計算します',
    category: 'financial',
    examples: ['=NPV(0.1,-10000,3000,4200,6800)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  XNPV: {
    name: 'XNPV',
    syntax: 'XNPV(rate, values, dates)',
    params: [
      { name: 'rate', desc: '割引率' },
      { name: 'values', desc: 'キャッシュフローの範囲' },
      { name: 'dates', desc: '日付の範囲' }
    ],
    description: '不定期なキャッシュフローの正味現在価値を計算します',
    category: 'financial',
    examples: ['=XNPV(0.09,B1:B5,A1:A5)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  IRR: {
    name: 'IRR',
    syntax: 'IRR(values, [guess])',
    params: [
      { name: 'values', desc: 'キャッシュフローの範囲' },
      { name: 'guess', desc: '推定値', optional: true }
    ],
    description: '内部収益率を計算します',
    category: 'financial',
    examples: ['=IRR(A1:A5)', '=IRR(A1:A10,0.1)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  XIRR: {
    name: 'XIRR',
    syntax: 'XIRR(values, dates, [guess])',
    params: [
      { name: 'values', desc: 'キャッシュフローの範囲' },
      { name: 'dates', desc: '日付の範囲' },
      { name: 'guess', desc: '推定値', optional: true }
    ],
    description: '不定期なキャッシュフローの内部収益率を計算します',
    category: 'financial',
    examples: ['=XIRR(B1:B5,A1:A5)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  MIRR: {
    name: 'MIRR',
    syntax: 'MIRR(values, finance_rate, reinvest_rate)',
    params: [
      { name: 'values', desc: 'キャッシュフローの範囲' },
      { name: 'finance_rate', desc: '財務費用率' },
      { name: 'reinvest_rate', desc: '再投資収益率' }
    ],
    description: '修正内部収益率を計算します',
    category: 'financial',
    examples: ['=MIRR(A1:A5,0.1,0.12)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  // 減価償却関数
  SLN: {
    name: 'SLN',
    syntax: 'SLN(cost, salvage, life)',
    params: [
      { name: 'cost', desc: '取得価額' },
      { name: 'salvage', desc: '残存価額' },
      { name: 'life', desc: '耐用年数' }
    ],
    description: '定額法による減価償却費を計算します',
    category: 'financial',
    examples: ['=SLN(30000,7500,10)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  SYD: {
    name: 'SYD',
    syntax: 'SYD(cost, salvage, life, per)',
    params: [
      { name: 'cost', desc: '取得価額' },
      { name: 'salvage', desc: '残存価額' },
      { name: 'life', desc: '耐用年数' },
      { name: 'per', desc: '対象期' }
    ],
    description: '級数法による減価償却費を計算します',
    category: 'financial',
    examples: ['=SYD(30000,7500,10,1)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  DB: {
    name: 'DB',
    syntax: 'DB(cost, salvage, life, period, [month])',
    params: [
      { name: 'cost', desc: '取得価額' },
      { name: 'salvage', desc: '残存価額' },
      { name: 'life', desc: '耐用年数' },
      { name: 'period', desc: '対象期' },
      { name: 'month', desc: '第1年度の月数', optional: true }
    ],
    description: '定率法による減価償却費を計算します',
    category: 'financial',
    examples: ['=DB(1000000,100000,6,1,7)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  DDB: {
    name: 'DDB',
    syntax: 'DDB(cost, salvage, life, period, [factor])',
    params: [
      { name: 'cost', desc: '取得価額' },
      { name: 'salvage', desc: '残存価額' },
      { name: 'life', desc: '耐用年数' },
      { name: 'period', desc: '対象期' },
      { name: 'factor', desc: '償却率', optional: true }
    ],
    description: '倍額定率法による減価償却費を計算します',
    category: 'financial',
    examples: ['=DDB(2400,300,10,1)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  VDB: {
    name: 'VDB',
    syntax: 'VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])',
    params: [
      { name: 'cost', desc: '取得価額' },
      { name: 'salvage', desc: '残存価額' },
      { name: 'life', desc: '耐用年数' },
      { name: 'start_period', desc: '開始期' },
      { name: 'end_period', desc: '終了期' },
      { name: 'factor', desc: '償却率', optional: true },
      { name: 'no_switch', desc: '定額法に切り替えない場合TRUE', optional: true }
    ],
    description: '可変減価償却費を計算します',
    category: 'financial',
    examples: ['=VDB(2400,300,10,0,0.875,2)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  // 証券関数
  ACCRINT: {
    name: 'ACCRINT',
    syntax: 'ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])',
    params: [
      { name: 'issue', desc: '証券の発行日' },
      { name: 'first_interest', desc: '初回利払日' },
      { name: 'settlement', desc: '受渡日' },
      { name: 'rate', desc: '年利率' },
      { name: 'par', desc: '額面価格' },
      { name: 'frequency', desc: '年間利払回数' },
      { name: 'basis', desc: '日数の計算基準', optional: true },
      { name: 'calc_method', desc: '計算方式', optional: true }
    ],
    description: '定期的に利息が支払われる証券の経過利息を計算します',
    category: 'financial',
    examples: ['=ACCRINT("2008/3/1","2008/8/31","2008/5/1",0.1,1000,2,0)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  DISC: {
    name: 'DISC',
    syntax: 'DISC(settlement, maturity, pr, redemption, [basis])',
    params: [
      { name: 'settlement', desc: '受渡日' },
      { name: 'maturity', desc: '満期日' },
      { name: 'pr', desc: '価格' },
      { name: 'redemption', desc: '償還価額' },
      { name: 'basis', desc: '日数の計算基準', optional: true }
    ],
    description: '割引証券の割引率を計算します',
    category: 'financial',
    examples: ['=DISC("2007/1/25","2007/6/15",97.975,100,1)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  PRICE: {
    name: 'PRICE',
    syntax: 'PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])',
    params: [
      { name: 'settlement', desc: '受渡日' },
      { name: 'maturity', desc: '満期日' },
      { name: 'rate', desc: '年利率' },
      { name: 'yld', desc: '年利回り' },
      { name: 'redemption', desc: '額面に対する償還価額' },
      { name: 'frequency', desc: '年間利払回数' },
      { name: 'basis', desc: '日数の計算基準', optional: true }
    ],
    description: '定期的に利息が支払われる証券の価格を計算します',
    category: 'financial',
    examples: ['=PRICE("2008/2/15","2017/11/15",0.0575,0.065,100,2,0)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  YIELD: {
    name: 'YIELD',
    syntax: 'YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])',
    params: [
      { name: 'settlement', desc: '受渡日' },
      { name: 'maturity', desc: '満期日' },
      { name: 'rate', desc: '年利率' },
      { name: 'pr', desc: '価格' },
      { name: 'redemption', desc: '額面に対する償還価額' },
      { name: 'frequency', desc: '年間利払回数' },
      { name: 'basis', desc: '日数の計算基準', optional: true }
    ],
    description: '定期的に利息が支払われる証券の利回りを計算します',
    category: 'financial',
    examples: ['=YIELD("2008/2/15","2016/11/15",0.0575,95.04287,100,2,0)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  DURATION: {
    name: 'DURATION',
    syntax: 'DURATION(settlement, maturity, rate, yld, frequency, [basis])',
    params: [
      { name: 'settlement', desc: '受渡日' },
      { name: 'maturity', desc: '満期日' },
      { name: 'rate', desc: '年利率' },
      { name: 'yld', desc: '年利回り' },
      { name: 'frequency', desc: '年間利払回数' },
      { name: 'basis', desc: '日数の計算基準', optional: true }
    ],
    description: 'デュレーションを計算します',
    category: 'financial',
    examples: ['=DURATION("2008/1/1","2016/1/1",0.08,0.09,2,1)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  MDURATION: {
    name: 'MDURATION',
    syntax: 'MDURATION(settlement, maturity, rate, yld, frequency, [basis])',
    params: [
      { name: 'settlement', desc: '受渡日' },
      { name: 'maturity', desc: '満期日' },
      { name: 'rate', desc: '年利率' },
      { name: 'yld', desc: '年利回り' },
      { name: 'frequency', desc: '年間利払回数' },
      { name: 'basis', desc: '日数の計算基準', optional: true }
    ],
    description: '修正デュレーションを計算します',
    category: 'financial',
    examples: ['=MDURATION("2008/1/1","2016/1/1",0.08,0.09,2,1)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  // その他の財務関数
  CUMIPMT: {
    name: 'CUMIPMT',
    syntax: 'CUMIPMT(rate, nper, pv, start_period, end_period, type)',
    params: [
      { name: 'rate', desc: '各期の利率' },
      { name: 'nper', desc: '支払い回数' },
      { name: 'pv', desc: '現在価値' },
      { name: 'start_period', desc: '開始期' },
      { name: 'end_period', desc: '終了期' },
      { name: 'type', desc: '支払期日（0=期末、1=期初）' }
    ],
    description: '指定期間の累計利息を計算します',
    category: 'financial',
    examples: ['=CUMIPMT(0.09/12,30*12,125000,13,24,0)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  CUMPRINC: {
    name: 'CUMPRINC',
    syntax: 'CUMPRINC(rate, nper, pv, start_period, end_period, type)',
    params: [
      { name: 'rate', desc: '各期の利率' },
      { name: 'nper', desc: '支払い回数' },
      { name: 'pv', desc: '現在価値' },
      { name: 'start_period', desc: '開始期' },
      { name: 'end_period', desc: '終了期' },
      { name: 'type', desc: '支払期日（0=期末、1=期初）' }
    ],
    description: '指定期間の累計元金を計算します',
    category: 'financial',
    examples: ['=CUMPRINC(0.09/12,30*12,125000,13,24,0)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  EFFECT: {
    name: 'EFFECT',
    syntax: 'EFFECT(nominal_rate, npery)',
    params: [
      { name: 'nominal_rate', desc: '名目年利率' },
      { name: 'npery', desc: '年間複利回数' }
    ],
    description: '実効年利率を計算します',
    category: 'financial',
    examples: ['=EFFECT(0.0525,4)'],
    colorScheme: COLOR_SCHEMES.financial
  },

  NOMINAL: {
    name: 'NOMINAL',
    syntax: 'NOMINAL(effect_rate, npery)',
    params: [
      { name: 'effect_rate', desc: '実効年利率' },
      { name: 'npery', desc: '年間複利回数' }
    ],
    description: '名目年利率を計算します',
    category: 'financial',
    examples: ['=NOMINAL(0.053543,4)'],
    colorScheme: COLOR_SCHEMES.financial
  }
};