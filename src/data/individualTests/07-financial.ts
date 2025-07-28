import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const financialTests: IndividualFunctionTest[] = [
  {
    name: 'PMT',
    category: '07. 財務',
    description: 'ローン定期支払額',
    data: [
      ['元本', '月利率', '期間(月)', '月額支払額'],
      [100000, 0.005, 12, '=PMT(B2,C2,-A2)']
    ],
    expectedValues: { 'D2': 8606.64 }
  },
  {
    name: 'FV',
    category: '07. 財務',
    description: '将来価値',
    data: [
      ['定期支払額', '月利率', '期間', '将来価値'],
      [1000, 0.004167, 10, '=FV(B2,C2,-A2)']
    ],
    expectedValues: { 'D2': 10189.60 }
  },
  {
    name: 'PV',
    category: '07. 財務',
    description: '現在価値',
    data: [
      ['定期支払額', '月利率', '期間', '現在価値'],
      [1000, 0.004167, 10, '=PV(B2,C2,-A2)']
    ],
    expectedValues: { 'D2': 9774.60 }
  },
  {
    name: 'RATE',
    category: '07. 財務',
    description: '利率',
    data: [
      ['期間', '支払額', '現在価値', '利率'],
      [12, -1000, 10000, '=RATE(A2,B2,C2)*12']
    ],
    expectedValues: { 'D2': 0.0253 }
  },
  {
    name: 'NPER',
    category: '07. 財務',
    description: '支払回数',
    data: [
      ['利率', '支払額', '現在価値', '期間'],
      [0.05, -1000, 10000, '=NPER(A2/12,B2,C2)']
    ],
    expectedValues: { 'D2': 10.475 }
  },
  {
    name: 'NPV',
    category: '07. 財務',
    description: '正味現在価値',
    data: [
      ['割引率', 'CF1', 'CF2', 'CF3', 'NPV'],
      [0.1, -10000, 3000, 4200, '=NPV(A2,B2:D2)']
    ],
    expectedValues: { 'E2': -4131.53 }
  },
  {
    name: 'IRR',
    category: '07. 財務',
    description: '内部収益率',
    data: [
      ['初期投資', 'CF1', 'CF2', 'CF3', 'IRR'],
      [-10000, 3000, 4200, 6800, '=IRR(A2:D2)']
    ],
    expectedValues: { 'E2': 0.1212 }
  },
  {
    name: 'XNPV',
    category: '07. 財務',
    description: '不定期のキャッシュフローの正味現在価値',
    data: [
      ['割引率', 0.1, '', ''],
      ['日付', 'キャッシュフロー', '', 'XNPV'],
      ['2024/1/1', -10000, '', '=XNPV(B1,B3:B6,A3:A6)'],
      ['2024/4/1', 3000, '', ''],
      ['2024/7/1', 4200, '', ''],
      ['2024/10/1', 6800, '', '']
    ],
    expectedValues: { 'D3': 3554.00 }
  },
  {
    name: 'XIRR',
    category: '07. 財務',
    description: '不定期のキャッシュフローの内部収益率',
    data: [
      ['日付', 'キャッシュフロー', 'XIRR'],
      ['2024/1/1', -10000, '=XIRR(B2:B5,A2:A5)'],
      ['2024/4/1', 3000, ''],
      ['2024/7/1', 4200, ''],
      ['2024/10/1', 6800, '']
    ],
    expectedValues: { 'C2': 0.5378 }
  },
  {
    name: 'IPMT',
    category: '07. 財務',
    description: '利息部分の支払額',
    data: [
      ['利率', '期', '期間', '現在価値', '利息支払額'],
      [0.06, 1, 12, -100000, '=IPMT(A2/12,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 500.00 }
  },
  {
    name: 'PPMT',
    category: '07. 財務',
    description: '元本部分の支払額',
    data: [
      ['利率', '期', '期間', '現在価値', '元本支払額'],
      [0.06, 1, 12, -100000, '=PPMT(A2/12,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 8110.66 }
  },
  {
    name: 'MIRR',
    category: '07. 財務',
    description: '修正内部収益率',
    data: [
      ['CF0', 'CF1', 'CF2', 'CF3', '投資利率', '再投資利率', 'MIRR'],
      [-10000, 3000, 4200, 6800, 0.1, 0.12, '=MIRR(A2:D2,E2,F2)']
    ],
    expectedValues: { 'G2': 0.1195 }
  },
  {
    name: 'SLN',
    category: '07. 財務',
    description: '定額法による減価償却',
    data: [
      ['取得価額', '残存価額', '耐用年数', '減価償却費'],
      [100000, 10000, 5, '=SLN(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 18000 }
  },
  {
    name: 'SYD',
    category: '07. 財務',
    description: '級数法による減価償却',
    data: [
      ['取得価額', '残存価額', '耐用年数', '期', '減価償却費'],
      [100000, 10000, 5, 1, '=SYD(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 30000 }
  },
  {
    name: 'DDB',
    category: '07. 財務',
    description: '倍率法による減価償却',
    data: [
      ['取得価額', '残存価額', '耐用年数', '期', '倍率', '減価償却費'],
      [100000, 10000, 5, 1, 2, '=DDB(A2,B2,C2,D2,E2)']
    ],
    expectedValues: { 'F2': 40000 }
  },
  {
    name: 'DB',
    category: '07. 財務',
    description: '定率法による減価償却',
    data: [
      ['取得価額', '残存価額', '耐用年数', '期', '月', '減価償却費'],
      [100000, 10000, 5, 1, 12, '=DB(A2,B2,C2,D2,E2)']
    ],
    expectedValues: { 'F2': 36900 }
  },
  {
    name: 'VDB',
    category: '07. 財務',
    description: '倍率法減価償却（期間指定）',
    data: [
      ['取得価額', '残存価額', '耐用年数', '開始期', '終了期', '倍率', '切替なし', '償却費'],
      [100000, 10000, 5, 0, 1, 2, 'FALSE', '=VDB(A2,B2,C2,D2,E2,F2,G2)']
    ],
    expectedValues: { 'H2': 40000 }
  },
  {
    name: 'ACCRINT',
    category: '07. 財務',
    description: '定期利息証券の未収利息額',
    data: [
      ['発行日', '初回利払日', '決済日', '利率', '額面', '頻度', '基準', '未収利息'],
      ['2024/1/1', '2024/7/1', '2024/4/1', 0.05, 1000, 2, 0, '=ACCRINT(A2,B2,C2,D2,E2,F2,G2)']
    ],
    expectedValues: { 'H2': 12.50 }
  },
  {
    name: 'ACCRINTM',
    category: '07. 財務',
    description: '満期一括払い証券の未収利息額',
    data: [
      ['発行日', '満期日', '利率', '額面', '基準', '未収利息'],
      ['2024/1/1', '2024/12/31', 0.05, 1000, 0, '=ACCRINTM(A2,B2,C2,D2,E2)']
    ],
    expectedValues: { 'F2': 50.00 }
  },
  {
    name: 'DISC',
    category: '07. 財務',
    description: '証券の割引率',
    data: [
      ['決済日', '満期日', '価格', '償還価格', '基準', '割引率'],
      ['2024/1/1', '2024/12/31', 95, 100, 0, '=DISC(A2,B2,C2,D2,E2)']
    ],
    expectedValues: { 'F2': 0.0503 }
  },
  {
    name: 'DURATION',
    category: '07. 財務',
    description: '定期利息証券の年間マコーレー・デュレーション',
    data: [
      ['決済日', '満期日', '利率', '利回り', '頻度', '基準', 'デュレーション'],
      ['2024/1/1', '2029/1/1', 0.05, 0.06, 2, 0, '=DURATION(A2,B2,C2,D2,E2,F2)']
    ],
    expectedValues: { 'G2': 4.438 }
  },
  {
    name: 'MDURATION',
    category: '07. 財務',
    description: '修正マコーレー・デュレーション',
    data: [
      ['決済日', '満期日', '利率', '利回り', '頻度', '基準', '修正デュレーション'],
      ['2024/1/1', '2029/1/1', 0.05, 0.06, 2, 0, '=MDURATION(A2,B2,C2,D2,E2,F2)']
    ],
    expectedValues: { 'G2': 4.309 }
  },
  {
    name: 'PRICE',
    category: '07. 財務',
    description: '定期利息証券の価格',
    data: [
      ['決済日', '満期日', '利率', '利回り', '償還価格', '頻度', '基準', '価格'],
      ['2024/1/1', '2029/1/1', 0.05, 0.06, 100, 2, 0, '=PRICE(A2,B2,C2,D2,E2,F2,G2)']
    ],
    expectedValues: { 'H2': 95.82 }
  },
  {
    name: 'YIELD',
    category: '07. 財務',
    description: '定期利息証券の利回り',
    data: [
      ['決済日', '満期日', '利率', '価格', '償還価格', '頻度', '基準', '利回り'],
      ['2024/1/1', '2029/1/1', 0.05, 95, 100, 2, 0, '=YIELD(A2,B2,C2,D2,E2,F2,G2)']
    ],
    expectedValues: { 'H2': 0.0608 }
  },
  {
    name: 'COUPDAYBS',
    category: '07. 財務',
    description: '利払期の開始日から決済日までの日数',
    data: [
      ['決済日', '満期日', '頻度', '基準', '日数'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPDAYBS(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 104 }
  },
  {
    name: 'COUPDAYS',
    category: '07. 財務',
    description: '決済日を含む利払期の日数',
    data: [
      ['決済日', '満期日', '頻度', '基準', '日数'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPDAYS(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 180 }
  },
  {
    name: 'COUPDAYSNC',
    category: '07. 財務',
    description: '決済日から次回利払日までの日数',
    data: [
      ['決済日', '満期日', '頻度', '基準', '日数'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPDAYSNC(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 76 }
  },
  {
    name: 'COUPNCD',
    category: '07. 財務',
    description: '決済日後の次回利払日',
    data: [
      ['決済日', '満期日', '頻度', '基準', '次回利払日'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPNCD(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 45107 }
  },
  {
    name: 'COUPNUM',
    category: '07. 財務',
    description: '決済日から満期日までの利払回数',
    data: [
      ['決済日', '満期日', '頻度', '基準', '利払回数'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPNUM(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 10 }
  },
  {
    name: 'COUPPCD',
    category: '07. 財務',
    description: '決済日直前の利払日',
    data: [
      ['決済日', '満期日', '頻度', '基準', '前回利払日'],
      ['2024/4/15', '2029/1/1', 2, 0, '=COUPPCD(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 45292 }
  },
  {
    name: 'AMORDEGRC',
    category: '07. 財務',
    description: '減価償却費（フランス会計システム）',
    data: [
      ['取得価額', '購入日', '最初の期末', '残存価額', '期', '率', '基準', '減価償却費'],
      [10000, '2024/1/1', '2024/12/31', 1000, 1, 0.15, 1, '=AMORDEGRC(A2,B2,C2,D2,E2,F2,G2)']
    ],
    expectedValues: { 'H2': 1500 }
  },
  {
    name: 'AMORLINC',
    category: '07. 財務',
    description: '減価償却費（フランス会計システム、線形）',
    data: [
      ['取得価額', '購入日', '最初の期末', '残存価額', '期', '率', '基準', '減価償却費'],
      [10000, '2024/1/1', '2024/12/31', 1000, 1, 0.15, 1, '=AMORLINC(A2,B2,C2,D2,E2,F2,G2)']
    ],
    expectedValues: { 'H2': 1350 }
  },
  {
    name: 'CUMIPMT',
    category: '07. 財務',
    description: '累計利息支払額',
    data: [
      ['利率', '期間', '現在価値', '開始期', '終了期', 'タイプ', '累計利息'],
      [0.06, 36, 100000, 1, 12, 0, '=CUMIPMT(A2/12,B2,C2,D2,E2,F2)']
    ],
    expectedValues: { 'G2': -2958.68 }
  },
  {
    name: 'CUMPRINC',
    category: '07. 財務',
    description: '累計元本支払額',
    data: [
      ['利率', '期間', '現在価値', '開始期', '終了期', 'タイプ', '累計元本'],
      [0.06, 36, 100000, 1, 12, 0, '=CUMPRINC(A2/12,B2,C2,D2,E2,F2)']
    ],
    expectedValues: { 'G2': -30405.29 }
  },
  {
    name: 'ODDFPRICE',
    category: '07. 財務',
    description: '期間が半端な最初の期の証券価格',
    data: [
      ['決済日', '満期日', '発行日', '初回利払日', '利率', '利回り', '償還価格', '頻度', '基準', '価格'],
      ['2024/2/15', '2029/1/1', '2024/1/1', '2024/7/1', 0.05, 0.06, 100, 2, 0, '=ODDFPRICE(A2,B2,C2,D2,E2,F2,G2,H2,I2)']
    ],
    expectedValues: { 'J2': 95.05 }
  },
  {
    name: 'ODDFYIELD',
    category: '07. 財務',
    description: '期間が半端な最初の期の証券利回り',
    data: [
      ['決済日', '満期日', '発行日', '初回利払日', '利率', '価格', '償還価格', '頻度', '基準', '利回り'],
      ['2024/2/15', '2029/1/1', '2024/1/1', '2024/7/1', 0.05, 95, 100, 2, 0, '=ODDFYIELD(A2,B2,C2,D2,E2,F2,G2,H2,I2)']
    ],
    expectedValues: { 'J2': 0.0622 }
  },
  {
    name: 'ODDLPRICE',
    category: '07. 財務',
    description: '期間が半端な最後の期の証券価格',
    data: [
      ['決済日', '満期日', '最終利払日', '利率', '利回り', '償還価格', '頻度', '基準', '価格'],
      ['2024/1/1', '2024/10/15', '2024/7/1', 0.05, 0.06, 100, 2, 0, '=ODDLPRICE(A2,B2,C2,D2,E2,F2,G2,H2)']
    ],
    expectedValues: { 'I2': 99.60 }
  },
  {
    name: 'ODDLYIELD',
    category: '07. 財務',
    description: '期間が半端な最後の期の証券利回り',
    data: [
      ['決済日', '満期日', '最終利払日', '利率', '価格', '償還価格', '頻度', '基準', '利回り'],
      ['2024/1/1', '2024/10/15', '2024/7/1', 0.05, 95, 100, 2, 0, '=ODDLYIELD(A2,B2,C2,D2,E2,F2,G2,H2)']
    ],
    expectedValues: { 'I2': 0.0585 }
  },
  {
    name: 'TBILLEQ',
    category: '07. 財務',
    description: '財務省短期証券の債券換算利回り',
    data: [
      ['決済日', '満期日', '割引率', '債券換算利回り'],
      ['2024/1/1', '2024/6/30', 0.045, '=TBILLEQ(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 0.0463 }
  },
  {
    name: 'TBILLPRICE',
    category: '07. 財務',
    description: '財務省短期証券の価格',
    data: [
      ['決済日', '満期日', '割引率', '価格'],
      ['2024/1/1', '2024/6/30', 0.045, '=TBILLPRICE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 97.77 }
  },
  {
    name: 'TBILLYIELD',
    category: '07. 財務',
    description: '財務省短期証券の利回り',
    data: [
      ['決済日', '満期日', '価格', '利回り'],
      ['2024/1/1', '2024/6/30', 97.75, '=TBILLYIELD(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 0.0459 }
  },
  {
    name: 'DOLLARDE',
    category: '07. 財務',
    description: '分数表記のドル価格を小数表記に変換',
    data: [
      ['分数価格', '分母', '小数価格'],
      [1.02, 16, '=DOLLARDE(A2,B2)'],
      [1.1, 32, '=DOLLARDE(A3,B3)']
    ],
    expectedValues: { 'C2': 1.125, 'C3': 1.3125 }
  },
  {
    name: 'DOLLARFR',
    category: '07. 財務',
    description: '小数表記のドル価格を分数表記に変換',
    data: [
      ['小数価格', '分母', '分数価格'],
      [1.125, 16, '=DOLLARFR(A2,B2)'],
      [1.3125, 32, '=DOLLARFR(A3,B3)']
    ],
    expectedValues: { 'C2': 1.02, 'C3': 1.10 }
  },
  {
    name: 'EFFECT',
    category: '07. 財務',
    description: '実効年利率',
    data: [
      ['名目利率', '年間複利回数', '実効利率'],
      [0.06, 4, '=EFFECT(A2,B2)'],
      [0.06, 12, '=EFFECT(A3,B3)']
    ],
    expectedValues: { 'C2': 0.0614, 'C3': 0.0618 }
  },
  {
    name: 'NOMINAL',
    category: '07. 財務',
    description: '名目年利率',
    data: [
      ['実効利率', '年間複利回数', '名目利率'],
      [0.0614, 4, '=NOMINAL(A2,B2)'],
      [0.0618, 12, '=NOMINAL(A3,B3)']
    ],
    expectedValues: { 'C2': 0.0591, 'C3': 0.0584 }
  },
  {
    name: 'PRICEDISC',
    category: '07. 財務',
    description: '割引証券の価格',
    data: [
      ['決済日', '満期日', '割引率', '償還価格', '基準', '価格'],
      ['2024/1/1', '2024/12/31', 0.05, 100, 0, '=PRICEDISC(A2,B2,C2,D2,E2)']
    ],
    expectedValues: { 'F2': 95.00 }
  },
  {
    name: 'RECEIVED',
    category: '07. 財務',
    description: '満期保有証券の受取額',
    data: [
      ['決済日', '満期日', '投資額', '割引率', '基準', '受取額'],
      ['2024/1/1', '2024/12/31', 95, 0.05, 0, '=RECEIVED(A2,B2,C2,D2,E2)']
    ],
    expectedValues: { 'F2': 100.00 }
  },
  {
    name: 'INTRATE',
    category: '07. 財務',
    description: '満期保有証券の利率',
    data: [
      ['決済日', '満期日', '投資額', '償還価格', '基準', '利率'],
      ['2024/1/1', '2024/12/31', 95, 100, 0, '=INTRATE(A2,B2,C2,D2,E2)']
    ],
    expectedValues: { 'F2': 0.0526 }
  },
  {
    name: 'PRICEMAT',
    category: '07. 財務',
    description: '満期利息付証券の価格',
    data: [
      ['決済日', '満期日', '発行日', '利率', '利回り', '基準', '価格'],
      ['2024/1/1', '2024/12/31', '2023/1/1', 0.05, 0.06, 0, '=PRICEMAT(A2,B2,C2,D2,E2,F2)']
    ],
    expectedValues: { 'G2': 98.85 }
  },
  {
    name: 'YIELDDISC',
    category: '07. 財務',
    description: '割引証券の年利回り',
    data: [
      ['決済日', '満期日', '価格', '償還価格', '基準', '利回り'],
      ['2024/1/1', '2024/12/31', 95, 100, 0, '=YIELDDISC(A2,B2,C2,D2,E2)']
    ],
    expectedValues: { 'F2': 0.0527 }
  },
  {
    name: 'YIELDMAT',
    category: '07. 財務',
    description: '満期利息付証券の年利回り',
    data: [
      ['決済日', '満期日', '発行日', '利率', '価格', '基準', '利回り'],
      ['2024/1/1', '2024/12/31', '2023/1/1', 0.05, 95, 0, '=YIELDMAT(A2,B2,C2,D2,E2,F2)']
    ],
    expectedValues: { 'G2': 0.0619 }
  }
];
