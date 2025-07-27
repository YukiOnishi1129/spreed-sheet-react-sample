import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const statisticsTests: IndividualFunctionTest[] = [
  {
    name: 'AVERAGE',
    category: '02. 統計',
    description: '平均値を計算',
    data: [
      ['値1', '値2', '値3', '値4', '平均'],
      [10, 20, 30, 40, '=AVERAGE(A2:D2)']
    ],
    expectedValues: { 'E2': 25 }
  },
  {
    name: 'AVERAGEIF',
    category: '02. 統計',
    description: '条件付き平均値',
    data: [
      ['値', '条件', '平均'],
      [10, 'A', ''],
      [20, 'B', ''],
      [30, 'A', ''],
      [40, 'A', '=AVERAGEIF(B2:B5,"A",A2:A5)']
    ],
    expectedValues: { 'C5': 26.666667 }
  },
  {
    name: 'AVERAGEIFS',
    category: '02. 統計',
    description: '複数条件付き平均値',
    data: [
      ['値', '条件1', '条件2', '平均'],
      [10, 'A', 1, ''],
      [20, 'B', 2, ''],
      [30, 'A', 1, ''],
      [40, 'A', 2, '=AVERAGEIFS(A2:A5,B2:B5,"A",C2:C5,1)']
    ],
    expectedValues: { 'D5': 20 }
  },
  {
    name: 'COUNT',
    category: '02. 統計',
    description: '数値の個数を数える',
    data: [
      ['値1', '値2', '値3', '値4', '個数'],
      [10, 20, 'テキスト', 30, '=COUNT(A2:D2)']
    ],
    expectedValues: { 'E2': 3 }
  },
  {
    name: 'COUNTA',
    category: '02. 統計',
    description: '空白以外のセル数',
    data: [
      ['値1', '値2', '値3', '値4', '個数'],
      [10, 20, 'テキスト', '', '=COUNTA(A2:D2)']
    ],
    expectedValues: { 'E2': 3 }
  },
  {
    name: 'COUNTBLANK',
    category: '02. 統計',
    description: '空白セルの個数',
    data: [
      ['値1', '値2', '値3', '値4', '空白数'],
      [10, '', 'テキスト', '', '=COUNTBLANK(A2:D2)']
    ],
    expectedValues: { 'E2': 2 }
  },
  {
    name: 'COUNTIF',
    category: '02. 統計',
    description: '条件に一致するセル数',
    data: [
      ['値', '個数'],
      [10, ''],
      [20, ''],
      [10, ''],
      [30, '=COUNTIF(A2:A5,10)']
    ],
    expectedValues: { 'B5': 2 }
  },
  {
    name: 'COUNTIFS',
    category: '02. 統計',
    description: '複数条件に一致するセル数',
    data: [
      ['値1', '値2', '個数'],
      [10, 'A', ''],
      [20, 'B', ''],
      [10, 'A', ''],
      [30, 'A', '=COUNTIFS(A2:A5,10,B2:B5,"A")']
    ],
    expectedValues: { 'C5': 2 }
  },
  {
    name: 'MAX',
    category: '02. 統計',
    description: '最大値を返す',
    data: [
      ['値1', '値2', '値3', '値4', '最大値'],
      [10, 30, 20, 40, '=MAX(A2:D2)']
    ],
    expectedValues: { 'E2': 40 }
  },
  {
    name: 'MIN',
    category: '02. 統計',
    description: '最小値を返す',
    data: [
      ['値1', '値2', '値3', '値4', '最小値'],
      [10, 30, 20, 40, '=MIN(A2:D2)']
    ],
    expectedValues: { 'E2': 10 }
  },
  {
    name: 'MAXIFS',
    category: '02. 統計',
    description: '条件付き最大値',
    data: [
      ['値', '条件', '最大値'],
      [10, 'A', ''],
      [20, 'B', ''],
      [30, 'A', ''],
      [40, 'A', '=MAXIFS(A2:A5,B2:B5,"A")']
    ],
    expectedValues: { 'C5': 40 }
  },
  {
    name: 'MINIFS',
    category: '02. 統計',
    description: '条件付き最小値',
    data: [
      ['値', '条件', '最小値'],
      [10, 'A', ''],
      [20, 'B', ''],
      [30, 'A', ''],
      [40, 'A', '=MINIFS(A2:A5,B2:B5,"A")']
    ],
    expectedValues: { 'C5': 10 }
  },
  {
    name: 'MEDIAN',
    category: '02. 統計',
    description: '中央値を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '中央値'],
      [1, 3, 3, 6, 7, '=MEDIAN(A2:E2)']
    ],
    expectedValues: { 'F2': 3 }
  },
  {
    name: 'MODE',
    category: '02. 統計',
    description: '最頻値を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '最頻値'],
      [1, 2, 3, 3, 4, '=MODE(A2:E2)']
    ],
    expectedValues: { 'F2': 3 }
  },
  {
    name: 'STDEV',
    category: '02. 統計',
    description: '標準偏差（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '標準偏差'],
      [10, 20, 30, 40, '=STDEV(A2:D2)']
    ],
    expectedValues: { 'E2': 12.90994449 }
  },
  {
    name: 'STDEV.S',
    category: '02. 統計',
    description: '標準偏差（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '標準偏差'],
      [10, 20, 30, 40, '=STDEV.S(A2:D2)']
    ],
    expectedValues: { 'E2': 12.90994449 }
  },
  {
    name: 'STDEV.P',
    category: '02. 統計',
    description: '標準偏差（母集団）',
    data: [
      ['値1', '値2', '値3', '値4', '標準偏差'],
      [10, 20, 30, 40, '=STDEV.P(A2:D2)']
    ],
    expectedValues: { 'E2': 11.18033989 }
  },
  {
    name: 'VAR',
    category: '02. 統計',
    description: '分散（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '分散'],
      [10, 20, 30, 40, '=VAR(A2:D2)']
    ],
    expectedValues: { 'E2': 166.6666667 }
  },
  {
    name: 'VAR.S',
    category: '02. 統計',
    description: '分散（標本）',
    data: [
      ['値1', '値2', '値3', '値4', '分散'],
      [10, 20, 30, 40, '=VAR.S(A2:D2)']
    ],
    expectedValues: { 'E2': 166.6666667 }
  },
  {
    name: 'VAR.P',
    category: '02. 統計',
    description: '分散（母集団）',
    data: [
      ['値1', '値2', '値3', '値4', '分散'],
      [10, 20, 30, 40, '=VAR.P(A2:D2)']
    ],
    expectedValues: { 'E2': 125 }
  },
  {
    name: 'CORREL',
    category: '02. 統計',
    description: '相関係数を計算',
    data: [
      ['X', 'Y', '', '相関係数'],
      [1, 2, '', ''],
      [2, 4, '', ''],
      [3, 6, '', ''],
      [4, 8, '', '=CORREL(A2:A5,B2:B5)']
    ],
    expectedValues: { 'D5': 1 }
  },
  {
    name: 'LARGE',
    category: '02. 統計',
    description: 'k番目に大きい値',
    data: [
      ['値', '', '', '', 'k', '結果'],
      [10, 30, 20, 40, 2, '=LARGE(A2:D2,E2)']
    ],
    expectedValues: { 'F2': 30 }
  },
  {
    name: 'SMALL',
    category: '02. 統計',
    description: 'k番目に小さい値',
    data: [
      ['値', '', '', '', 'k', '結果'],
      [10, 30, 20, 40, 2, '=SMALL(A2:D2,E2)']
    ],
    expectedValues: { 'F2': 20 }
  },
  {
    name: 'RANK',
    category: '02. 統計',
    description: '順位を返す',
    data: [
      ['値', '順位'],
      [10, '=RANK(A2,$A$2:$A$5)'],
      [30, '=RANK(A3,$A$2:$A$5)'],
      [20, '=RANK(A4,$A$2:$A$5)'],
      [40, '=RANK(A5,$A$2:$A$5)']
    ],
    expectedValues: { 'B2': 4, 'B3': 2, 'B4': 3, 'B5': 1 }
  },
  {
    name: 'PERCENTILE',
    category: '02. 統計',
    description: 'パーセンタイル値',
    data: [
      ['値', '', '', '', 'k', '結果'],
      [1, 2, 3, 4, 0.5, '=PERCENTILE(A2:D2,E2)']
    ],
    expectedValues: { 'F2': 2.5 }
  },
  {
    name: 'QUARTILE',
    category: '02. 統計',
    description: '四分位数',
    data: [
      ['値', '', '', '', '四分位', '結果'],
      [1, 2, 3, 4, 1, '=QUARTILE(A2:D2,E2)']
    ],
    expectedValues: { 'F2': 1.75 }
  },
  {
    name: 'SKEW',
    category: '02. 統計',
    description: '歪度を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '歪度'],
      [3, 4, 5, 2, 1, '=SKEW(A2:E2)']
    ]
  },
  {
    name: 'SKEW.P',
    category: '02. 統計',
    description: '歪度（母集団）',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '歪度'],
      [3, 4, 5, 2, 1, '=SKEW.P(A2:E2)']
    ]
  },
  {
    name: 'MODE.SNGL',
    category: '02. 統計',
    description: '最頻値（単一）',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '最頻値'],
      [1, 2, 2, 3, 3, '=MODE.SNGL(A2:E2)']
    ]
  },
  {
    name: 'QUARTILE.INC',
    category: '02. 統計',
    description: '四分位数（境界値含む）',
    data: [
      ['値1', '値2', '値3', '値4', '第1四分位'],
      [1, 2, 3, 4, '=QUARTILE.INC(A2:D2,1)']
    ]
  },
  {
    name: 'PERCENTILE.INC',
    category: '02. 統計',
    description: 'パーセンタイル（境界値含む）',
    data: [
      ['値1', '値2', '値3', '値4', '50%タイル'],
      [1, 2, 3, 4, '=PERCENTILE.INC(A2:D2,0.5)']
    ]
  },
  {
    name: 'RANK.EQ',
    category: '02. 統計',
    description: '順位（同順位は最小）',
    data: [
      ['値', '順位'],
      [85, '=RANK.EQ(A2,$A$2:$A$5)'],
      [92, '=RANK.EQ(A3,$A$2:$A$5)'],
      [78, '=RANK.EQ(A4,$A$2:$A$5)'],
      [92, '=RANK.EQ(A5,$A$2:$A$5)']
    ]
  },
  {
    name: 'KURT',
    category: '02. 統計',
    description: '尖度を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '尖度'],
      [3, 4, 5, 2, 1, '=KURT(A2:E2)']
    ]
  },
  {
    name: 'STANDARDIZE',
    category: '02. 統計',
    description: '標準化する',
    data: [
      ['値', '平均', '標準偏差', '標準化値'],
      [42, 40, 1.5, '=STANDARDIZE(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 1.333333 }
  },
  {
    name: 'DEVSQ',
    category: '02. 統計',
    description: '偏差平方和を計算',
    data: [
      ['値1', '値2', '値3', '値4', '偏差平方和'],
      [4, 5, 8, 7, '=DEVSQ(A2:D2)']
    ],
    expectedValues: { 'E2': 10 }
  },
  {
    name: 'GEOMEAN',
    category: '02. 統計',
    description: '幾何平均を計算',
    data: [
      ['値1', '値2', '値3', '値4', '幾何平均'],
      [4, 5, 8, 7, '=GEOMEAN(A2:D2)']
    ]
  },
  {
    name: 'HARMEAN',
    category: '02. 統計',
    description: '調和平均を計算',
    data: [
      ['値1', '値2', '値3', '値4', '調和平均'],
      [4, 5, 8, 7, '=HARMEAN(A2:D2)']
    ]
  },
  {
    name: 'TRIMMEAN',
    category: '02. 統計',
    description: 'トリム平均を計算',
    data: [
      ['値1', '値2', '値3', '値4', '値5', '値6', '除外率', 'トリム平均'],
      [1, 2, 3, 4, 5, 100, 0.2, '=TRIMMEAN(A2:F2,G2)']
    ]
  },
  {
    name: 'NORM.DIST',
    category: '02. 統計',
    description: '正規分布',
    data: [
      ['値', '平均', '標準偏差', '累積', '確率密度'],
      [42, 40, 1.5, 'FALSE', '=NORM.DIST(A2,B2,C2,D2)'],
      [42, 40, 1.5, 'TRUE', '=NORM.DIST(A3,B3,C3,D3)']
    ]
  },
  {
    name: 'NORM.INV',
    category: '02. 統計',
    description: '正規分布の逆関数',
    data: [
      ['確率', '平均', '標準偏差', '値'],
      [0.908789, 40, 1.5, '=NORM.INV(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 42 }
  },
  {
    name: 'NORM.S.DIST',
    category: '02. 統計',
    description: '標準正規分布',
    data: [
      ['z値', '累積', '確率'],
      [1.333333, 'TRUE', '=NORM.S.DIST(A2,B2)'],
      [1.333333, 'FALSE', '=NORM.S.DIST(A3,B3)']
    ]
  },
  {
    name: 'NORM.S.INV',
    category: '02. 統計',
    description: '標準正規分布の逆関数',
    data: [
      ['確率', 'z値'],
      [0.908789, '=NORM.S.INV(A2)']
    ],
    expectedValues: { 'B2': 1.333333 }
  },
  {
    name: 'T.DIST',
    category: '02. 統計',
    description: 't分布（左側）',
    data: [
      ['値', '自由度', '累積', '確率'],
      [1.96, 60, 'TRUE', '=T.DIST(A2,B2,C2)']
    ]
  },
  {
    name: 'T.DIST.2T',
    category: '02. 統計',
    description: 't分布（両側）',
    data: [
      ['値', '自由度', '確率'],
      [1.96, 60, '=T.DIST.2T(A2,B2)']
    ]
  },
  {
    name: 'T.INV',
    category: '02. 統計',
    description: 't分布の逆関数（左側）',
    data: [
      ['確率', '自由度', '値'],
      [0.975, 60, '=T.INV(A2,B2)']
    ],
    expectedValues: { 'C2': 2.000298 }
  },
  {
    name: 'T.INV.2T',
    category: '02. 統計',
    description: 't分布の逆関数（両側）',
    data: [
      ['確率', '自由度', '値'],
      [0.05, 60, '=T.INV.2T(A2,B2)']
    ],
    expectedValues: { 'C2': 2.000298 }
  },
  {
    name: 'CHISQ.DIST',
    category: '02. 統計',
    description: 'カイ二乗分布',
    data: [
      ['値', '自由度', '累積', '確率'],
      [18.307, 10, 'TRUE', '=CHISQ.DIST(A2,B2,C2)']
    ]
  },
  {
    name: 'CHISQ.INV',
    category: '02. 統計',
    description: 'カイ二乗分布の逆関数',
    data: [
      ['確率', '自由度', '値'],
      [0.95, 10, '=CHISQ.INV(A2,B2)']
    ],
    expectedValues: { 'C2': 18.307 }
  },
  {
    name: 'F.DIST',
    category: '02. 統計',
    description: 'F分布',
    data: [
      ['値', '自由度1', '自由度2', '累積', '確率'],
      [2.5, 5, 10, 'TRUE', '=F.DIST(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': 0.8823 }
  },
  {
    name: 'F.INV',
    category: '02. 統計',
    description: 'F分布の逆関数',
    data: [
      ['確率', '自由度1', '自由度2', '値'],
      [0.95, 5, 10, '=F.INV(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 3.3258 }
  },
  {
    name: 'EXPON.DIST',
    category: '02. 統計',
    description: '指数分布',
    data: [
      ['値', 'λ', '累積', '確率'],
      [0.2, 10, 'TRUE', '=EXPON.DIST(A2,B2,C2)']
    ],
    expectedValues: { 'D2': 0.8647 }
  },
  {
    name: 'COVARIANCE.P',
    category: '02. 統計',
    description: '母共分散',
    data: [
      ['データ1', 'データ2', '', '共分散'],
      [1, 10, '', '=COVARIANCE.P(A2:A5,B2:B5)'],
      [2, 15, '', ''],
      [3, 20, '', ''],
      [4, 25, '', '']
    ],
    expectedValues: { 'D2': 6.25 }
  },
  {
    name: 'COVARIANCE.S',
    category: '02. 統計',
    description: '標本共分散',
    data: [
      ['データ1', 'データ2', '', '共分散'],
      [1, 10, '', '=COVARIANCE.S(A2:A5,B2:B5)'],
      [2, 15, '', ''],
      [3, 20, '', ''],
      [4, 25, '', '']
    ]
  },
  {
    name: 'MODE.MULT',
    category: '02. 統計',
    description: '最頻値（複数）',
    data: [
      ['データ', '最頻値'],
      [1, '=MODE.MULT(A2:A8)'],
      [2, ''],
      [2, ''],
      [3, ''],
      [3, ''],
      [3, ''],
      [4, '']
    ]
  },
  {
    name: 'RANK.AVG',
    category: '02. 統計',
    description: '順位（平均）',
    data: [
      ['データ', '順位'],
      [89, '=RANK.AVG(A2,$A$2:$A$6)'],
      [92, '=RANK.AVG(A3,$A$2:$A$6)'],
      [85, '=RANK.AVG(A4,$A$2:$A$6)'],
      [92, '=RANK.AVG(A5,$A$2:$A$6)'],
      [88, '=RANK.AVG(A6,$A$2:$A$6)']
    ],
    expectedValues: { 'B2': 3, 'B3': 1.5, 'B4': 5, 'B5': 1.5, 'B6': 4 }
  },
  {
    name: 'LOGNORM.DIST',
    category: '02. 統計',
    description: '対数正規分布',
    data: [
      ['値', '平均', '標準偏差', '累積', '確率'],
      [4, 3.5, 1.2, 'TRUE', '=LOGNORM.DIST(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'LOGNORM.INV',
    category: '02. 統計',
    description: '対数正規分布の逆関数',
    data: [
      ['確率', '平均', '標準偏差', '値'],
      [0.5, 3.5, 1.2, '=LOGNORM.INV(A2,B2,C2)']
    ]
  },
  {
    name: 'T.DIST.RT',
    category: '02. 統計',
    description: 'T分布（右側）',
    data: [
      ['値', '自由度', '確率'],
      [2, 10, '=T.DIST.RT(A2,B2)']
    ]
  },
  {
    name: 'CHISQ.DIST.RT',
    category: '02. 統計',
    description: 'カイ二乗分布（右側）',
    data: [
      ['値', '自由度', '確率'],
      [10, 5, '=CHISQ.DIST.RT(A2,B2)']
    ]
  },
  {
    name: 'CHISQ.INV.RT',
    category: '02. 統計',
    description: 'カイ二乗分布の逆関数（右側）',
    data: [
      ['確率', '自由度', '値'],
      [0.05, 5, '=CHISQ.INV.RT(A2,B2)']
    ]
  },
  {
    name: 'F.DIST.RT',
    category: '02. 統計',
    description: 'F分布（右側）',
    data: [
      ['値', '自由度1', '自由度2', '確率'],
      [2, 5, 10, '=F.DIST.RT(A2,B2,C2)']
    ]
  },
  {
    name: 'F.INV.RT',
    category: '02. 統計',
    description: 'F分布の逆関数（右側）',
    data: [
      ['確率', '自由度1', '自由度2', '値'],
      [0.05, 5, 10, '=F.INV.RT(A2,B2,C2)']
    ]
  },
  {
    name: 'BETA.DIST',
    category: '02. 統計',
    description: 'ベータ分布',
    data: [
      ['値', 'α', 'β', '下限', '上限', '累積', '確率'],
      [0.5, 2, 5, 0, 1, 'TRUE', '=BETA.DIST(A2,B2,C2,D2,E2,F2)']
    ]
  },
  {
    name: 'BETA.INV',
    category: '02. 統計',
    description: 'ベータ分布の逆関数',
    data: [
      ['確率', 'α', 'β', '下限', '上限', '値'],
      [0.5, 2, 5, 0, 1, '=BETA.INV(A2,B2,C2,D2,E2)']
    ]
  },
  {
    name: 'GAMMA.DIST',
    category: '02. 統計',
    description: 'ガンマ分布',
    data: [
      ['値', 'α', 'β', '累積', '確率'],
      [10, 9, 2, 'TRUE', '=GAMMA.DIST(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'GAMMA.INV',
    category: '02. 統計',
    description: 'ガンマ分布の逆関数',
    data: [
      ['確率', 'α', 'β', '値'],
      [0.5, 9, 2, '=GAMMA.INV(A2,B2,C2)']
    ]
  },
  {
    name: 'WEIBULL.DIST',
    category: '02. 統計',
    description: 'ワイブル分布',
    data: [
      ['値', 'α', 'β', '累積', '確率'],
      [100, 20, 100, 'TRUE', '=WEIBULL.DIST(A2,B2,C2,D2)']
    ]
  },
  {
    name: 'PERCENTILE.EXC',
    category: '02. 統計',
    description: 'パーセンタイル（除外）',
    data: [
      ['データ', '', 'パーセンタイル'],
      [1, '', ''],
      [2, '', ''],
      [3, '', ''],
      [4, '', ''],
      [5, 0.25, '=PERCENTILE.EXC(A2:A6,B6)']
    ],
    expectedValues: { 'C6': 1.75 }
  },
  {
    name: 'PERCENTRANK.INC',
    category: '02. 統計',
    description: 'パーセントランク（含む）',
    data: [
      ['データ', '値', 'ランク'],
      [1, 3, '=PERCENTRANK.INC(A2:A6,B2)'],
      [2, '', ''],
      [3, '', ''],
      [4, '', ''],
      [5, '', '']
    ],
    expectedValues: { 'C2': 0.5 }
  },
  {
    name: 'PERCENTRANK.EXC',
    category: '02. 統計',
    description: 'パーセントランク（除外）',
    data: [
      ['データ', '値', 'ランク'],
      [1, 3, '=PERCENTRANK.EXC(A2:A6,B2)'],
      [2, '', ''],
      [3, '', ''],
      [4, '', ''],
      [5, '', '']
    ]
  },
  {
    name: 'QUARTILE.EXC',
    category: '02. 統計',
    description: '四分位数（除外）',
    data: [
      ['データ', '四分位', '値'],
      [1, 1, '=QUARTILE.EXC(A2:A9,B2)'],
      [2, 2, '=QUARTILE.EXC(A2:A9,B3)'],
      [3, 3, '=QUARTILE.EXC(A2:A9,B4)'],
      [4, '', ''],
      [5, '', ''],
      [6, '', ''],
      [7, '', ''],
      [8, '', '']
    ]
  },
  {
    name: 'STDEVA',
    category: '02. 統計',
    description: '標準偏差（テキスト含む）',
    data: [
      ['データ', '標準偏差'],
      [10, '=STDEVA(A2:A6)'],
      [20, ''],
      ['テキスト', ''],
      [30, ''],
      [40, '']
    ]
  },
  {
    name: 'STDEVPA',
    category: '02. 統計',
    description: '母標準偏差（テキスト含む）',
    data: [
      ['データ', '母標準偏差'],
      [10, '=STDEVPA(A2:A6)'],
      [20, ''],
      ['テキスト', ''],
      [30, ''],
      [40, '']
    ]
  },
  {
    name: 'VARA',
    category: '02. 統計',
    description: '分散（テキスト含む）',
    data: [
      ['データ', '分散'],
      [10, '=VARA(A2:A6)'],
      [20, ''],
      ['テキスト', ''],
      [30, ''],
      [40, '']
    ]
  },
  {
    name: 'VARPA',
    category: '02. 統計',
    description: '母分散（テキスト含む）',
    data: [
      ['データ', '母分散'],
      [10, '=VARPA(A2:A6)'],
      [20, ''],
      ['テキスト', ''],
      [30, ''],
      [40, '']
    ]
  },
  {
    name: 'COUNT',
    category: '02. 統計',
    description: '数値の個数',
    data: [
      ['データ', '個数'],
      [10, '=COUNT(A2:A7)'],
      [20, ''],
      ['テキスト', ''],
      [30, ''],
      ['', ''],
      [40, '']
    ],
    expectedValues: { 'B2': 4 }
  },
  {
    name: 'AVEDEV',
    category: '02. 統計',
    description: '平均偏差',
    data: [
      ['データ', '平均偏差'],
      [10, '=AVEDEV(A2:A5)'],
      [20, ''],
      [30, ''],
      [40, '']
    ],
    expectedValues: { 'B2': 10 }
  },
  {
    name: 'BETA.DIST',
    category: '02. 統計',
    description: 'ベータ分布',
    data: [
      ['値', 'α', 'β', '累積', '結果'],
      [0.5, 2, 3, 'TRUE', '=BETA.DIST(A2,B2,C2,D2)'],
      [0.5, 2, 3, 'FALSE', '=BETA.DIST(A3,B3,C3,D3)']
    ],
    expectedValues: { 'E2': 0.3125, 'E3': 1.5 }
  },
  {
    name: 'BETA.INV',
    category: '02. 統計',
    description: 'ベータ分布の逆関数',
    data: [
      ['確率', 'α', 'β', '結果'],
      [0.5, 2, 3, '=BETA.INV(A2,B2,C2)'],
      [0.8, 2, 3, '=BETA.INV(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 0.448, 'D3': 0.618 }
  },
  {
    name: 'BINOM.DIST',
    category: '02. 統計',
    description: '二項分布',
    data: [
      ['成功数', '試行回数', '成功率', '累積', '結果'],
      [3, 10, 0.5, 'TRUE', '=BINOM.DIST(A2,B2,C2,D2)'],
      [3, 10, 0.5, 'FALSE', '=BINOM.DIST(A3,B3,C3,D3)']
    ],
    expectedValues: { 'E2': 0.171875, 'E3': 0.117188 }
  },
  {
    name: 'BINOM.INV',
    category: '02. 統計',
    description: '二項分布の逆関数',
    data: [
      ['試行回数', '成功率', '確率', '結果'],
      [10, 0.5, 0.5, '=BINOM.INV(A2,B2,C2)'],
      [10, 0.5, 0.8, '=BINOM.INV(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 5, 'D3': 6 }
  },
  {
    name: 'CHISQ.DIST',
    category: '02. 統計',
    description: 'カイ二乗分布',
    data: [
      ['値', '自由度', '累積', '結果'],
      [5, 3, 'TRUE', '=CHISQ.DIST(A2,B2,C2)'],
      [5, 3, 'FALSE', '=CHISQ.DIST(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 0.828, 'D3': 0.064 }
  },
  {
    name: 'CHISQ.DIST.RT',
    category: '02. 統計',
    description: 'カイ二乗分布（右側）',
    data: [
      ['値', '自由度', '結果'],
      [5, 3, '=CHISQ.DIST.RT(A2,B2)'],
      [10, 5, '=CHISQ.DIST.RT(A3,B3)']
    ],
    expectedValues: { 'C2': 0.172, 'C3': 0.075 }
  },
  {
    name: 'CHISQ.TEST',
    category: '02. 統計',
    description: 'カイ二乗検定',
    data: [
      ['観測値', '期待値', '', '', '検定結果'],
      [10, 12, '', '', '=CHISQ.TEST(A2:B4,D2:E4)'],
      [15, 14, '', '', ''],
      [20, 19, '', '', ''],
      [12, 10, '', '', ''],
      [14, 15, '', '', ''],
      [19, 20, '', '', '']
    ],
    expectedValues: { 'F2': 0.807 }
  },
  {
    name: 'CONFIDENCE.NORM',
    category: '02. 統計',
    description: '正規分布の信頼区間',
    data: [
      ['有意水準', '標準偏差', 'サンプル数', '信頼区間'],
      [0.05, 2.5, 50, '=CONFIDENCE.NORM(A2,B2,C2)'],
      [0.01, 2.5, 50, '=CONFIDENCE.NORM(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 0.693, 'D3': 0.911 }
  },
  {
    name: 'CONFIDENCE.T',
    category: '02. 統計',
    description: 'T分布の信頼区間',
    data: [
      ['有意水準', '標準偏差', 'サンプル数', '信頼区間'],
      [0.05, 2.5, 10, '=CONFIDENCE.T(A2,B2,C2)'],
      [0.01, 2.5, 10, '=CONFIDENCE.T(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 1.789, 'D3': 2.574 }
  },
  {
    name: 'COVARIANCE.P',
    category: '02. 統計',
    description: '母共分散',
    data: [
      ['X', 'Y', '', '母共分散'],
      [1, 2, '', '=COVARIANCE.P(A2:A5,B2:B5)'],
      [2, 4, '', ''],
      [3, 6, '', ''],
      [4, 8, '', '']
    ],
    expectedValues: { 'D2': 2.5 }
  },
  {
    name: 'COVARIANCE.S',
    category: '02. 統計',
    description: '標本共分散',
    data: [
      ['X', 'Y', '', '標本共分散'],
      [1, 2, '', '=COVARIANCE.S(A2:A5,B2:B5)'],
      [2, 4, '', ''],
      [3, 6, '', ''],
      [4, 8, '', '']
    ],
    expectedValues: { 'D2': 3.333 }
  },
  {
    name: 'DEVSQ',
    category: '02. 統計',
    description: '平方偏差の合計',
    data: [
      ['データ', '平方偏差の合計'],
      [10, '=DEVSQ(A2:A5)'],
      [20, ''],
      [30, ''],
      [40, '']
    ],
    expectedValues: { 'B2': 500 }
  },
  {
    name: 'EXPON.DIST',
    category: '02. 統計',
    description: '指数分布',
    data: [
      ['値', 'λ', '累積', '結果'],
      [0.5, 1, 'TRUE', '=EXPON.DIST(A2,B2,C2)'],
      [0.5, 1, 'FALSE', '=EXPON.DIST(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 0.393469, 'D3': 0.606531 }
  },
  {
    name: 'F.TEST',
    category: '02. 統計',
    description: 'F検定',
    data: [
      ['データ1', 'データ2', '', 'F検定'],
      [10, 12, '', '=F.TEST(A2:A5,B2:B5)'],
      [20, 18, '', ''],
      [30, 32, '', ''],
      [40, 38, '', '']
    ],
    expectedValues: { 'D2': 0.646 }
  },
  {
    name: 'FISHER',
    category: '02. 統計',
    description: 'フィッシャー変換',
    data: [
      ['相関係数', 'フィッシャー変換'],
      [0.5, '=FISHER(A2)'],
      [0.75, '=FISHER(A3)'],
      [0.9, '=FISHER(A4)']
    ],
    expectedValues: { 'B2': 0.549306, 'B3': 0.972955, 'B4': 1.472219 }
  },
  {
    name: 'FISHERINV',
    category: '02. 統計',
    description: 'フィッシャー変換の逆関数',
    data: [
      ['値', '相関係数'],
      [0.549306, '=FISHERINV(A2)'],
      [0.972955, '=FISHERINV(A3)'],
      [1.472219, '=FISHERINV(A4)']
    ],
    expectedValues: { 'B2': 0.5, 'B3': 0.75, 'B4': 0.9 }
  },
  {
    name: 'GAMMA',
    category: '02. 統計',
    description: 'ガンマ関数',
    data: [
      ['値', 'ガンマ関数'],
      [1, '=GAMMA(A2)'],
      [2, '=GAMMA(A3)'],
      [3, '=GAMMA(A4)'],
      [4, '=GAMMA(A5)']
    ],
    expectedValues: { 'B2': 1, 'B3': 1, 'B4': 2, 'B5': 6 }
  },
  {
    name: 'GAMMA.DIST',
    category: '02. 統計',
    description: 'ガンマ分布',
    data: [
      ['値', 'α', 'β', '累積', '結果'],
      [5, 3, 2, 'TRUE', '=GAMMA.DIST(A2,B2,C2,D2)'],
      [5, 3, 2, 'FALSE', '=GAMMA.DIST(A3,B3,C3,D3)']
    ],
    expectedValues: { 'E2': 0.761, 'E3': 0.065 }
  },
  {
    name: 'GAMMA.INV',
    category: '02. 統計',
    description: 'ガンマ分布の逆関数',
    data: [
      ['確率', 'α', 'β', '結果'],
      [0.5, 3, 2, '=GAMMA.INV(A2,B2,C2)'],
      [0.75, 3, 2, '=GAMMA.INV(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 4.671, 'D3': 6.727 }
  },
  {
    name: 'GAMMALN',
    category: '02. 統計',
    description: 'ガンマ関数の自然対数',
    data: [
      ['値', '自然対数'],
      [2, '=GAMMALN(A2)'],
      [3, '=GAMMALN(A3)'],
      [4, '=GAMMALN(A4)'],
      [5, '=GAMMALN(A5)']
    ],
    expectedValues: { 'B2': 0, 'B3': 0.693147, 'B4': 1.791759, 'B5': 3.178054 }
  },
  {
    name: 'GAMMALN.PRECISE',
    category: '02. 統計',
    description: 'ガンマ関数の自然対数（高精度）',
    data: [
      ['値', '自然対数'],
      [2, '=GAMMALN.PRECISE(A2)'],
      [3, '=GAMMALN.PRECISE(A3)'],
      [4, '=GAMMALN.PRECISE(A4)']
    ],
    expectedValues: { 'B2': 0, 'B3': 0.693147, 'B4': 1.791759 }
  },
  {
    name: 'GAUSS',
    category: '02. 統計',
    description: '標準正規分布（0.5を引いた値）',
    data: [
      ['z値', 'ガウス分布'],
      [0, '=GAUSS(A2)'],
      [1, '=GAUSS(A3)'],
      [2, '=GAUSS(A4)']
    ],
    expectedValues: { 'B2': 0, 'B3': 0.341345, 'B4': 0.477250 }
  },
  {
    name: 'HYPGEOM.DIST',
    category: '02. 統計',
    description: '超幾何分布',
    data: [
      ['成功数', '標本数', '母成功数', '母集団数', '累積', '結果'],
      [2, 5, 3, 10, 'FALSE', '=HYPGEOM.DIST(A2,B2,C2,D2,E2)'],
      [2, 5, 3, 10, 'TRUE', '=HYPGEOM.DIST(A3,B3,C3,D3,E3)']
    ],
    expectedValues: { 'F2': 0.238095, 'F3': 0.916667 }
  },
  {
    name: 'KURT',
    category: '02. 統計',
    description: '尖度',
    data: [
      ['データ', '尖度'],
      [10, '=KURT(A2:A6)'],
      [20, ''],
      [30, ''],
      [40, ''],
      [50, '']
    ],
    expectedValues: { 'B2': -1.2 }
  },
  {
    name: 'LARGE',
    category: '02. 統計',
    description: 'k番目に大きな値',
    data: [
      ['データ', 'k', '結果'],
      [10, 1, '=LARGE($A$2:$A$6,B2)'],
      [20, 2, '=LARGE($A$2:$A$6,B3)'],
      [30, 3, '=LARGE($A$2:$A$6,B4)'],
      [40, '', ''],
      [50, '', '']
    ],
    expectedValues: { 'C2': 50, 'C3': 40, 'C4': 30 }
  },
  {
    name: 'MAX',
    category: '02. 統計',
    description: '最大値',
    data: [
      ['データ', '最大値'],
      [10, '=MAX(A2:A5)'],
      [20, ''],
      [30, ''],
      [40, '']
    ],
    expectedValues: { 'B2': 40 }
  },
  {
    name: 'MIN',
    category: '02. 統計',
    description: '最小値',
    data: [
      ['データ', '最小値'],
      [10, '=MIN(A2:A5)'],
      [20, ''],
      [30, ''],
      [40, '']
    ],
    expectedValues: { 'B2': 10 }
  },
  {
    name: 'PEARSON',
    category: '02. 統計',
    description: 'ピアソン相関係数',
    data: [
      ['X', 'Y', '', '相関係数'],
      [1, 2, '', '=PEARSON(A2:A5,B2:B5)'],
      [2, 4, '', ''],
      [3, 6, '', ''],
      [4, 8, '', '']
    ],
    expectedValues: { 'D2': 1 }
  },
  {
    name: 'PERCENTILE.EXC',
    category: '02. 統計',
    description: '百分位数（除外）',
    data: [
      ['データ', '百分位', '結果'],
      [10, 0.25, '=PERCENTILE.EXC($A$2:$A$6,B2)'],
      [20, 0.5, '=PERCENTILE.EXC($A$2:$A$6,B3)'],
      [30, 0.75, '=PERCENTILE.EXC($A$2:$A$6,B4)'],
      [40, '', ''],
      [50, '', '']
    ],
    expectedValues: { 'C2': 15, 'C3': 30, 'C4': 45 }
  },
  {
    name: 'PERCENTILE.INC',
    category: '02. 統計',
    description: '百分位数（含む）',
    data: [
      ['データ', '百分位', '結果'],
      [10, 0.25, '=PERCENTILE.INC($A$2:$A$6,B2)'],
      [20, 0.5, '=PERCENTILE.INC($A$2:$A$6,B3)'],
      [30, 0.75, '=PERCENTILE.INC($A$2:$A$6,B4)'],
      [40, '', ''],
      [50, '', '']
    ],
    expectedValues: { 'C2': 20, 'C3': 30, 'C4': 40 }
  },
  {
    name: 'PERMUT',
    category: '02. 統計',
    description: '順列',
    data: [
      ['n', 'k', '順列'],
      [5, 2, '=PERMUT(A2,B2)'],
      [10, 3, '=PERMUT(A3,B3)'],
      [7, 4, '=PERMUT(A4,B4)']
    ],
    expectedValues: { 'C2': 20, 'C3': 720, 'C4': 840 }
  },
  {
    name: 'PERMUTATIONA',
    category: '02. 統計',
    description: '重複順列',
    data: [
      ['n', 'k', '重複順列'],
      [5, 2, '=PERMUTATIONA(A2,B2)'],
      [10, 3, '=PERMUTATIONA(A3,B3)'],
      [7, 4, '=PERMUTATIONA(A4,B4)']
    ],
    expectedValues: { 'C2': 25, 'C3': 1000, 'C4': 2401 }
  },
  {
    name: 'PHI',
    category: '02. 統計',
    description: '標準正規分布の密度関数',
    data: [
      ['z値', '密度'],
      [0, '=PHI(A2)'],
      [1, '=PHI(A3)'],
      [-1, '=PHI(A4)']
    ],
    expectedValues: { 'B2': 0.398942, 'B3': 0.241971, 'B4': 0.241971 }
  },
  {
    name: 'POISSON.DIST',
    category: '02. 統計',
    description: 'ポアソン分布',
    data: [
      ['x', '平均', '累積', '結果'],
      [2, 3, 'TRUE', '=POISSON.DIST(A2,B2,C2)'],
      [2, 3, 'FALSE', '=POISSON.DIST(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 0.423190, 'D3': 0.224042 }
  },
  {
    name: 'QUARTILE.EXC',
    category: '02. 統計',
    description: '四分位数（除外）',
    data: [
      ['データ', '四分位', '結果'],
      [10, 1, '=QUARTILE.EXC($A$2:$A$6,B2)'],
      [20, 2, '=QUARTILE.EXC($A$2:$A$6,B3)'],
      [30, 3, '=QUARTILE.EXC($A$2:$A$6,B4)'],
      [40, '', ''],
      [50, '', '']
    ],
    expectedValues: { 'C2': 15, 'C3': 30, 'C4': 45 }
  },
  {
    name: 'QUARTILE.INC',
    category: '02. 統計',
    description: '四分位数（含む）',
    data: [
      ['データ', '四分位', '結果'],
      [10, 1, '=QUARTILE.INC($A$2:$A$6,B2)'],
      [20, 2, '=QUARTILE.INC($A$2:$A$6,B3)'],
      [30, 3, '=QUARTILE.INC($A$2:$A$6,B4)'],
      [40, '', ''],
      [50, '', '']
    ],
    expectedValues: { 'C2': 20, 'C3': 30, 'C4': 40 }
  },
  {
    name: 'RANK.AVG',
    category: '02. 統計',
    description: '平均順位',
    data: [
      ['スコア', '順位'],
      [85, '=RANK.AVG(A2,$A$2:$A$6)'],
      [90, '=RANK.AVG(A3,$A$2:$A$6)'],
      [85, '=RANK.AVG(A4,$A$2:$A$6)'],
      [80, '=RANK.AVG(A5,$A$2:$A$6)'],
      [95, '=RANK.AVG(A6,$A$2:$A$6)']
    ],
    expectedValues: { 'B2': 3.5, 'B3': 2, 'B4': 3.5, 'B5': 5, 'B6': 1 }
  },
  {
    name: 'RANK.EQ',
    category: '02. 統計',
    description: '同順位',
    data: [
      ['スコア', '順位'],
      [85, '=RANK.EQ(A2,$A$2:$A$6)'],
      [90, '=RANK.EQ(A3,$A$2:$A$6)'],
      [85, '=RANK.EQ(A4,$A$2:$A$6)'],
      [80, '=RANK.EQ(A5,$A$2:$A$6)'],
      [95, '=RANK.EQ(A6,$A$2:$A$6)']
    ],
    expectedValues: { 'B2': 3, 'B3': 2, 'B4': 3, 'B5': 5, 'B6': 1 }
  },
  {
    name: 'RSQ',
    category: '02. 統計',
    description: '決定係数',
    data: [
      ['X', 'Y', '', '決定係数'],
      [1, 2, '', '=RSQ(A2:A5,B2:B5)'],
      [2, 4, '', ''],
      [3, 6, '', ''],
      [4, 8, '', '']
    ],
    expectedValues: { 'D2': 1 }
  },
  {
    name: 'SKEW.P',
    category: '02. 統計',
    description: '歪度（母集団）',
    data: [
      ['データ', '歪度'],
      [10, '=SKEW.P(A2:A6)'],
      [20, ''],
      [30, ''],
      [40, ''],
      [50, '']
    ],
    expectedValues: { 'B2': 0 }
  },
  {
    name: 'SLOPE',
    category: '02. 統計',
    description: '回帰直線の傾き',
    data: [
      ['Y', 'X', '', '傾き'],
      [2, 1, '', '=SLOPE(A2:A5,B2:B5)'],
      [4, 2, '', ''],
      [6, 3, '', ''],
      [8, 4, '', '']
    ],
    expectedValues: { 'D2': 2 }
  },
  {
    name: 'SMALL',
    category: '02. 統計',
    description: 'k番目に小さな値',
    data: [
      ['データ', 'k', '結果'],
      [50, 1, '=SMALL($A$2:$A$6,B2)'],
      [40, 2, '=SMALL($A$2:$A$6,B3)'],
      [30, 3, '=SMALL($A$2:$A$6,B4)'],
      [20, '', ''],
      [10, '', '']
    ],
    expectedValues: { 'C2': 10, 'C3': 20, 'C4': 30 }
  },
  {
    name: 'STANDARDIZE',
    category: '02. 統計',
    description: '標準化',
    data: [
      ['値', '平均', '標準偏差', '標準化'],
      [85, 75, 10, '=STANDARDIZE(A2,B2,C2)'],
      [65, 75, 10, '=STANDARDIZE(A3,B3,C3)'],
      [75, 75, 10, '=STANDARDIZE(A4,B4,C4)']
    ],
    expectedValues: { 'D2': 1, 'D3': -1, 'D4': 0 }
  },
  {
    name: 'STEYX',
    category: '02. 統計',
    description: '回帰の標準誤差',
    data: [
      ['Y', 'X', '', '標準誤差'],
      [2, 1, '', '=STEYX(A2:A5,B2:B5)'],
      [4.1, 2, '', ''],
      [5.9, 3, '', ''],
      [8, 4, '', '']
    ],
    expectedValues: { 'D2': 0.05 }
  },
  {
    name: 'T.TEST',
    category: '02. 統計',
    description: 'T検定',
    data: [
      ['データ1', 'データ2', '', 'T検定'],
      [10, 12, '', '=T.TEST(A2:A5,B2:B5,2,1)'],
      [20, 18, '', ''],
      [30, 32, '', ''],
      [40, 38, '', '']
    ],
    expectedValues: { 'D2': 0.423 }
  },
  {
    name: 'TRIMMEAN',
    category: '02. 統計',
    description: '調整平均',
    data: [
      ['データ', '除外率', '調整平均'],
      [10, 0.2, '=TRIMMEAN($A$2:$A$9,B2)'],
      [20, '', ''],
      [30, '', ''],
      [40, '', ''],
      [50, '', ''],
      [60, '', ''],
      [70, '', ''],
      [80, '', '']
    ],
    expectedValues: { 'C2': 45 }
  },
  {
    name: 'WEIBULL.DIST',
    category: '02. 統計',
    description: 'ワイブル分布',
    data: [
      ['値', '形状', '尺度', '累積', '結果'],
      [1, 2, 1, 'TRUE', '=WEIBULL.DIST(A2,B2,C2,D2)'],
      [1, 2, 1, 'FALSE', '=WEIBULL.DIST(A3,B3,C3,D3)']
    ],
    expectedValues: { 'E2': 0.632121, 'E3': 0.735759 }
  },
  {
    name: 'Z.TEST',
    category: '02. 統計',
    description: 'Z検定',
    data: [
      ['データ', '', '', 'Z検定'],
      [10, '', '', '=Z.TEST(A2:A6,25)'],
      [20, '', '', ''],
      [30, '', '', ''],
      [40, '', '', ''],
      [50, '', '', '']
    ],
    expectedValues: { 'D2': 0.066 }
  }
];
