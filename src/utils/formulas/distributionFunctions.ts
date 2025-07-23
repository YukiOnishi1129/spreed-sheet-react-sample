// 分布関数・検定関数・予測関数の実装

import type { CustomFormula } from './types';

// === 分布関数 ===

// NORM.DIST関数（正規分布）
export const NORM_DIST: CustomFormula = {
  name: 'NORM.DIST',
  pattern: /NORM\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// NORM.INV関数（正規分布の逆関数）
export const NORM_INV: CustomFormula = {
  name: 'NORM.INV',
  pattern: /NORM\.INV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// NORM.S.DIST関数（標準正規分布）
export const NORM_S_DIST: CustomFormula = {
  name: 'NORM.S.DIST',
  pattern: /NORM\.S\.DIST\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// NORM.S.INV関数（標準正規分布の逆関数）
export const NORM_S_INV: CustomFormula = {
  name: 'NORM.S.INV',
  pattern: /NORM\.S\.INV\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// LOGNORM.DIST関数（対数正規分布）
export const LOGNORM_DIST: CustomFormula = {
  name: 'LOGNORM.DIST',
  pattern: /LOGNORM\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// LOGNORM.INV関数（対数正規分布の逆関数）
export const LOGNORM_INV: CustomFormula = {
  name: 'LOGNORM.INV',
  pattern: /LOGNORM\.INV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// T.DIST関数（t分布・左側）
export const T_DIST: CustomFormula = {
  name: 'T.DIST',
  pattern: /T\.DIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// T.DIST.2T関数（t分布・両側）
export const T_DIST_2T: CustomFormula = {
  name: 'T.DIST.2T',
  pattern: /T\.DIST\.2T\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// T.DIST.RT関数（t分布・右側）
export const T_DIST_RT: CustomFormula = {
  name: 'T.DIST.RT',
  pattern: /T\.DIST\.RT\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// T.INV関数（t分布の逆関数・左側）
export const T_INV: CustomFormula = {
  name: 'T.INV',
  pattern: /T\.INV\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// T.INV.2T関数（t分布の逆関数・両側）
export const T_INV_2T: CustomFormula = {
  name: 'T.INV.2T',
  pattern: /T\.INV\.2T\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// CHISQ.DIST関数（カイ二乗分布）
export const CHISQ_DIST: CustomFormula = {
  name: 'CHISQ.DIST',
  pattern: /CHISQ\.DIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// CHISQ.DIST.RT関数（カイ二乗分布・右側）
export const CHISQ_DIST_RT: CustomFormula = {
  name: 'CHISQ.DIST.RT',
  pattern: /CHISQ\.DIST\.RT\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// CHISQ.INV関数（カイ二乗分布の逆関数）
export const CHISQ_INV: CustomFormula = {
  name: 'CHISQ.INV',
  pattern: /CHISQ\.INV\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// CHISQ.INV.RT関数（カイ二乗分布の逆関数・右側）
export const CHISQ_INV_RT: CustomFormula = {
  name: 'CHISQ.INV.RT',
  pattern: /CHISQ\.INV\.RT\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// F.DIST関数（F分布）
export const F_DIST: CustomFormula = {
  name: 'F.DIST',
  pattern: /F\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// F.DIST.RT関数（F分布・右側）
export const F_DIST_RT: CustomFormula = {
  name: 'F.DIST.RT',
  pattern: /F\.DIST\.RT\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// F.INV関数（F分布の逆関数）
export const F_INV: CustomFormula = {
  name: 'F.INV',
  pattern: /F\.INV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// F.INV.RT関数（F分布の逆関数・右側）
export const F_INV_RT: CustomFormula = {
  name: 'F.INV.RT',
  pattern: /F\.INV\.RT\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// BETA.DIST関数（ベータ分布）
export const BETA_DIST: CustomFormula = {
  name: 'BETA.DIST',
  pattern: /BETA\.DIST\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// BETA.INV関数（ベータ分布の逆関数）
export const BETA_INV: CustomFormula = {
  name: 'BETA.INV',
  pattern: /BETA\.INV\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// GAMMA.DIST関数（ガンマ分布）
export const GAMMA_DIST: CustomFormula = {
  name: 'GAMMA.DIST',
  pattern: /GAMMA\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// GAMMA.INV関数（ガンマ分布の逆関数）
export const GAMMA_INV: CustomFormula = {
  name: 'GAMMA.INV',
  pattern: /GAMMA\.INV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// EXPON.DIST関数（指数分布）
export const EXPON_DIST: CustomFormula = {
  name: 'EXPON.DIST',
  pattern: /EXPON\.DIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// WEIBULL.DIST関数（ワイブル分布）
export const WEIBULL_DIST: CustomFormula = {
  name: 'WEIBULL.DIST',
  pattern: /WEIBULL\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// BINOM.DIST関数（二項分布）
export const BINOM_DIST: CustomFormula = {
  name: 'BINOM.DIST',
  pattern: /BINOM\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// BINOM.INV関数（二項分布の逆関数）
export const BINOM_INV: CustomFormula = {
  name: 'BINOM.INV',
  pattern: /BINOM\.INV\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// NEGBINOM.DIST関数（負の二項分布）
export const NEGBINOM_DIST: CustomFormula = {
  name: 'NEGBINOM.DIST',
  pattern: /NEGBINOM\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// POISSON.DIST関数（ポアソン分布）
export const POISSON_DIST: CustomFormula = {
  name: 'POISSON.DIST',
  pattern: /POISSON\.DIST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// HYPGEOM.DIST関数（超幾何分布）
export const HYPGEOM_DIST: CustomFormula = {
  name: 'HYPGEOM.DIST',
  pattern: /HYPGEOM\.DIST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// === 検定・推定関数 ===

// CONFIDENCE.NORM関数（信頼区間・正規分布）
export const CONFIDENCE_NORM: CustomFormula = {
  name: 'CONFIDENCE.NORM',
  pattern: /CONFIDENCE\.NORM\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// CONFIDENCE.T関数（信頼区間・t分布）
export const CONFIDENCE_T: CustomFormula = {
  name: 'CONFIDENCE.T',
  pattern: /CONFIDENCE\.T\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// Z.TEST関数（z検定）
export const Z_TEST: CustomFormula = {
  name: 'Z.TEST',
  pattern: /Z\.TEST\(([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\)/i,
  calculate: () => null // HyperFormulaが処理
};

// T.TEST関数（t検定）
export const T_TEST: CustomFormula = {
  name: 'T.TEST',
  pattern: /T\.TEST\(([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// F.TEST関数（F検定）
export const F_TEST: CustomFormula = {
  name: 'F.TEST',
  pattern: /F\.TEST\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// CHISQ.TEST関数（カイ二乗検定）
export const CHISQ_TEST: CustomFormula = {
  name: 'CHISQ.TEST',
  pattern: /CHISQ\.TEST\(([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// === 予測・回帰関数 ===

// FORECAST関数（予測値を計算）
export const FORECAST: CustomFormula = {
  name: 'FORECAST',
  pattern: /FORECAST\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// FORECAST.LINEAR関数（線形予測）
export const FORECAST_LINEAR: CustomFormula = {
  name: 'FORECAST.LINEAR',
  pattern: /FORECAST\.LINEAR\(([^,]+),\s*([^,]+),\s*([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// FORECAST.ETS関数（指数平滑法による予測）
export const FORECAST_ETS: CustomFormula = {
  name: 'FORECAST.ETS',
  pattern: /FORECAST\.ETS\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// FORECAST.ETS.CONFINT関数（ETS予測の信頼区間）
export const FORECAST_ETS_CONFINT: CustomFormula = {
  name: 'FORECAST.ETS.CONFINT',
  pattern: /FORECAST\.ETS\.CONFINT\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// FORECAST.ETS.SEASONALITY関数（ETS季節性の長さ）
export const FORECAST_ETS_SEASONALITY: CustomFormula = {
  name: 'FORECAST.ETS.SEASONALITY',
  pattern: /FORECAST\.ETS\.SEASONALITY\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// FORECAST.ETS.STAT関数（ETS統計値）
export const FORECAST_ETS_STAT: CustomFormula = {
  name: 'FORECAST.ETS.STAT',
  pattern: /FORECAST\.ETS\.STAT\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// TREND関数（線形トレンド値）
export const TREND: CustomFormula = {
  name: 'TREND',
  pattern: /TREND\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// GROWTH関数（指数成長値）
export const GROWTH: CustomFormula = {
  name: 'GROWTH',
  pattern: /GROWTH\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// LINEST関数（線形回帰統計値）
export const LINEST: CustomFormula = {
  name: 'LINEST',
  pattern: /LINEST\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};

// LOGEST関数（指数回帰統計値）
export const LOGEST: CustomFormula = {
  name: 'LOGEST',
  pattern: /LOGEST\(([^)]+)\)/i,
  calculate: () => null // HyperFormulaが処理
};