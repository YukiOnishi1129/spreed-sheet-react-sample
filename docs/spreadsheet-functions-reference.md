# Excel & Google Sheets 関数リファレンス

このドキュメントは、ExcelとGoogle Sheetsで利用可能なすべての関数を網羅的にまとめたものです。各関数には手動計算ロジックの実装状況を示すマークを付けています。

## 凡例
- ✅ = 手動計算ロジック実装済み
- ❌ = 手動計算ロジック未実装
- 🔵 = Google Sheets専用関数
- 🟠 = Excel専用関数

## 1. 数学・三角関数 (Math & Trigonometry)

### 基本的な数学関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| SUM | 数値の合計を計算 | `=SUM(数値1, [数値2], ...)` | ✅ |
| SUMIF | 条件に一致するセルの合計 | `=SUMIF(範囲, 条件, [合計範囲])` | ✅ |
| SUMIFS | 複数条件に一致するセルの合計 | `=SUMIFS(合計範囲, 条件範囲1, 条件1, ...)` | ❌ |
| PRODUCT | 数値の積を計算 | `=PRODUCT(数値1, [数値2], ...)` | ❌ |
| SQRT | 平方根を計算 | `=SQRT(数値)` | ✅ |
| POWER | べき乗を計算 | `=POWER(数値, 指数)` | ✅ |
| ABS | 絶対値を返す | `=ABS(数値)` | ✅ |
| ROUND | 指定桁数で四捨五入 | `=ROUND(数値, 桁数)` | ✅ |
| ROUNDUP | 指定桁数で切り上げ | `=ROUNDUP(数値, 桁数)` | ✅ |
| ROUNDDOWN | 指定桁数で切り下げ | `=ROUNDDOWN(数値, 桁数)` | ✅ |
| CEILING | 基準値の倍数に切り上げ | `=CEILING(数値, 基準値)` | ✅ |
| FLOOR | 基準値の倍数に切り下げ | `=FLOOR(数値, 基準値)` | ✅ |
| MOD | 剰余を返す | `=MOD(数値, 除数)` | ✅ |
| INT | 整数部分を返す | `=INT(数値)` | ✅ |
| TRUNC | 小数部分を切り捨て | `=TRUNC(数値, [桁数])` | ✅ |
| SIGN | 数値の符号を返す | `=SIGN(数値)` | ✅ |
| RAND | 0以上1未満の乱数 | `=RAND()` | ✅ |
| RANDBETWEEN | 指定範囲の整数乱数 | `=RANDBETWEEN(最小値, 最大値)` | ✅ |
| EXP | eのべき乗を計算 | `=EXP(数値)` | ✅ |
| LN | 自然対数を計算 | `=LN(数値)` | ✅ |
| LOG | 対数を計算 | `=LOG(数値, [底])` | ✅ |
| LOG10 | 常用対数を計算 | `=LOG10(数値)` | ❌ |
| FACT | 階乗を計算 | `=FACT(数値)` | ✅ |
| COMBIN | 組み合わせ数を計算 | `=COMBIN(総数, 選択数)` | ❌ |
| PERMUT | 順列数を計算 | `=PERMUT(総数, 選択数)` | ❌ |
| GCD | 最大公約数を計算 | `=GCD(数値1, 数値2, ...)` | ❌ |
| LCM | 最小公倍数を計算 | `=LCM(数値1, 数値2, ...)` | ❌ |
| QUOTIENT | 商の整数部分を返す | `=QUOTIENT(被除数, 除数)` | ❌ |
| MROUND | 倍数に丸める | `=MROUND(数値, 倍数)` | ❌ |
| SUMSQ | 平方和を計算 | `=SUMSQ(数値1, [数値2], ...)` | ❌ |
| SUMPRODUCT | 配列の積の和を計算 | `=SUMPRODUCT(配列1, [配列2], ...)` | ❌ |
| SUMX2MY2 | x^2-y^2の和を計算 | `=SUMX2MY2(配列x, 配列y)` | ❌ |
| SUMX2PY2 | x^2+y^2の和を計算 | `=SUMX2PY2(配列x, 配列y)` | ❌ |
| SUMXMY2 | (x-y)^2の和を計算 | `=SUMXMY2(配列x, 配列y)` | ❌ |

### 三角関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| SIN | 正弦を計算 | `=SIN(角度)` | ✅ |
| COS | 余弦を計算 | `=COS(角度)` | ✅ |
| TAN | 正接を計算 | `=TAN(角度)` | ✅ |
| ASIN | 逆正弦を計算 | `=ASIN(数値)` | ✅ |
| ACOS | 逆余弦を計算 | `=ACOS(数値)` | ✅ |
| ATAN | 逆正接を計算 | `=ATAN(数値)` | ✅ |
| ATAN2 | x,y座標から角度を計算 | `=ATAN2(x座標, y座標)` | ❌ |
| SINH | 双曲線正弦を計算 | `=SINH(数値)` | ❌ |
| COSH | 双曲線余弦を計算 | `=COSH(数値)` | ❌ |
| TANH | 双曲線正接を計算 | `=TANH(数値)` | ❌ |
| ASINH | 双曲線逆正弦を計算 | `=ASINH(数値)` | ❌ |
| ACOSH | 双曲線逆余弦を計算 | `=ACOSH(数値)` | ❌ |
| ATANH | 双曲線逆正接を計算 | `=ATANH(数値)` | ❌ |
| DEGREES | ラジアンを度に変換 | `=DEGREES(角度)` | ✅ |
| RADIANS | 度をラジアンに変換 | `=RADIANS(角度)` | ✅ |
| PI | 円周率を返す | `=PI()` | ✅ |
| CSC | 余割を計算 | `=CSC(角度)` | ❌ |
| SEC | 正割を計算 | `=SEC(角度)` | ❌ |
| COT | 余接を計算 | `=COT(角度)` | ❌ |
| ACOT | 逆余接を計算 | `=ACOT(数値)` | ❌ |
| CSCH | 双曲線余割を計算 | `=CSCH(数値)` | ❌ |
| SECH | 双曲線正割を計算 | `=SECH(数値)` | ❌ |
| COTH | 双曲線余接を計算 | `=COTH(数値)` | ❌ |
| ACOTH | 双曲線逆余接を計算 | `=ACOTH(数値)` | ❌ |

### 行列関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| MDETERM | 行列式を計算 | `=MDETERM(配列)` | ❌ |
| MINVERSE | 逆行列を計算 | `=MINVERSE(配列)` | ❌ |
| MMULT | 行列の積を計算 | `=MMULT(配列1, 配列2)` | ❌ |
| MUNIT | 単位行列を作成 | `=MUNIT(サイズ)` | ❌ |

## 2. 統計関数 (Statistical Functions)

### 基本統計関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| AVERAGE | 平均値を計算 | `=AVERAGE(数値1, [数値2], ...)` | ✅ |
| AVERAGEIF | 条件付き平均値 | `=AVERAGEIF(範囲, 条件, [平均範囲])` | ✅ |
| AVERAGEIFS | 複数条件付き平均値 | `=AVERAGEIFS(平均範囲, 条件範囲1, 条件1, ...)` | ❌ |
| COUNT | 数値の個数を数える | `=COUNT(値1, [値2], ...)` | ✅ |
| COUNTA | 空白以外のセル数 | `=COUNTA(値1, [値2], ...)` | ✅ |
| COUNTBLANK | 空白セルの個数 | `=COUNTBLANK(範囲)` | ✅ |
| COUNTIF | 条件に一致するセル数 | `=COUNTIF(範囲, 条件)` | ✅ |
| COUNTIFS | 複数条件に一致するセル数 | `=COUNTIFS(条件範囲1, 条件1, ...)` | ❌ |
| MAX | 最大値を返す | `=MAX(数値1, [数値2], ...)` | ✅ |
| MIN | 最小値を返す | `=MIN(数値1, [数値2], ...)` | ✅ |
| MAXIFS | 条件付き最大値 | `=MAXIFS(最大範囲, 条件範囲1, 条件1, ...)` | ❌ |
| MINIFS | 条件付き最小値 | `=MINIFS(最小範囲, 条件範囲1, 条件1, ...)` | ❌ |
| MEDIAN | 中央値を計算 | `=MEDIAN(数値1, [数値2], ...)` | ✅ |
| MODE | 最頻値を計算 | `=MODE(数値1, [数値2], ...)` | ✅ |
| MODE.SNGL | 最頻値（単一） | `=MODE.SNGL(数値1, [数値2], ...)` | ❌ |
| MODE.MULT | 最頻値（複数） | `=MODE.MULT(数値1, [数値2], ...)` | ❌ |
| STDEV | 標準偏差（標本） | `=STDEV(数値1, [数値2], ...)` | ✅ |
| STDEV.S | 標準偏差（標本） | `=STDEV.S(数値1, [数値2], ...)` | ❌ |
| STDEV.P | 標準偏差（母集団） | `=STDEV.P(数値1, [数値2], ...)` | ❌ |
| STDEVA | 標準偏差（文字列含む） | `=STDEVA(値1, [値2], ...)` | ❌ |
| STDEVPA | 標準偏差（母集団、文字列含む） | `=STDEVPA(値1, [値2], ...)` | ❌ |
| VAR | 分散（標本） | `=VAR(数値1, [数値2], ...)` | ✅ |
| VAR.S | 分散（標本） | `=VAR.S(数値1, [数値2], ...)` | ❌ |
| VAR.P | 分散（母集団） | `=VAR.P(数値1, [数値2], ...)` | ❌ |
| VARA | 分散（文字列含む） | `=VARA(値1, [値2], ...)` | ❌ |
| VARPA | 分散（母集団、文字列含む） | `=VARPA(値1, [値2], ...)` | ❌ |

### 高度な統計関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| CORREL | 相関係数を計算 | `=CORREL(配列1, 配列2)` | ❌ |
| COVARIANCE.P | 母共分散を計算 | `=COVARIANCE.P(配列1, 配列2)` | ❌ |
| COVARIANCE.S | 標本共分散を計算 | `=COVARIANCE.S(配列1, 配列2)` | ❌ |
| DEVSQ | 偏差平方和を計算 | `=DEVSQ(数値1, [数値2], ...)` | ❌ |
| KURT | 尖度を計算 | `=KURT(数値1, [数値2], ...)` | ❌ |
| SKEW | 歪度を計算 | `=SKEW(数値1, [数値2], ...)` | ❌ |
| SKEW.P | 歪度（母集団） | `=SKEW.P(数値1, [数値2], ...)` | ❌ |
| STANDARDIZE | 標準化する | `=STANDARDIZE(x, 平均, 標準偏差)` | ❌ |
| LARGE | k番目に大きい値 | `=LARGE(配列, k)` | ✅ |
| SMALL | k番目に小さい値 | `=SMALL(配列, k)` | ✅ |
| RANK | 順位を返す | `=RANK(数値, 参照, [順序])` | ✅ |
| RANK.EQ | 順位（同順位は最小） | `=RANK.EQ(数値, 参照, [順序])` | ❌ |
| RANK.AVG | 順位（同順位は平均） | `=RANK.AVG(数値, 参照, [順序])` | ❌ |
| PERCENTILE | パーセンタイル値 | `=PERCENTILE(配列, k)` | ❌ |
| PERCENTILE.INC | パーセンタイル（境界値含む） | `=PERCENTILE.INC(配列, k)` | ❌ |
| PERCENTILE.EXC | パーセンタイル（境界値除く） | `=PERCENTILE.EXC(配列, k)` | ❌ |
| PERCENTRANK | パーセント順位 | `=PERCENTRANK(配列, x, [有効桁数])` | ❌ |
| PERCENTRANK.INC | パーセント順位（境界値含む） | `=PERCENTRANK.INC(配列, x, [有効桁数])` | ❌ |
| PERCENTRANK.EXC | パーセント順位（境界値除く） | `=PERCENTRANK.EXC(配列, x, [有効桁数])` | ❌ |
| QUARTILE | 四分位数 | `=QUARTILE(配列, 戻り値)` | ❌ |
| QUARTILE.INC | 四分位数（境界値含む） | `=QUARTILE.INC(配列, 戻り値)` | ❌ |
| QUARTILE.EXC | 四分位数（境界値除く） | `=QUARTILE.EXC(配列, 戻り値)` | ❌ |

### 分布関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| NORM.DIST | 正規分布 | `=NORM.DIST(x, 平均, 標準偏差, 累積)` | ❌ |
| NORM.INV | 正規分布の逆関数 | `=NORM.INV(確率, 平均, 標準偏差)` | ❌ |
| NORM.S.DIST | 標準正規分布 | `=NORM.S.DIST(z, 累積)` | ❌ |
| NORM.S.INV | 標準正規分布の逆関数 | `=NORM.S.INV(確率)` | ❌ |
| LOGNORM.DIST | 対数正規分布 | `=LOGNORM.DIST(x, 平均, 標準偏差, 累積)` | ❌ |
| LOGNORM.INV | 対数正規分布の逆関数 | `=LOGNORM.INV(確率, 平均, 標準偏差)` | ❌ |
| T.DIST | t分布（左側） | `=T.DIST(x, 自由度, 累積)` | ❌ |
| T.DIST.2T | t分布（両側） | `=T.DIST.2T(x, 自由度)` | ❌ |
| T.DIST.RT | t分布（右側） | `=T.DIST.RT(x, 自由度)` | ❌ |
| T.INV | t分布の逆関数（左側） | `=T.INV(確率, 自由度)` | ❌ |
| T.INV.2T | t分布の逆関数（両側） | `=T.INV.2T(確率, 自由度)` | ❌ |
| CHISQ.DIST | カイ二乗分布 | `=CHISQ.DIST(x, 自由度, 累積)` | ❌ |
| CHISQ.DIST.RT | カイ二乗分布（右側） | `=CHISQ.DIST.RT(x, 自由度)` | ❌ |
| CHISQ.INV | カイ二乗分布の逆関数 | `=CHISQ.INV(確率, 自由度)` | ❌ |
| CHISQ.INV.RT | カイ二乗分布の逆関数（右側） | `=CHISQ.INV.RT(確率, 自由度)` | ❌ |
| F.DIST | F分布 | `=F.DIST(x, 自由度1, 自由度2, 累積)` | ❌ |
| F.DIST.RT | F分布（右側） | `=F.DIST.RT(x, 自由度1, 自由度2)` | ❌ |
| F.INV | F分布の逆関数 | `=F.INV(確率, 自由度1, 自由度2)` | ❌ |
| F.INV.RT | F分布の逆関数（右側） | `=F.INV.RT(確率, 自由度1, 自由度2)` | ❌ |
| BETA.DIST | ベータ分布 | `=BETA.DIST(x, α, β, 累積, [A], [B])` | ❌ |
| BETA.INV | ベータ分布の逆関数 | `=BETA.INV(確率, α, β, [A], [B])` | ❌ |
| GAMMA.DIST | ガンマ分布 | `=GAMMA.DIST(x, α, β, 累積)` | ❌ |
| GAMMA.INV | ガンマ分布の逆関数 | `=GAMMA.INV(確率, α, β)` | ❌ |
| EXPON.DIST | 指数分布 | `=EXPON.DIST(x, λ, 累積)` | ❌ |
| WEIBULL.DIST | ワイブル分布 | `=WEIBULL.DIST(x, α, β, 累積)` | ❌ |
| BINOM.DIST | 二項分布 | `=BINOM.DIST(成功数, 試行回数, 成功率, 累積)` | ❌ |
| BINOM.INV | 二項分布の逆関数 | `=BINOM.INV(試行回数, 成功率, α)` | ❌ |
| NEGBINOM.DIST | 負の二項分布 | `=NEGBINOM.DIST(失敗数, 成功数, 成功率, 累積)` | ❌ |
| POISSON.DIST | ポアソン分布 | `=POISSON.DIST(x, 平均, 累積)` | ❌ |
| HYPGEOM.DIST | 超幾何分布 | `=HYPGEOM.DIST(標本成功数, 標本数, 母集団成功数, 母集団数, 累積)` | ❌ |

### 検定・推定関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| CONFIDENCE.NORM | 信頼区間（正規分布） | `=CONFIDENCE.NORM(α, 標準偏差, 標本数)` | ❌ |
| CONFIDENCE.T | 信頼区間（t分布） | `=CONFIDENCE.T(α, 標準偏差, 標本数)` | ❌ |
| Z.TEST | z検定 | `=Z.TEST(配列, x, [σ])` | ❌ |
| T.TEST | t検定 | `=T.TEST(配列1, 配列2, 尾部, 検定の種類)` | ❌ |
| F.TEST | F検定 | `=F.TEST(配列1, 配列2)` | ❌ |
| CHISQ.TEST | カイ二乗検定 | `=CHISQ.TEST(実測値範囲, 期待値範囲)` | ❌ |

### 予測・回帰関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| FORECAST | 予測値を計算 | `=FORECAST(x, 既知のy, 既知のx)` | ❌ |
| FORECAST.LINEAR | 線形予測 | `=FORECAST.LINEAR(x, 既知のy, 既知のx)` | ❌ |
| FORECAST.ETS | 指数平滑法による予測 | `=FORECAST.ETS(目標日付, 値, タイムライン, [季節性], [データ補完], [集計])` | ❌ |
| FORECAST.ETS.CONFINT | ETS予測の信頼区間 | `=FORECAST.ETS.CONFINT(目標日付, 値, タイムライン, [信頼度], [季節性], [データ補完], [集計])` | ❌ |
| FORECAST.ETS.SEASONALITY | ETS季節性の長さ | `=FORECAST.ETS.SEASONALITY(値, タイムライン, [データ補完], [集計])` | ❌ |
| FORECAST.ETS.STAT | ETS統計値 | `=FORECAST.ETS.STAT(値, タイムライン, 統計タイプ, [季節性], [データ補完], [集計])` | ❌ |
| TREND | 線形トレンド値 | `=TREND(既知のy, [既知のx], [新しいx], [定数])` | ❌ |
| GROWTH | 指数成長値 | `=GROWTH(既知のy, [既知のx], [新しいx], [定数])` | ❌ |
| LINEST | 線形回帰統計値 | `=LINEST(既知のy, [既知のx], [定数], [補正])` | ❌ |
| LOGEST | 指数回帰統計値 | `=LOGEST(既知のy, [既知のx], [定数], [補正])` | ❌ |
| SLOPE | 回帰直線の傾き | `=SLOPE(既知のy, 既知のx)` | ❌ |
| INTERCEPT | 回帰直線の切片 | `=INTERCEPT(既知のy, 既知のx)` | ❌ |
| RSQ | 決定係数 | `=RSQ(既知のy, 既知のx)` | ❌ |
| PEARSON | ピアソン相関係数 | `=PEARSON(配列1, 配列2)` | ❌ |
| STEYX | 回帰の標準誤差 | `=STEYX(既知のy, 既知のx)` | ❌ |

### その他の統計関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| AVEDEV | 平均偏差 | `=AVEDEV(数値1, [数値2], ...)` | ❌ |
| GEOMEAN | 幾何平均 | `=GEOMEAN(数値1, [数値2], ...)` | ❌ |
| HARMEAN | 調和平均 | `=HARMEAN(数値1, [数値2], ...)` | ❌ |
| TRIMMEAN | トリム平均 | `=TRIMMEAN(配列, 割合)` | ❌ |
| PROB | 確率の計算 | `=PROB(x範囲, 確率範囲, [下限], [上限])` | ❌ |
| FISHER | フィッシャー変換 | `=FISHER(x)` | ❌ |
| FISHERINV | フィッシャー変換の逆関数 | `=FISHERINV(y)` | ❌ |
| PHI | 標準正規分布の密度関数 | `=PHI(x)` | ❌ |
| GAUSS | ガウス関数 | `=GAUSS(z)` | ❌ |
| PERMUT | 順列 | `=PERMUT(総数, 選択数)` | ❌ |
| PERMUTATIONA | 重複順列 | `=PERMUTATIONA(総数, 選択数)` | ❌ |
| COMBINA | 重複組合せ | `=COMBINA(総数, 選択数)` | ❌ |

## 3. 文字列操作関数 (Text Functions)

| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| CONCATENATE | 文字列を結合 | `=CONCATENATE(文字列1, [文字列2], ...)` | ✅ |
| CONCAT | 文字列を結合（新版） | `=CONCAT(文字列1, [文字列2], ...)` | ✅ |
| TEXTJOIN | 区切り文字で結合 | `=TEXTJOIN(区切り文字, 空白無視, 文字列1, ...)` | ✅ |
| LEFT | 左から文字を抽出 | `=LEFT(文字列, [文字数])` | ✅ |
| RIGHT | 右から文字を抽出 | `=RIGHT(文字列, [文字数])` | ✅ |
| MID | 中間の文字を抽出 | `=MID(文字列, 開始位置, 文字数)` | ✅ |
| LEN | 文字数を返す | `=LEN(文字列)` | ✅ |
| LENB | バイト数を返す | `=LENB(文字列)` | ❌ |
| FIND | 文字位置を検索（大小区別） | `=FIND(検索文字列, 対象文字列, [開始位置])` | ✅ |
| FINDB | バイト位置を検索 | `=FINDB(検索文字列, 対象文字列, [開始位置])` | ❌ |
| SEARCH | 文字位置を検索（大小区別なし） | `=SEARCH(検索文字列, 対象文字列, [開始位置])` | ✅ |
| SEARCHB | バイト位置を検索 | `=SEARCHB(検索文字列, 対象文字列, [開始位置])` | ❌ |
| REPLACE | 文字を置換（位置指定） | `=REPLACE(文字列, 開始位置, 文字数, 新文字列)` | ❌ |
| REPLACEB | バイト単位で置換 | `=REPLACEB(文字列, 開始位置, バイト数, 新文字列)` | ❌ |
| SUBSTITUTE | 文字を置換（文字指定） | `=SUBSTITUTE(文字列, 検索文字列, 置換文字列, [置換対象])` | ✅ |
| UPPER | 大文字に変換 | `=UPPER(文字列)` | ✅ |
| LOWER | 小文字に変換 | `=LOWER(文字列)` | ✅ |
| PROPER | 先頭を大文字に変換 | `=PROPER(文字列)` | ✅ |
| TRIM | 余分なスペースを削除 | `=TRIM(文字列)` | ✅ |
| CLEAN | 印刷不可文字を削除 | `=CLEAN(文字列)` | ❌ |
| TEXT | 数値を書式付き文字列に変換 | `=TEXT(値, 表示形式)` | ✅ |
| VALUE | 文字列を数値に変換 | `=VALUE(文字列)` | ✅ |
| NUMBERVALUE | 文字列を数値に変換（地域設定対応） | `=NUMBERVALUE(文字列, [小数点], [桁区切り])` | ❌ |
| DOLLAR | 通貨書式に変換 | `=DOLLAR(数値, [桁数])` | ❌ |
| FIXED | 固定小数点表示 | `=FIXED(数値, [桁数], [桁区切りなし])` | ❌ |
| REPT | 文字列を繰り返す | `=REPT(文字列, 繰り返し回数)` | ✅ |
| CHAR | 文字コードから文字を返す | `=CHAR(数値)` | ❌ |
| UNICHAR | Unicode文字を返す | `=UNICHAR(数値)` | ❌ |
| CODE | 文字から文字コードを返す | `=CODE(文字列)` | ❌ |
| UNICODE | Unicode値を返す | `=UNICODE(文字列)` | ❌ |
| EXACT | 文字列が同一か判定 | `=EXACT(文字列1, 文字列2)` | ❌ |
| T | 文字列を返す | `=T(値)` | ❌ |
| TEXTBEFORE | 区切り文字の前を抽出 | `=TEXTBEFORE(文字列, 区切り文字, [インスタンス番号], [一致モード], [一致しない場合], [大文字小文字])` | ❌ |
| TEXTAFTER | 区切り文字の後を抽出 | `=TEXTAFTER(文字列, 区切り文字, [インスタンス番号], [一致モード], [一致しない場合], [大文字小文字])` | ❌ |
| TEXTSPLIT | 文字列を分割 | `=TEXTSPLIT(文字列, 列区切り文字, [行区切り文字], [空白無視], [一致モード], [余白を埋める])` | ❌ |
| ASC | 全角を半角に変換 | `=ASC(文字列)` | ❌ |
| JIS | 半角を全角に変換 | `=JIS(文字列)` | ❌ |
| DBCS | 半角を全角に変換 | `=DBCS(文字列)` | ❌ |
| PHONETIC | ふりがなを抽出 | `=PHONETIC(参照)` | ❌ |
| BAHTTEXT | 数値をタイ語文字列に変換 | `=BAHTTEXT(数値)` | ❌ |
| ARABIC | ローマ数字をアラビア数字に変換 | `=ARABIC(文字列)` | ❌ |
| ROMAN | アラビア数字をローマ数字に変換 | `=ROMAN(数値, [書式])` | ❌ |

## 4. 日付・時刻関数 (Date & Time Functions)

### 基本的な日付・時刻関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| TODAY | 今日の日付を返す | `=TODAY()` | ✅ |
| NOW | 現在の日時を返す | `=NOW()` | ✅ |
| DATE | 日付を作成 | `=DATE(年, 月, 日)` | ✅ |
| TIME | 時刻を作成 | `=TIME(時, 分, 秒)` | ❌ |
| DATEVALUE | 日付文字列を日付値に変換 | `=DATEVALUE(日付文字列)` | ❌ |
| TIMEVALUE | 時刻文字列を時刻値に変換 | `=TIMEVALUE(時刻文字列)` | ❌ |
| YEAR | 年を抽出 | `=YEAR(日付)` | ✅ |
| MONTH | 月を抽出 | `=MONTH(日付)` | ✅ |
| DAY | 日を抽出 | `=DAY(日付)` | ✅ |
| HOUR | 時を抽出 | `=HOUR(時刻)` | ❌ |
| MINUTE | 分を抽出 | `=MINUTE(時刻)` | ❌ |
| SECOND | 秒を抽出 | `=SECOND(時刻)` | ❌ |
| WEEKDAY | 曜日を数値で返す | `=WEEKDAY(日付, [種類])` | ✅ |
| WEEKNUM | 週番号を返す | `=WEEKNUM(日付, [週の基準])` | ❌ |
| ISOWEEKNUM | ISO週番号を返す | `=ISOWEEKNUM(日付)` | ❌ |

### 日付計算関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| DATEDIF | 日付間の期間を計算 | `=DATEDIF(開始日, 終了日, 単位)` | ✅ |
| DAYS | 日数差を計算 | `=DAYS(終了日, 開始日)` | ✅ |
| DAYS360 | 360日基準の日数 | `=DAYS360(開始日, 終了日, [方式])` | ❌ |
| EDATE | 月数後の日付 | `=EDATE(開始日, 月数)` | ✅ |
| EOMONTH | 月末日を返す | `=EOMONTH(開始日, 月数)` | ❌ |
| NETWORKDAYS | 稼働日数を計算 | `=NETWORKDAYS(開始日, 終了日, [祝日])` | ✅ |
| NETWORKDAYS.INTL | 稼働日数（国際版） | `=NETWORKDAYS.INTL(開始日, 終了日, [週末], [祝日])` | ❌ |
| WORKDAY | 稼働日を計算 | `=WORKDAY(開始日, 日数, [祝日])` | ❌ |
| WORKDAY.INTL | 稼働日（国際版） | `=WORKDAY.INTL(開始日, 日数, [週末], [祝日])` | ❌ |
| YEARFRAC | 年の割合を計算 | `=YEARFRAC(開始日, 終了日, [基準])` | ❌ |

## 5. 論理関数 (Logical Functions)

| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| IF | 条件分岐 | `=IF(論理式, 真の場合, 偽の場合)` | ✅ |
| IFS | 複数条件分岐 | `=IFS(条件1, 値1, [条件2, 値2], ...)` | ✅ |
| AND | 論理積（すべて真） | `=AND(論理値1, [論理値2], ...)` | ✅ |
| OR | 論理和（いずれか真） | `=OR(論理値1, [論理値2], ...)` | ✅ |
| NOT | 論理否定 | `=NOT(論理値)` | ✅ |
| XOR | 排他的論理和 | `=XOR(論理値1, [論理値2], ...)` | ✅ |
| TRUE | 真を返す | `=TRUE()` | ✅ |
| FALSE | 偽を返す | `=FALSE()` | ✅ |
| IFERROR | エラー時の値を指定 | `=IFERROR(値, エラーの場合の値)` | ✅ |
| IFNA | #N/Aエラー時の値 | `=IFNA(値, #N/Aの場合の値)` | ✅ |
| SWITCH | 値に応じて切り替え | `=SWITCH(式, 値1, 結果1, [値2, 結果2], ..., [既定])` | ❌ |
| LAMBDA | カスタム関数を作成 | `=LAMBDA([パラメータ1, ...], 計算式)` | ❌ |
| LET | 変数を定義して使用 | `=LET(名前1, 値1, 計算式)` | ❌ |

## 6. 検索・参照関数 (Lookup & Reference Functions)

### 基本的な検索関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| VLOOKUP | 垂直方向検索 | `=VLOOKUP(検索値, 範囲, 列番号, [検索方法])` | ✅ |
| HLOOKUP | 水平方向検索 | `=HLOOKUP(検索値, 範囲, 行番号, [検索方法])` | ✅ |
| XLOOKUP | 柔軟な検索 | `=XLOOKUP(検索値, 検索範囲, 戻り範囲, [見つからない場合], [一致モード], [検索モード])` | ✅ |
| LOOKUP | 検索（旧式） | `=LOOKUP(検索値, 検索範囲, [結果範囲])` | ✅ |
| INDEX | 位置から値を取得 | `=INDEX(配列, 行番号, [列番号])` | ✅ |
| MATCH | 値の位置を検索 | `=MATCH(検索値, 検索範囲, [照合の種類])` | ✅ |
| XMATCH | 値の位置を検索（拡張版） | `=XMATCH(検索値, 検索配列, [一致モード], [検索モード])` | ❌ |

### 参照関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| CHOOSE | リストから選択 | `=CHOOSE(インデックス番号, 値1, [値2], ...)` | ❌ |
| CHOOSEROWS | 行を選択 | `=CHOOSEROWS(配列, 行番号1, [行番号2], ...)` | ❌ |
| CHOOSECOLS | 列を選択 | `=CHOOSECOLS(配列, 列番号1, [列番号2], ...)` | ❌ |
| OFFSET | オフセット参照 | `=OFFSET(参照, 行数, 列数, [高さ], [幅])` | ❌ |
| INDIRECT | 文字列から参照を作成 | `=INDIRECT(参照文字列, [参照形式])` | ❌ |
| ROW | 行番号を返す | `=ROW([参照])` | ❌ |
| ROWS | 行数を返す | `=ROWS(配列)` | ❌ |
| COLUMN | 列番号を返す | `=COLUMN([参照])` | ❌ |
| COLUMNS | 列数を返す | `=COLUMNS(配列)` | ❌ |
| ADDRESS | セルアドレスを作成 | `=ADDRESS(行番号, 列番号, [絶対参照], [参照形式], [シート名])` | ❌ |
| AREAS | 領域数を返す | `=AREAS(参照)` | ❌ |
| FORMULATEXT | 数式を文字列で返す | `=FORMULATEXT(参照)` | ❌ |
| GETPIVOTDATA | ピボットテーブルからデータ取得 | `=GETPIVOTDATA(データフィールド, ピボットテーブル, [フィールド1, アイテム1], ...)` | ❌ |
| HYPERLINK | ハイパーリンクを作成 | `=HYPERLINK(リンク先, [表示文字列])` | ❌ |
| TRANSPOSE | 行列を入れ替え | `=TRANSPOSE(配列)` | ❌ |

### 動的配列関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| FILTER | 条件でフィルタ | `=FILTER(配列, 条件, [空の場合])` | ❌ |
| SORT | 並べ替え | `=SORT(配列, [並べ替えインデックス], [並べ替え順序], [列で並べ替え])` | ❌ |
| SORTBY | 別の配列で並べ替え | `=SORTBY(配列, 基準配列1, [並べ替え順序1], ...)` | ❌ |
| UNIQUE | 一意の値を抽出 | `=UNIQUE(配列, [列の比較], [1回のみ])` | ❌ |
| SEQUENCE | 連続値を生成 | `=SEQUENCE(行, [列], [開始], [ステップ])` | ❌ |
| RANDARRAY | ランダム配列を生成 | `=RANDARRAY([行], [列], [最小], [最大], [整数])` | ❌ |
| TAKE | 行/列を取得 | `=TAKE(配列, 行, [列])` | ❌ |
| DROP | 行/列を除外 | `=DROP(配列, 行, [列])` | ❌ |
| EXPAND | 配列を拡張 | `=EXPAND(配列, 行, [列], [パディング値])` | ❌ |
| HSTACK | 水平方向に結合 | `=HSTACK(配列1, [配列2], ...)` | ❌ |
| VSTACK | 垂直方向に結合 | `=VSTACK(配列1, [配列2], ...)` | ❌ |
| TOCOL | 列に変換 | `=TOCOL(配列, [無視], [スキャンバイ列])` | ❌ |
| TOROW | 行に変換 | `=TOROW(配列, [無視], [スキャンバイ列])` | ❌ |
| WRAPROWS | 行で折り返し | `=WRAPROWS(ベクトル, 折り返し数, [パディング値])` | ❌ |
| WRAPCOLS | 列で折り返し | `=WRAPCOLS(ベクトル, 折り返し数, [パディング値])` | ❌ |

## 7. 情報関数 (Information Functions)

| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| ISBLANK | 空白セルか判定 | `=ISBLANK(値)` | ✅ |
| ISERROR | エラー値か判定 | `=ISERROR(値)` | ✅ |
| ISERR | エラー値か判定（#N/A以外） | `=ISERR(値)` | ❌ |
| ISNA | #N/Aエラーか判定 | `=ISNA(値)` | ✅ |
| ISTEXT | 文字列か判定 | `=ISTEXT(値)` | ✅ |
| ISNONTEXT | 文字列以外か判定 | `=ISNONTEXT(値)` | ❌ |
| ISNUMBER | 数値か判定 | `=ISNUMBER(値)` | ✅ |
| ISLOGICAL | 論理値か判定 | `=ISLOGICAL(値)` | ✅ |
| ISREF | 参照か判定 | `=ISREF(値)` | ❌ |
| ISFORMULA | 数式か判定 | `=ISFORMULA(参照)` | ❌ |
| ISEVEN | 偶数か判定 | `=ISEVEN(数値)` | ✅ |
| ISODD | 奇数か判定 | `=ISODD(数値)` | ✅ |
| INFO | システム情報を返す | `=INFO(検査の種類)` | ❌ |
| TYPE | データ型を返す | `=TYPE(値)` | ✅ |
| N | 数値に変換 | `=N(値)` | ✅ |
| NA | #N/Aエラーを返す | `=NA()` | ❌ |
| ERROR.TYPE | エラーの種類を返す | `=ERROR.TYPE(エラー値)` | ❌ |
| SHEET | シート番号を返す | `=SHEET([値])` | ❌ |
| SHEETS | シート数を返す | `=SHEETS([参照])` | ❌ |
| CELL | セル情報を返す | `=CELL(検査の種類, [参照])` | ❌ |

## 8. データベース関数 (Database Functions)

| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| DSUM | 条件付き合計 | `=DSUM(データベース, フィールド, 条件)` | ❌ |
| DAVERAGE | 条件付き平均 | `=DAVERAGE(データベース, フィールド, 条件)` | ❌ |
| DCOUNT | 条件付きカウント | `=DCOUNT(データベース, フィールド, 条件)` | ❌ |
| DCOUNTA | 条件付きカウント（空白以外） | `=DCOUNTA(データベース, フィールド, 条件)` | ❌ |
| DMAX | 条件付き最大値 | `=DMAX(データベース, フィールド, 条件)` | ❌ |
| DMIN | 条件付き最小値 | `=DMIN(データベース, フィールド, 条件)` | ❌ |
| DPRODUCT | 条件付き積 | `=DPRODUCT(データベース, フィールド, 条件)` | ❌ |
| DSTDEV | 条件付き標準偏差（標本） | `=DSTDEV(データベース, フィールド, 条件)` | ❌ |
| DSTDEVP | 条件付き標準偏差（母集団） | `=DSTDEVP(データベース, フィールド, 条件)` | ❌ |
| DVAR | 条件付き分散（標本） | `=DVAR(データベース, フィールド, 条件)` | ❌ |
| DVARP | 条件付き分散（母集団） | `=DVARP(データベース, フィールド, 条件)` | ❌ |
| DGET | 条件に一致する値を取得 | `=DGET(データベース, フィールド, 条件)` | ❌ |

## 9. 財務関数 (Financial Functions)

### 基本的な財務関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| PV | 現在価値 | `=PV(利率, 期間, 定期支払額, [将来価値], [支払期日])` | ❌ |
| FV | 将来価値 | `=FV(利率, 期間, 定期支払額, [現在価値], [支払期日])` | ❌ |
| PMT | 定期支払額 | `=PMT(利率, 期間, 現在価値, [将来価値], [支払期日])` | ❌ |
| PPMT | 元金返済額 | `=PPMT(利率, 期, 期間, 現在価値, [将来価値], [支払期日])` | ❌ |
| IPMT | 利息支払額 | `=IPMT(利率, 期, 期間, 現在価値, [将来価値], [支払期日])` | ❌ |
| RATE | 利率 | `=RATE(期間, 定期支払額, 現在価値, [将来価値], [支払期日], [推定値])` | ❌ |
| NPER | 支払回数 | `=NPER(利率, 定期支払額, 現在価値, [将来価値], [支払期日])` | ❌ |
| NPV | 正味現在価値 | `=NPV(割引率, 値1, [値2], ...)` | ❌ |
| XNPV | 正味現在価値（日付指定） | `=XNPV(割引率, 値, 日付)` | ❌ |
| IRR | 内部収益率 | `=IRR(値, [推定値])` | ❌ |
| XIRR | 内部収益率（日付指定） | `=XIRR(値, 日付, [推定値])` | ❌ |
| MIRR | 修正内部収益率 | `=MIRR(値, 財務費用率, 再投資収益率)` | ❌ |

### 減価償却関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| SLN | 定額法 | `=SLN(取得価額, 残存価額, 耐用年数)` | ❌ |
| SYD | 級数法 | `=SYD(取得価額, 残存価額, 耐用年数, 期)` | ❌ |
| DB | 定率法 | `=DB(取得価額, 残存価額, 耐用年数, 期, [月])` | ❌ |
| DDB | 倍額定率法 | `=DDB(取得価額, 残存価額, 耐用年数, 期, [率])` | ❌ |
| VDB | 可変減価償却 | `=VDB(取得価額, 残存価額, 耐用年数, 開始期, 終了期, [率], [切り替えなし])` | ❌ |
| AMORDEGRC | フランス式減価償却 | `=AMORDEGRC(取得価額, 購入日, 第1期終了日, 残存価額, 期, 率, [基準])` | ❌ |
| AMORLINC | フランス式定額償却 | `=AMORLINC(取得価額, 購入日, 第1期終了日, 残存価額, 期, 率, [基準])` | ❌ |

### 証券関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| ACCRINT | 経過利息 | `=ACCRINT(発行日, 初回利払日, 受渡日, 利率, 額面, 頻度, [基準], [計算方式])` | ❌ |
| ACCRINTM | 満期一括払証券の経過利息 | `=ACCRINTM(発行日, 満期日, 利率, 額面, [基準])` | ❌ |
| COUPDAYBS | 直前利払日から受渡日までの日数 | `=COUPDAYBS(受渡日, 満期日, 頻度, [基準])` | ❌ |
| COUPDAYS | 利払期間の日数 | `=COUPDAYS(受渡日, 満期日, 頻度, [基準])` | ❌ |
| COUPDAYSNC | 受渡日から次回利払日までの日数 | `=COUPDAYSNC(受渡日, 満期日, 頻度, [基準])` | ❌ |
| COUPNCD | 次回利払日 | `=COUPNCD(受渡日, 満期日, 頻度, [基準])` | ❌ |
| COUPNUM | 利払回数 | `=COUPNUM(受渡日, 満期日, 頻度, [基準])` | ❌ |
| COUPPCD | 直前利払日 | `=COUPPCD(受渡日, 満期日, 頻度, [基準])` | ❌ |
| DISC | 割引率 | `=DISC(受渡日, 満期日, 価格, 償還価額, [基準])` | ❌ |
| DOLLARDE | ドル小数表記に変換 | `=DOLLARDE(小数値, 分母)` | ❌ |
| DOLLARFR | ドル分数表記に変換 | `=DOLLARFR(小数値, 分母)` | ❌ |
| DURATION | デュレーション | `=DURATION(受渡日, 満期日, 利率, 利回り, 頻度, [基準])` | ❌ |
| EFFECT | 実効年利率 | `=EFFECT(名目利率, 年間複利回数)` | ❌ |
| INTRATE | 利率 | `=INTRATE(受渡日, 満期日, 投資額, 償還価額, [基準])` | ❌ |
| MDURATION | 修正デュレーション | `=MDURATION(受渡日, 満期日, 利率, 利回り, 頻度, [基準])` | ❌ |
| NOMINAL | 名目年利率 | `=NOMINAL(実効利率, 年間複利回数)` | ❌ |
| ODDFPRICE | 変則初回期の価格 | `=ODDFPRICE(受渡日, 満期日, 発行日, 初回利払日, 利率, 利回り, 償還価額, 頻度, [基準])` | ❌ |
| ODDFYIELD | 変則初回期の利回り | `=ODDFYIELD(受渡日, 満期日, 発行日, 初回利払日, 利率, 価格, 償還価額, 頻度, [基準])` | ❌ |
| ODDLPRICE | 変則最終期の価格 | `=ODDLPRICE(受渡日, 満期日, 最終利払日, 利率, 利回り, 償還価額, 頻度, [基準])` | ❌ |
| ODDLYIELD | 変則最終期の利回り | `=ODDLYIELD(受渡日, 満期日, 最終利払日, 利率, 価格, 償還価額, 頻度, [基準])` | ❌ |
| PDURATION | 投資期間 | `=PDURATION(利率, 現在価値, 将来価値)` | ❌ |
| PRICE | 定期利付証券の価格 | `=PRICE(受渡日, 満期日, 利率, 利回り, 償還価額, 頻度, [基準])` | ❌ |
| PRICEDISC | 割引証券の価格 | `=PRICEDISC(受渡日, 満期日, 割引率, 償還価額, [基準])` | ❌ |
| PRICEMAT | 満期利付証券の価格 | `=PRICEMAT(受渡日, 満期日, 発行日, 利率, 利回り, [基準])` | ❌ |
| RECEIVED | 受取金額 | `=RECEIVED(受渡日, 満期日, 投資額, 割引率, [基準])` | ❌ |
| RRI | 投資成長率 | `=RRI(期間, 現在価値, 将来価値)` | ❌ |
| TBILLEQ | 米国財務省短期証券の債券換算利回り | `=TBILLEQ(受渡日, 満期日, 割引率)` | ❌ |
| TBILLPRICE | 米国財務省短期証券の価格 | `=TBILLPRICE(受渡日, 満期日, 割引率)` | ❌ |
| TBILLYIELD | 米国財務省短期証券の利回り | `=TBILLYIELD(受渡日, 満期日, 価格)` | ❌ |
| YIELD | 定期利付証券の利回り | `=YIELD(受渡日, 満期日, 利率, 価格, 償還価額, 頻度, [基準])` | ❌ |
| YIELDDISC | 割引証券の利回り | `=YIELDDISC(受渡日, 満期日, 価格, 償還価額, [基準])` | ❌ |
| YIELDMAT | 満期利付証券の利回り | `=YIELDMAT(受渡日, 満期日, 発行日, 利率, 価格, [基準])` | ❌ |

### その他の財務関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| CUMIPMT | 累計利息 | `=CUMIPMT(利率, 期間, 現在価値, 開始期, 終了期, 支払期日)` | ❌ |
| CUMPRINC | 累計元金 | `=CUMPRINC(利率, 期間, 現在価値, 開始期, 終了期, 支払期日)` | ❌ |
| ISPMT | 利息支払額（元金均等） | `=ISPMT(利率, 期, 期間, 現在価値)` | ❌ |

## 10. エンジニアリング関数 (Engineering Functions)

### 単位変換関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| CONVERT | 単位変換 | `=CONVERT(数値, 変換元単位, 変換先単位)` | ❌ |

### 数値システム変換関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| BIN2DEC | 2進数→10進数 | `=BIN2DEC(数値)` | ❌ |
| BIN2HEX | 2進数→16進数 | `=BIN2HEX(数値, [桁数])` | ❌ |
| BIN2OCT | 2進数→8進数 | `=BIN2OCT(数値, [桁数])` | ❌ |
| DEC2BIN | 10進数→2進数 | `=DEC2BIN(数値, [桁数])` | ❌ |
| DEC2HEX | 10進数→16進数 | `=DEC2HEX(数値, [桁数])` | ❌ |
| DEC2OCT | 10進数→8進数 | `=DEC2OCT(数値, [桁数])` | ❌ |
| HEX2BIN | 16進数→2進数 | `=HEX2BIN(数値, [桁数])` | ❌ |
| HEX2DEC | 16進数→10進数 | `=HEX2DEC(数値)` | ❌ |
| HEX2OCT | 16進数→8進数 | `=HEX2OCT(数値, [桁数])` | ❌ |
| OCT2BIN | 8進数→2進数 | `=OCT2BIN(数値, [桁数])` | ❌ |
| OCT2DEC | 8進数→10進数 | `=OCT2DEC(数値)` | ❌ |
| OCT2HEX | 8進数→16進数 | `=OCT2HEX(数値, [桁数])` | ❌ |

### 複素数関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| COMPLEX | 複素数を作成 | `=COMPLEX(実部, 虚部, [虚数単位])` | ❌ |
| IMABS | 複素数の絶対値 | `=IMABS(複素数)` | ❌ |
| IMAGINARY | 虚部を返す | `=IMAGINARY(複素数)` | ❌ |
| IMARGUMENT | 偏角を返す | `=IMARGUMENT(複素数)` | ❌ |
| IMCONJUGATE | 複素共役 | `=IMCONJUGATE(複素数)` | ❌ |
| IMCOS | 複素数の余弦 | `=IMCOS(複素数)` | ❌ |
| IMCOSH | 複素数の双曲線余弦 | `=IMCOSH(複素数)` | ❌ |
| IMCOT | 複素数の余接 | `=IMCOT(複素数)` | ❌ |
| IMCSC | 複素数の余割 | `=IMCSC(複素数)` | ❌ |
| IMCSCH | 複素数の双曲線余割 | `=IMCSCH(複素数)` | ❌ |
| IMDIV | 複素数の除算 | `=IMDIV(複素数1, 複素数2)` | ❌ |
| IMEXP | 複素数の指数関数 | `=IMEXP(複素数)` | ❌ |
| IMLN | 複素数の自然対数 | `=IMLN(複素数)` | ❌ |
| IMLOG10 | 複素数の常用対数 | `=IMLOG10(複素数)` | ❌ |
| IMLOG2 | 複素数の2を底とする対数 | `=IMLOG2(複素数)` | ❌ |
| IMPOWER | 複素数のべき乗 | `=IMPOWER(複素数, 数値)` | ❌ |
| IMPRODUCT | 複素数の積 | `=IMPRODUCT(複素数1, [複素数2], ...)` | ❌ |
| IMREAL | 実部を返す | `=IMREAL(複素数)` | ❌ |
| IMSEC | 複素数の正割 | `=IMSEC(複素数)` | ❌ |
| IMSECH | 複素数の双曲線正割 | `=IMSECH(複素数)` | ❌ |
| IMSIN | 複素数の正弦 | `=IMSIN(複素数)` | ❌ |
| IMSINH | 複素数の双曲線正弦 | `=IMSINH(複素数)` | ❌ |
| IMSQRT | 複素数の平方根 | `=IMSQRT(複素数)` | ❌ |
| IMSUB | 複素数の減算 | `=IMSUB(複素数1, 複素数2)` | ❌ |
| IMSUM | 複素数の和 | `=IMSUM(複素数1, [複素数2], ...)` | ❌ |
| IMTAN | 複素数の正接 | `=IMTAN(複素数)` | ❌ |

### ベッセル関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| BESSELI | 修正ベッセル関数In(x) | `=BESSELI(x, n)` | ❌ |
| BESSELJ | ベッセル関数Jn(x) | `=BESSELJ(x, n)` | ❌ |
| BESSELK | 修正ベッセル関数Kn(x) | `=BESSELK(x, n)` | ❌ |
| BESSELY | ベッセル関数Yn(x) | `=BESSELY(x, n)` | ❌ |

### その他のエンジニアリング関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| DELTA | クロネッカーのデルタ | `=DELTA(数値1, [数値2])` | ❌ |
| ERF | 誤差関数 | `=ERF(下限, [上限])` | ❌ |
| ERF.PRECISE | 誤差関数（精密） | `=ERF.PRECISE(x)` | ❌ |
| ERFC | 相補誤差関数 | `=ERFC(x)` | ❌ |
| ERFC.PRECISE | 相補誤差関数（精密） | `=ERFC.PRECISE(x)` | ❌ |
| GESTEP | ステップ関数 | `=GESTEP(数値, [ステップ])` | ❌ |
| BITAND | ビット単位AND | `=BITAND(数値1, 数値2)` | ❌ |
| BITOR | ビット単位OR | `=BITOR(数値1, 数値2)` | ❌ |
| BITXOR | ビット単位XOR | `=BITXOR(数値1, 数値2)` | ❌ |
| BITLSHIFT | ビット左シフト | `=BITLSHIFT(数値, シフト量)` | ❌ |
| BITRSHIFT | ビット右シフト | `=BITRSHIFT(数値, シフト量)` | ❌ |

## 11. キューブ関数 (Cube Functions) 🟠

| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| CUBEKPIMEMBER | KPIプロパティを返す | `=CUBEKPIMEMBER(接続, KPI名, KPIプロパティ, [キャプション])` | ❌ |
| CUBEMEMBER | メンバーを返す | `=CUBEMEMBER(接続, メンバー式, [キャプション])` | ❌ |
| CUBEMEMBERPROPERTY | メンバープロパティを返す | `=CUBEMEMBERPROPERTY(接続, メンバー式, プロパティ)` | ❌ |
| CUBERANKEDMEMBER | n番目のメンバーを返す | `=CUBERANKEDMEMBER(接続, セット式, ランク, [キャプション])` | ❌ |
| CUBESET | セットを定義 | `=CUBESET(接続, セット式, [キャプション], [並べ替え順序], [並べ替え基準])` | ❌ |
| CUBESETCOUNT | セット内のアイテム数 | `=CUBESETCOUNT(セット)` | ❌ |
| CUBEVALUE | 集計値を返す | `=CUBEVALUE(接続, [メンバー式1], [メンバー式2], ...)` | ❌ |

## 12. Web関数 (Web Functions)

| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| WEBSERVICE | Webサービスからデータ取得 🟠 | `=WEBSERVICE(URL)` | ❌ |
| FILTERXML | XMLからデータ抽出 🟠 | `=FILTERXML(XML, XPath)` | ❌ |
| ENCODEURL | URLエンコード 🟠 | `=ENCODEURL(文字列)` | ❌ |

## 13. 互換性関数 (Compatibility Functions)

これらの関数は後方互換性のために残されています。新しいバージョンの関数を使用することを推奨します。

| 関数名 | 説明 | 新しい関数 | 実装状況 |
|--------|------|----------|----------|
| BETADIST | ベータ分布 | BETA.DIST | ❌ |
| BETAINV | ベータ分布の逆関数 | BETA.INV | ❌ |
| BINOMDIST | 二項分布 | BINOM.DIST | ❌ |
| CHIDIST | カイ二乗分布 | CHISQ.DIST.RT | ❌ |
| CHIINV | カイ二乗分布の逆関数 | CHISQ.INV.RT | ❌ |
| CHITEST | カイ二乗検定 | CHISQ.TEST | ❌ |
| CONFIDENCE | 信頼区間 | CONFIDENCE.NORM | ❌ |
| COVAR | 共分散 | COVARIANCE.P | ❌ |
| CRITBINOM | 二項分布の臨界値 | BINOM.INV | ❌ |
| EXPONDIST | 指数分布 | EXPON.DIST | ❌ |
| FDIST | F分布 | F.DIST.RT | ❌ |
| FINV | F分布の逆関数 | F.INV.RT | ❌ |
| FTEST | F検定 | F.TEST | ❌ |
| GAMMADIST | ガンマ分布 | GAMMA.DIST | ❌ |
| GAMMAINV | ガンマ分布の逆関数 | GAMMA.INV | ❌ |
| HYPGEOMDIST | 超幾何分布 | HYPGEOM.DIST | ❌ |
| LOGINV | 対数正規分布の逆関数 | LOGNORM.INV | ❌ |
| LOGNORMDIST | 対数正規分布 | LOGNORM.DIST | ❌ |
| MODE | 最頻値 | MODE.SNGL | ❌ |
| NEGBINOMDIST | 負の二項分布 | NEGBINOM.DIST | ❌ |
| NORMDIST | 正規分布 | NORM.DIST | ❌ |
| NORMINV | 正規分布の逆関数 | NORM.INV | ❌ |
| NORMSDIST | 標準正規分布 | NORM.S.DIST | ❌ |
| NORMSINV | 標準正規分布の逆関数 | NORM.S.INV | ❌ |
| PERCENTILE | パーセンタイル | PERCENTILE.INC | ❌ |
| PERCENTRANK | パーセント順位 | PERCENTRANK.INC | ❌ |
| POISSON | ポアソン分布 | POISSON.DIST | ❌ |
| QUARTILE | 四分位数 | QUARTILE.INC | ❌ |
| RANK | 順位 | RANK.EQ | ❌ |
| STDEV | 標準偏差 | STDEV.S | ❌ |
| STDEVP | 標準偏差（母集団） | STDEV.P | ❌ |
| TDIST | t分布 | T.DIST.2T | ❌ |
| TINV | t分布の逆関数 | T.INV.2T | ❌ |
| TTEST | t検定 | T.TEST | ❌ |
| VAR | 分散 | VAR.S | ❌ |
| VARP | 分散（母集団） | VAR.P | ❌ |
| WEIBULL | ワイブル分布 | WEIBULL.DIST | ❌ |
| ZTEST | z検定 | Z.TEST | ❌ |

## 14. Google Sheets専用関数 🔵

### 配列・フィルタ関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| ARRAYFORMULA | 配列数式を適用 | `=ARRAYFORMULA(配列数式)` | ❌ |
| QUERY | SQLライクなクエリ実行 | `=QUERY(データ, クエリ, [ヘッダー])` | ❌ |
| SORTN | 上位N件を並べ替えて返す | `=SORTN(範囲, [n], [表示タイ], [並べ替えインデックス], [昇順])` | ❌ |
| FLATTEN | 配列を1次元に変換 | `=FLATTEN(範囲1, [範囲2], ...)` | ❌ |

### データ取得関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| IMPORTDATA | CSVやTSVをインポート | `=IMPORTDATA(URL)` | ❌ |
| IMPORTFEED | RSSやAtomフィードを取得 | `=IMPORTFEED(URL, [クエリ], [ヘッダー], [アイテム数])` | ❌ |
| IMPORTHTML | HTMLテーブルやリストを取得 | `=IMPORTHTML(URL, クエリ, インデックス)` | ❌ |
| IMPORTXML | XMLデータを取得 | `=IMPORTXML(URL, XPathクエリ)` | ❌ |
| IMPORTRANGE | 他のスプレッドシートから取得 | `=IMPORTRANGE(スプレッドシートURL, 範囲文字列)` | ❌ |
| IMAGE | 画像を挿入 | `=IMAGE(URL, [モード], [高さ], [幅])` | ❌ |

### Google固有の関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| GOOGLEFINANCE | 株価情報を取得 | `=GOOGLEFINANCE(銘柄, [属性], [開始日], [終了日], [間隔])` | ❌ |
| GOOGLETRANSLATE | テキストを翻訳 | `=GOOGLETRANSLATE(テキスト, ソース言語, ターゲット言語)` | ❌ |
| DETECTLANGUAGE | 言語を検出 | `=DETECTLANGUAGE(テキスト)` | ❌ |
| SPARKLINE | スパークラインを作成 | `=SPARKLINE(データ, [オプション])` | ❌ |

### その他のGoogle Sheets関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| SPLIT | 文字列を分割 | `=SPLIT(テキスト, 区切り文字, [各文字で分割], [空のテキストを削除])` | ✅ |
| JOIN | 文字列を結合 | `=JOIN(区切り文字, 値1, [値2], ...)` | ❌ |
| REGEXEXTRACT | 正規表現で抽出 | `=REGEXEXTRACT(テキスト, 正規表現)` | ❌ |
| REGEXMATCH | 正規表現でマッチ判定 | `=REGEXMATCH(テキスト, 正規表現)` | ❌ |
| REGEXREPLACE | 正規表現で置換 | `=REGEXREPLACE(テキスト, 正規表現, 置換テキスト)` | ❌ |
| TO_DATE | 数値を日付に変換 | `=TO_DATE(値)` | ❌ |
| TO_DOLLARS | 数値をドル表記に変換 | `=TO_DOLLARS(値)` | ❌ |
| TO_PERCENT | 数値をパーセント表記に変換 | `=TO_PERCENT(値)` | ❌ |
| TO_TEXT | 値をテキストに変換 | `=TO_TEXT(値)` | ❌ |
| ISBETWEEN | 値が範囲内か判定 | `=ISBETWEEN(値, 下限, 上限, [下限含む], [上限含む])` | ❌ |

## 15. 新しい関数・その他

### Excel 365の新関数
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| BYROW | 行ごとにLAMBDA適用 | `=BYROW(配列, LAMBDA)` | ❌ |
| BYCOL | 列ごとにLAMBDA適用 | `=BYCOL(配列, LAMBDA)` | ❌ |
| MAKEARRAY | 配列を生成 | `=MAKEARRAY(行, 列, LAMBDA)` | ❌ |
| MAP | 配列の各要素にLAMBDA適用 | `=MAP(配列1, [配列2], ..., LAMBDA)` | ❌ |
| REDUCE | 配列を累積処理 | `=REDUCE(初期値, 配列, LAMBDA)` | ❌ |
| SCAN | 配列をスキャン | `=SCAN(初期値, 配列, LAMBDA)` | ❌ |
| ISOMITTED | 引数が省略されたか判定 | `=ISOMITTED(引数)` | ❌ |
| STOCKHISTORY | 株価履歴を取得 | `=STOCKHISTORY(銘柄, 開始日, [終了日], [間隔], [列], [プロパティ])` | ❌ |

### AI関連関数（Excel Labs）
| 関数名 | 説明 | 構文 | 実装状況 |
|--------|------|------|----------|
| GPT | ChatGPTを呼び出し | `=GPT(プロンプト, [温度], [最大トークン])` | ❌ |

## まとめ

このドキュメントには、ExcelとGoogle Sheetsで利用可能な主要な関数をカテゴリ別にまとめました。

### 統計情報  
- 総関数数: 約500+
- カテゴリ数: 15
- **実装済み: 101関数**（SUM, AVERAGE, COUNT, MAX, MIN, IF, ABS, SQRT, POWER, MOD, INT, TRUNC, SIN, COS, TAN, LOG, EXP, RANDBETWEEN, DEGREES, RADIANS等）
- 未実装: 400+関数

## 実装状況サマリー

現在実装済みの関数数: **約100関数** ✅

### 実装済みカテゴリ
- **数学・三角関数**: ABS, SQRT, POWER, ROUND系, MOD, 三角関数, 対数関数など
- **統計関数**: SUM, AVERAGE, COUNT系, MEDIAN, MODE, STDEV, VAR, LARGE, SMALL, RANKなど  
- **テキスト関数**: CONCATENATE, LEFT/RIGHT/MID, FIND/SEARCH, SUBSTITUTE, 大小文字変換など
- **日付関数**: TODAY, NOW, DATE, YEAR/MONTH/DAY, WEEKDAY, DATEDIF, DAYS, EDATEなど
- **論理関数**: IF, IFS, AND, OR, NOT, XOR, TRUE/FALSE, IFERROR, IFNAなど
- **検索関数**: VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, LOOKUPなど
- **情報関数**: ISBLANK, ISERROR, ISTEXT, ISNUMBER, TYPE, Nなど

**注意**: 実装状況は2024年7月時点での /src/utils/formulas/ ディレクトリの内容に基づいて更新されています。

### 次のステップ
1. 各関数の手動計算ロジックの実装優先順位を決定
2. 使用頻度の高い関数から順次実装
3. テストケースの作成と検証
4. パフォーマンスの最適化

注意: このリストは主要な関数を網羅していますが、完全なリストではありません。ExcelとGoogle Sheetsは継続的に新機能を追加しているため、最新の関数については公式ドキュメントを参照してください。