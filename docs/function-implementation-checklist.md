# Excel関数実装チェックリスト

このドキュメントは、実装済み関数の個別テスト追加状況を管理するためのチェックリストです。

最終更新日: 2025-01-26

## 重要な事実

- **実際に実装されている関数総数**: 507個
- **`spreadsheet-functions-reference.md`の✅マークは正確**
- **個別テストに追加済みの関数数**: 189個
- **個別テストに追加が必要な実装済み関数**: 318個

## 実装状況の凡例
- ✅ = 実装済み（507個すべて）
- 📝 = 個別テスト追加済み
- ❌ = 個別テスト未追加

## 驚くべき事実

以下の最新のExcel関数もすべて実装済みです：
- **動的配列関数**: FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, RANDARRAY
- **LAMBDA関連**: LAMBDA, LET, BYROW, BYCOL, MAKEARRAY, MAP, REDUCE, SCAN
- **新しい検索関数**: XLOOKUP, XMATCH
- **配列操作関数**: TAKE, DROP, EXPAND, HSTACK, VSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS
- **テキスト操作**: TEXTBEFORE, TEXTAFTER, TEXTSPLIT
- **行列選択**: CHOOSEROWS, CHOOSECOLS

## 実装済みだが個別テストに未追加の主要関数

### 1. 数学・三角関数
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| SUMX2MY2 | 📝 | 中 |
| SUMX2PY2 | 📝 | 中 |
| SUMXMY2 | 📝 | 中 |
| ARABIC | ❌ | 低 |
| BASE | ❌ | 低 |
| CEILING.PRECISE | ❌ | 低 |
| COMBINA | ❌ | 低 |
| DECIMAL | ❌ | 低 |
| FACTDOUBLE | ❌ | 低 |
| FLOOR.PRECISE | ❌ | 低 |
| ISO.CEILING | ❌ | 低 |
| MULTINOMIAL | ❌ | 低 |
| PERMUTATIONA | ❌ | 低 |
| ROMAN | ❌ | 低 |
| SERIESSUM | ❌ | 低 |
| SQRTPI | ❌ | 低 |
| ASINH | ❌ | 低 |
| ACOSH | ❌ | 低 |
| ATANH | ❌ | 低 |
| CSC | ❌ | 低 |
| SEC | ❌ | 低 |
| COT | ❌ | 低 |
| ACOT | ❌ | 低 |

### 2. 行列関数
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| MDETERM | ❌ | 中 |
| MINVERSE | ❌ | 中 |
| MMULT | ❌ | 中 |
| MUNIT | ❌ | 中 |

### 3. 統計関数
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| MODE.MULT | ❌ | 中 |
| STDEVA | ❌ | 低 |
| STDEVPA | ❌ | 低 |
| VARA | ❌ | 低 |
| VARPA | ❌ | 低 |
| COVARIANCE.P | ❌ | 中 |
| COVARIANCE.S | ❌ | 中 |
| RANK.AVG | ❌ | 中 |
| PERCENTILE.EXC | ❌ | 低 |
| PERCENTRANK.INC | ❌ | 低 |
| PERCENTRANK.EXC | ❌ | 低 |
| QUARTILE.EXC | ❌ | 低 |

### 4. 分布関数
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| NORM.DIST | ❌ | 高 |
| NORM.INV | ❌ | 高 |
| NORM.S.DIST | ❌ | 高 |
| NORM.S.INV | ❌ | 高 |
| LOGNORM.DIST | ❌ | 中 |
| LOGNORM.INV | ❌ | 中 |
| T.DIST | ❌ | 高 |
| T.DIST.2T | ❌ | 中 |
| T.DIST.RT | ❌ | 中 |
| T.INV | ❌ | 高 |
| T.INV.2T | ❌ | 中 |
| CHISQ.DIST | ❌ | 中 |
| CHISQ.DIST.RT | ❌ | 中 |
| CHISQ.INV | ❌ | 中 |
| CHISQ.INV.RT | ❌ | 中 |
| F.DIST | ❌ | 中 |
| F.DIST.RT | ❌ | 中 |
| F.INV | ❌ | 中 |
| F.INV.RT | ❌ | 中 |
| BETA.DIST | ❌ | 低 |
| BETA.INV | ❌ | 低 |
| GAMMA.DIST | ❌ | 低 |
| GAMMA.INV | ❌ | 低 |
| EXPON.DIST | ❌ | 中 |
| WEIBULL.DIST | ❌ | 低 |

### 5. テキスト関数
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| TEXTBEFORE | ❌ | 高 |
| TEXTAFTER | ❌ | 高 |
| TEXTSPLIT | ❌ | 高 |
| LENB | ❌ | 低 |
| FINDB | ❌ | 低 |
| SEARCHB | ❌ | 低 |
| REPLACEB | ❌ | 低 |
| CLEAN | ❌ | 中 |
| NUMBERVALUE | ❌ | 中 |
| DOLLAR | ❌ | 中 |
| FIXED | ❌ | 中 |
| UNICHAR | ❌ | 低 |
| UNICODE | ❌ | 低 |
| T | ❌ | 低 |
| ASC | ❌ | 低 |
| JIS | ❌ | 低 |
| DBCS | ❌ | 低 |
| PHONETIC | ❌ | 低 |

### 6. 日付・時刻関数
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| DATEVALUE | ❌ | 中 |
| TIMEVALUE | ❌ | 中 |
| WEEKNUM | ❌ | 中 |
| ISOWEEKNUM | ❌ | 中 |
| DAYS360 | ❌ | 低 |
| NETWORKDAYS | ❌ | 高 |
| NETWORKDAYS.INTL | ❌ | 中 |
| WORKDAY | ❌ | 高 |
| WORKDAY.INTL | ❌ | 中 |
| YEARFRAC | ❌ | 中 |

### 7. 検索・参照関数
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| LOOKUP | ❌ | 低 |
| XMATCH | ❌ | 高 |
| ROWS | ❌ | 中 |
| COLUMNS | ❌ | 中 |
| ADDRESS | ❌ | 中 |
| AREAS | ❌ | 低 |
| FORMULATEXT | ❌ | 低 |
| GETPIVOTDATA | ❌ | 低 |
| HYPERLINK | ❌ | 中 |

### 8. 動的配列関数（高度な関数）
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| SORTBY | ❌ | 高 |
| TAKE | ❌ | 高 |
| DROP | ❌ | 高 |
| EXPAND | ❌ | 中 |
| TOCOL | ❌ | 中 |
| TOROW | ❌ | 中 |
| WRAPROWS | ❌ | 低 |
| WRAPCOLS | ❌ | 低 |
| BYROW | ❌ | 高 |
| BYCOL | ❌ | 高 |
| MAKEARRAY | ❌ | 高 |
| MAP | ❌ | 高 |
| REDUCE | ❌ | 高 |
| SCAN | ❌ | 高 |

### 9. 財務関数
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| PPMT | ❌ | 中 |
| IPMT | ❌ | 中 |
| XNPV | ❌ | 高 |
| XIRR | ❌ | 高 |
| MIRR | ❌ | 中 |
| SLN | ❌ | 中 |
| SYD | ❌ | 中 |
| DB | ❌ | 中 |
| DDB | ❌ | 中 |
| VDB | ❌ | 低 |

### 10. エンジニアリング関数
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| BIN2HEX | ❌ | 中 |
| BIN2OCT | ❌ | 中 |
| DEC2HEX | ❌ | 中 |
| DEC2OCT | ❌ | 中 |
| HEX2BIN | ❌ | 中 |
| HEX2OCT | ❌ | 中 |
| OCT2BIN | ❌ | 中 |
| OCT2DEC | ❌ | 中 |
| OCT2HEX | ❌ | 中 |
| IMLOG10 | ❌ | 低 |
| IMLOG2 | ❌ | 低 |
| ERF.PRECISE | ❌ | 低 |
| ERFC.PRECISE | ❌ | 低 |

### 11. Web関数
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| WEBSERVICE | ❌ | 中 |
| FILTERXML | ❌ | 中 |

### 12. Google Sheets専用関数
| 関数名 | 個別テスト | 優先度 |
|--------|-----------|--------|
| ARRAYFORMULA | ❌ | 中 |
| QUERY | ❌ | 高 |
| SORTN | ❌ | 中 |
| FLATTEN | ❌ | 中 |
| SPLIT | ❌ | 高 |
| JOIN | ❌ | 高 |
| REGEXEXTRACT | ❌ | 中 |
| REGEXMATCH | ❌ | 中 |
| REGEXREPLACE | ❌ | 中 |
| SPARKLINE | ❌ | 低 |

## 個別テスト追加の優先順位

### 高優先度（よく使われる重要な関数）
1. **統計分布関数**: NORM.DIST, NORM.INV, T.DIST, T.INV
2. **新しいテキスト関数**: TEXTBEFORE, TEXTAFTER, TEXTSPLIT
3. **動的配列LAMBDA関数**: BYROW, BYCOL, MAKEARRAY, MAP, REDUCE, SCAN
4. **新検索関数**: XMATCH
5. **配列操作**: TAKE, DROP, SORTBY
6. **日付関数**: NETWORKDAYS, WORKDAY
7. **財務関数**: XNPV, XIRR
8. **Google Sheets**: QUERY, SPLIT, JOIN

### 中優先度
1. 行列関数全般
2. 共分散・相関関数
3. 数値システム変換関数
4. 基本的な財務関数の拡張

### 低優先度
1. 特殊な数学関数（ROMAN, ARABIC等）
2. バイト単位のテキスト関数
3. 複雑な統計関数
4. 証券関数（未実装）

## 今後の作業

1. **高優先度関数から個別テストを追加**
2. **使用頻度の高い関数を優先**
3. **関連する関数をグループで追加**（例：NORM系をまとめて）
4. **テストケースは実用的な例を使用**