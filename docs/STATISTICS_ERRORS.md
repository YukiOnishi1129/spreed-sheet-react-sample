# 統計関数のエラー状況

## エラーが発生している関数（17個）

### 未実装の関数
1. **AVERAGEIFS** - 複数条件付き平均値 → calculate: () => null
2. **COUNTIFS** - 複数条件に一致するセル数 → calculate: () => null
3. **CONFIDENCE.NORM** - 正規分布の信頼区間 → calculate: () => null
4. **CONFIDENCE.T** - T分布の信頼区間 → calculate: () => null
5. **F.TEST** - F検定 → calculate: () => null
6. **T.TEST** - T検定 → calculate: () => null
7. **Z.TEST** - Z検定 → calculate: () => null
8. **CHISQ.TEST** - カイ二乗検定 → calculate: () => null

### 実装済みだが動作に問題がある可能性のある関数
9. **MAXIFS** - 条件付き最大値 → #N/A エラー
10. **MINIFS** - 条件付き最小値 → #N/A エラー
11. **T.INV** - t分布の逆関数（左側）
12. **T.INV.2T** - t分布の逆関数（両側）
13. **F.DIST** - F分布
14. **F.INV** - F分布の逆関数
15. **PERCENTILE.EXC** - パーセンタイル（除外）
16. **HYPGEOM.DIST** - 超幾何分布

### 実装済み（修正済み）
17. **GAMMALN** - ガンマ関数の自然対数 ✓

## 修正方針

### 1. 未実装関数の実装（8個）
- AVERAGEIFS, COUNTIFS の実装
- 検定関数（CONFIDENCE.NORM, CONFIDENCE.T, F.TEST, T.TEST, Z.TEST, CHISQ.TEST）の実装

### 2. 実装済み関数の動作確認
- MAXIFS, MINIFS の #N/A エラーの原因調査
- その他の関数の動作確認

## 進捗状況
- [x] GAMMALN関数の実装完了
- [ ] AVERAGEIFS関数の実装
- [ ] COUNTIFS関数の実装
- [ ] その他の未実装関数の実装