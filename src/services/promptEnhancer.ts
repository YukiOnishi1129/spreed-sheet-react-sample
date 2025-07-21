// ユーザー入力を分析して、より具体的なプロンプトに変換する
export const enhanceUserPrompt = (userInput: string): string => {
  const input = userInput.toLowerCase().trim();
  
  // 基本的なプロンプト構造
  let enhancedPrompt = `以下のユーザー要求に基づいて、実用的なスプレッドシートを作成してください。

ユーザー要求: "${userInput}"

以下の制約を厳守してください：
1. 8行×8列の構造（必須）
2. 1行目はヘッダー行
3. 2-6行目はデータ行（5行分のサンプルデータ）
4. 7-8行目は空行または合計行
5. 循環参照を絶対に避ける

`;

  // キーワードベースでプロンプトを強化
  if (input.includes('合計') || input.includes('sum')) {
    enhancedPrompt += `
【SUM関数の要件】
- SUM関数で数値の合計を計算
- 合計行は最下部に配置
- 例：=SUM(B2:B6) で2-6行目のB列を合計

`;
  }

  if (input.includes('条件') || input.includes('if') || input.includes('判定') || input.includes('承認') || input.includes('評価') || input.includes('分類') || input.includes('チェック')) {
    enhancedPrompt += `
【IF関数の要件（必須）】
- 判定や評価は必ずIF関数を使用してください
- 固定値ではなく、数式で動的に判定してください
- 複数条件の場合はネストしたIF関数を使用
- 例：=IF(C2<=5000,"自動承認",IF(C2>10000,"要承認","確認中"))

`;
  }

  if (input.includes('検索') || input.includes('vlookup') || input.includes('lookup')) {
    enhancedPrompt += `
【VLOOKUP関数の要件】
- VLOOKUP関数で検索機能を実装
- 検索用のテーブルと検索実行部分を分ける
- 例：=VLOOKUP(E2,A2:C6,2,0)

`;
  }

  if (input.includes('ランキング') || input.includes('順位') || input.includes('rank')) {
    enhancedPrompt += `
【注意】RANK関数は使用しないでください。代わりに以下を使用：
- 大きい順に数値を配置して手動でランク表示
- または条件分岐で「上位」「中位」「下位」を判定

`;
  }

  if (input.includes('営業') || input.includes('売上')) {
    enhancedPrompt += `
【営業・売上データの例】
- A列：営業担当者名（田中、佐藤、鈴木、山田、伊藤）
- B列：売上金額（80000-150000程度の数値）
- C列：評価や判定結果

`;
  }

  if (input.includes('在庫') || input.includes('商品')) {
    enhancedPrompt += `
【在庫・商品データの例】
- A列：商品名またはコード
- B列：在庫数（10-100程度の数値）
- C列：単価
- D列：判定結果（要発注、充分など）

`;
  }

  // 共通の構造指示を追加
  enhancedPrompt += `
【重要な注意事項】
- 数式内の文字列は必ずダブルクォートで囲む
- JSON形式で正確に出力
- 数式セルは必ず "f" プロパティに記載
- 通常の値は "v" プロパティに記載
- 循環参照は絶対に避ける（セルが自分自身を参照しない）

JSONのみを返してください。説明文は不要です。`;

  return enhancedPrompt;
};