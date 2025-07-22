export interface FunctionTemplate {
  id: string;
  title: string;
  description: string;
  category: string;
  icon: string;
  functions: string[];
  prompt: string;
  difficulty: 'beginner' | 'intermediate' | 'advanced';
  tags: string[];
  // 固定のスプレッドシートデータ（ChatGPT APIに依存しない）
  fixedData?: {
    function_name: string;
    description: string;
    syntax: string;
    syntax_detail?: string;
    category: string;
    spreadsheet_data: any[][];
    examples: string[];
  };
}

export interface TemplateCategory {
  id: string;
  name: string;
  icon: string;
  description: string;
  templates: FunctionTemplate[];
}

export const TEMPLATE_CATEGORIES: TemplateCategory[] = [
  {
    id: 'sales',
    name: '売上分析',
    icon: '📊',
    description: '売上データの分析と集計',
    templates: [
      {
        id: 'sales-performance',
        title: '営業成績評価',
        description: 'IF関数で目標達成判定、SUM関数で合計計算',
        category: 'sales',
        icon: '🎯',
        functions: ['IF', 'SUM', 'AVERAGE'],
        prompt: '営業担当者の売上管理表を作成してください。以下の要件で8行8列のスプレッドシートを作成してください：1) 営業担当者名、売上金額、評価、合計の列を含める 2) IF関数で売上100000以上なら「目標達成」、未満なら「要改善」と表示 3) SUM関数で全体の売上合計を計算 4) 5人の営業担当者データを含める',
        difficulty: 'beginner',
        tags: ['売上', '目標管理', 'KPI'],
        fixedData: {
          function_name: 'IF & SUM',
          description: '営業担当者の売上を評価し、合計を計算します',
          syntax: 'IF(logical_test, value_if_true, value_if_false) & SUM(range)',
          category: '論理関数 & 数学関数',
          spreadsheet_data: [
            [
              { v: "営業担当者", ct: { t: "s" }, bg: "#E3F2FD" },
              { v: "売上", ct: { t: "s" }, bg: "#E3F2FD" },
              { v: "評価", ct: { t: "s" }, bg: "#E3F2FD" },
              null, null, null, null, null
            ],
            [
              { v: "田中", ct: { t: "s" } },
              { v: 120000, ct: { t: "n" } },
              { v: null, f: "=IF(B2>=100000,\"目標達成\",\"要改善\")", bg: "#E8F5E8", fc: "#2E7D32" },
              null, null, null, null, null
            ],
            [
              { v: "佐藤", ct: { t: "s" } },
              { v: 85000, ct: { t: "n" } },
              { v: null, f: "=IF(B3>=100000,\"目標達成\",\"要改善\")", bg: "#E8F5E8", fc: "#2E7D32" },
              null, null, null, null, null
            ],
            [
              { v: "鈴木", ct: { t: "s" } },
              { v: 150000, ct: { t: "n" } },
              { v: null, f: "=IF(B4>=100000,\"目標達成\",\"要改善\")", bg: "#E8F5E8", fc: "#2E7D32" },
              null, null, null, null, null
            ],
            [
              { v: "山田", ct: { t: "s" } },
              { v: 95000, ct: { t: "n" } },
              { v: null, f: "=IF(B5>=100000,\"目標達成\",\"要改善\")", bg: "#E8F5E8", fc: "#2E7D32" },
              null, null, null, null, null
            ],
            [
              { v: "合計", ct: { t: "s" }, bg: "#E8F5E8" },
              { v: null, f: "=SUM(B2:B5)", bg: "#FFE0B2", fc: "#D84315" },
              null, null, null, null, null, null
            ],
            [null, null, null, null, null, null, null, null],
            [null, null, null, null, null, null, null, null]
          ],
          examples: ['=IF(B2>=100000,"目標達成","要改善")', '=SUM(B2:B5)']
        }
      },
      {
        id: 'sales-ranking',
        title: '売上評価（条件分岐）',
        description: 'IF関数で売上レベル判定、COUNTIF関数で人数集計',
        category: 'sales',
        icon: '🏆',
        functions: ['IF', 'COUNTIF', 'SUM'],
        prompt: '営業担当者の売上評価表を作成してください。8行8列の構造で以下を含めてください：1) A列：営業担当者名（5人分） 2) B列：売上金額 3) C列：IF関数で評価（120000以上=優秀、100000以上=良い、未満=普通） 4) 最下行：SUM関数でB列の合計。循環参照を避けてください。',
        difficulty: 'intermediate',
        tags: ['評価', '条件分岐', '売上']
      }
    ]
  },
  {
    id: 'inventory',
    name: '在庫管理',
    icon: '📋',
    description: '在庫データの管理と発注判定',
    templates: [
      {
        id: 'stock-alert',
        title: '在庫アラート',
        description: 'IF関数で発注判定、COUNTIF関数で発注商品数',
        category: 'inventory',
        icon: '⚠️',
        functions: ['IF', 'COUNTIF', 'SUM'],
        prompt: '商品の在庫管理表を作成してください。在庫が10個以下なら「要発注」、20個以上なら「充分」、その間は「普通」とIF関数で判定し、COUNTIF関数で要発注商品の数をカウントするスプレッドシートを作成してください。',
        difficulty: 'beginner',
        tags: ['在庫', '発注', 'アラート']
      },
      {
        id: 'inventory-value',
        title: '在庫金額計算',
        description: '在庫数×単価の計算とVLOOKUP検索',
        category: 'inventory',
        icon: '💰',
        functions: ['VLOOKUP', 'SUM', 'IF'],
        prompt: '商品マスタと在庫データを使用して、VLOOKUP関数で商品名と単価を検索し、在庫数×単価で在庫金額を計算するスプレッドシートを作成してください。全体の在庫金額合計も表示してください。',
        difficulty: 'intermediate',
        tags: ['在庫', '金額計算', '検索']
      }
    ]
  },
  {
    id: 'hr',
    name: '人事評価',
    icon: '👥',
    description: '社員データの評価と分析',
    templates: [
      {
        id: 'performance-review',
        title: '成績評価',
        description: 'IF関数で等級判定、COUNTIF関数で人数集計',
        category: 'hr',
        icon: '📝',
        functions: ['IF', 'COUNTIF', 'AVERAGE'],
        prompt: '社員の成績評価表を作成してください。点数が90点以上は「S」、80点以上は「A」、70点以上は「B」、未満は「C」とIF関数で判定し、COUNTIF関数で各等級の人数をカウントするスプレッドシートを作成してください。',
        difficulty: 'beginner',
        tags: ['評価', '成績', '人事']
      },
      {
        id: 'overtime-calc',
        title: '残業代計算',
        description: 'IF関数で残業判定、複雑な時給計算',
        category: 'hr',
        icon: '⏰',
        functions: ['IF', 'SUM', 'VLOOKUP'],
        prompt: '社員の勤務時間管理表を作成してください。IF関数で勤務時間8時間超過を判定し、8時間を超えた分の残業時間を計算し、残業時間×1.25で残業手当係数を算出するスプレッドシートを作成してください。',
        difficulty: 'advanced',
        tags: ['残業', '給与計算', '時間管理']
      },
      {
        id: 'employee-tenure-vacation',
        title: '従業員勤続年数・有給計算',
        description: 'DATEDIF関数で勤続年数計算、NETWORKDAYS関数で営業日数計算',
        category: 'hr',
        icon: '📅',
        functions: ['DATEDIF', 'NETWORKDAYS', 'TODAY', 'IF'],
        prompt: '従業員の勤続年数と有給休暇を管理する表を作成してください。DATEDIF関数で勤続年数を計算し、NETWORKDAYS関数で営業日数を計算してください。',
        difficulty: 'intermediate',
        tags: ['勤続年数', '有給休暇', '営業日数', '人事管理'],
        fixedData: {
          function_name: 'DATEDIF & NETWORKDAYS',
          description: '従業員の勤続年数と有給休暇日数を自動計算します',
          syntax: 'DATEDIF(開始日, 終了日, "単位") / NETWORKDAYS(開始日, 終了日)',
          syntax_detail: '- DATEDIF: 2つの日付間の期間を計算（Y=年、M=月、D=日）\n- NETWORKDAYS: 土日を除く営業日数を計算\n- TODAY(): 現在の日付を取得',
          category: 'date',
          spreadsheet_data: [
            [
              { v: "従業員ID", ct: { t: "s" } },
              { v: "従業員名", ct: { t: "s" } },
              { v: "入社日", ct: { t: "s" } },
              { v: "勤続年数", ct: { t: "s" } },
              { v: "有給開始日", ct: { t: "s" } },
              { v: "有給終了日", ct: { t: "s" } },
              { v: "有給営業日数", ct: { t: "s" } },
              { v: "勤続評価", ct: { t: "s" } }
            ],
            [
              { v: "EMP001", ct: { t: "s" } },
              { v: "田中太郎", ct: { t: "s" } },
              { v: "2020-04-01", ct: { t: "s" } },
              { f: "=DATEDIF(C2,TODAY(),\"Y\")" },
              { v: "2024-01-15", ct: { t: "s" } },
              { v: "2024-01-25", ct: { t: "s" } },
              { f: "=NETWORKDAYS(E2,F2)" },
              { f: "=IF(D2>=5,\"ベテラン\",IF(D2>=3,\"経験者\",\"新人\"))" }
            ],
            [
              { v: "EMP002", ct: { t: "s" } },
              { v: "佐藤花子", ct: { t: "s" } },
              { v: "2018-07-15", ct: { t: "s" } },
              { f: "=DATEDIF(C3,TODAY(),\"Y\")" },
              { v: "2024-02-05", ct: { t: "s" } },
              { v: "2024-02-16", ct: { t: "s" } },
              { f: "=NETWORKDAYS(E3,F3)" },
              { f: "=IF(D3>=5,\"ベテラン\",IF(D3>=3,\"経験者\",\"新人\"))" }
            ],
            [
              { v: "EMP003", ct: { t: "s" } },
              { v: "山田次郎", ct: { t: "s" } },
              { v: "2022-10-01", ct: { t: "s" } },
              { f: "=DATEDIF(C4,TODAY(),\"Y\")" },
              { v: "2024-03-10", ct: { t: "s" } },
              { v: "2024-03-20", ct: { t: "s" } },
              { f: "=NETWORKDAYS(E4,F4)" },
              { f: "=IF(D4>=5,\"ベテラン\",IF(D4>=3,\"経験者\",\"新人\"))" }
            ],
            [
              { v: "EMP004", ct: { t: "s" } },
              { v: "鈴木一郎", ct: { t: "s" } },
              { v: "2015-03-01", ct: { t: "s" } },
              { f: "=DATEDIF(C5,TODAY(),\"Y\")" },
              { v: "2024-04-01", ct: { t: "s" } },
              { v: "2024-04-12", ct: { t: "s" } },
              { f: "=NETWORKDAYS(E5,F5)" },
              { f: "=IF(D5>=5,\"ベテラン\",IF(D5>=3,\"経験者\",\"新人\"))" }
            ],
            [
              { v: "EMP005", ct: { t: "s" } },
              { v: "高橋美咲", ct: { t: "s" } },
              { v: "2021-09-15", ct: { t: "s" } },
              { f: "=DATEDIF(C6,TODAY(),\"Y\")" },
              { v: "2024-05-07", ct: { t: "s" } },
              { v: "2024-05-21", ct: { t: "s" } },
              { f: "=NETWORKDAYS(E6,F6)" },
              { f: "=IF(D6>=5,\"ベテラン\",IF(D6>=3,\"経験者\",\"新人\"))" }
            ],
            [
              { v: "集計", ct: { t: "s" } },
              { v: "平均勤続年数", ct: { t: "s" } },
              { f: "=AVERAGE(D2:D6)" },
              { v: "年", ct: { t: "s" } },
              { v: "総有給日数", ct: { t: "s" } },
              { f: "=SUM(G2:G6)" },
              { v: "日", ct: { t: "s" } },
              null
            ],
            [
              { v: "備考", ct: { t: "s" } },
              { v: "勤続3年未満: 新人", ct: { t: "s" } },
              { v: "3-5年: 経験者", ct: { t: "s" } },
              { v: "5年以上: ベテラン", ct: { t: "s" } },
              null,
              null,
              null,
              null
            ]
          ],
          examples: [
            'DATEDIF(C2,TODAY(),"Y") - 入社日から現在までの勤続年数を計算',
            'NETWORKDAYS(E2,F2) - 有給開始日から終了日までの営業日数を計算',
            'IF(D2>=5,"ベテラン",IF(D2>=3,"経験者","新人")) - 勤続年数による評価分類',
            'AVERAGE(D2:D6) - 従業員全体の平均勤続年数を計算'
          ]
        }
      }
    ]
  },
  {
    id: 'inventory',
    name: '在庫管理',
    icon: '📦',
    description: '商品在庫の監視とアラート管理',
    templates: [
      {
        id: 'inventory-alert',
        title: '商品在庫アラート管理',
        description: 'AND・OR関数で複合条件による在庫アラートシステム',
        category: 'inventory',
        icon: '🚨',
        functions: ['AND', 'OR', 'IF', 'CONCATENATE'],
        prompt: '商品の在庫管理でAND関数（在庫数<10かつ売れ筋商品）、OR関数（季節商品または特価商品）の条件でアラート表示するスプレッドシートを作成してください。',
        difficulty: 'intermediate',
        tags: ['在庫管理', '条件判定', 'アラート', '論理関数'],
        fixedData: {
          function_name: 'AND & OR',
          description: '複合条件による在庫アラートシステムを構築します',
          syntax: 'AND(条件1, 条件2, ...) / OR(条件1, 条件2, ...)',
          syntax_detail: '- AND: すべての条件がTRUEの場合にTRUE\n- OR: いずれかの条件がTRUEの場合にTRUE\n- IF: 条件に基づいて異なる値を返す\n- 複合条件で高度な判定が可能',
          category: 'logic',
          spreadsheet_data: [
            [
              { v: "商品名", ct: { t: "s" } },
              { v: "在庫数", ct: { t: "s" } },
              { v: "単価", ct: { t: "s" } },
              { v: "商品区分", ct: { t: "s" } },
              { v: "売れ筋", ct: { t: "s" } },
              { v: "在庫アラート", ct: { t: "s" } },
              { v: "特別アラート", ct: { t: "s" } },
              { v: "総合判定", ct: { t: "s" } }
            ],
            [
              { v: "商品A", ct: { t: "s" } },
              { v: 8, ct: { t: "n" } },
              { v: 1500, ct: { t: "n" } },
              { v: "季節商品", ct: { t: "s" } },
              { v: "売れ筋", ct: { t: "s" } },
              { f: "=IF(AND(B2<10,E2=\"売れ筋\"),\"要発注\",\"充分\")" },
              { f: "=IF(OR(D2=\"季節商品\",D2=\"特価商品\"),\"要注意\",\"通常\")" },
              { f: "=IF(AND(B2<10,OR(D2=\"季節商品\",E2=\"売れ筋\")),\"緊急発注\",IF(B2<5,\"発注検討\",\"在庫OK\"))" }
            ],
            [
              { v: "商品B", ct: { t: "s" } },
              { v: 15, ct: { t: "n" } },
              { v: 2000, ct: { t: "n" } },
              { v: "通常商品", ct: { t: "s" } },
              { v: "通常", ct: { t: "s" } },
              { f: "=IF(AND(B3<10,E3=\"売れ筋\"),\"要発注\",\"充分\")" },
              { f: "=IF(OR(D3=\"季節商品\",D3=\"特価商品\"),\"要注意\",\"通常\")" },
              { f: "=IF(AND(B3<10,OR(D3=\"季節商品\",E3=\"売れ筋\")),\"緊急発注\",IF(B3<5,\"発注検討\",\"在庫OK\"))" }
            ],
            [
              { v: "商品C", ct: { t: "s" } },
              { v: 5, ct: { t: "n" } },
              { v: 3000, ct: { t: "n" } },
              { v: "特価商品", ct: { t: "s" } },
              { v: "売れ筋", ct: { t: "s" } },
              { f: "=IF(AND(B4<10,E4=\"売れ筋\"),\"要発注\",\"充分\")" },
              { f: "=IF(OR(D4=\"季節商品\",D4=\"特価商品\"),\"要注意\",\"通常\")" },
              { f: "=IF(AND(B4<10,OR(D4=\"季節商品\",E4=\"売れ筋\")),\"緊急発注\",IF(B4<5,\"発注検討\",\"在庫OK\"))" }
            ],
            [
              { v: "商品D", ct: { t: "s" } },
              { v: 20, ct: { t: "n" } },
              { v: 2500, ct: { t: "n" } },
              { v: "通常商品", ct: { t: "s" } },
              { v: "通常", ct: { t: "s" } },
              { f: "=IF(AND(B5<10,E5=\"売れ筋\"),\"要発注\",\"充分\")" },
              { f: "=IF(OR(D5=\"季節商品\",D5=\"特価商品\"),\"要注意\",\"通常\")" },
              { f: "=IF(AND(B5<10,OR(D5=\"季節商品\",E5=\"売れ筋\")),\"緊急発注\",IF(B5<5,\"発注検討\",\"在庫OK\"))" }
            ],
            [
              { v: "商品E", ct: { t: "s" } },
              { v: 3, ct: { t: "n" } },
              { v: 1800, ct: { t: "n" } },
              { v: "季節商品", ct: { t: "s" } },
              { v: "通常", ct: { t: "s" } },
              { f: "=IF(AND(B6<10,E6=\"売れ筋\"),\"要発注\",\"充分\")" },
              { f: "=IF(OR(D6=\"季節商品\",D6=\"特価商品\"),\"要注意\",\"通常\")" },
              { f: "=IF(AND(B6<10,OR(D6=\"季節商品\",E6=\"売れ筋\")),\"緊急発注\",IF(B6<5,\"発注検討\",\"在庫OK\"))" }
            ],
            [
              { v: "集計", ct: { t: "s" } },
              { v: "総在庫数", ct: { t: "s" } },
              { f: "=SUM(B2:B6)" },
              { v: "平均単価", ct: { t: "s" } },
              { f: "=AVERAGE(C2:C6)" },
              { v: "要発注件数", ct: { t: "s" } },
              { f: "=COUNTIF(F2:F6,\"要発注\")" },
              { v: "緊急発注件数", ct: { t: "s" } },
              { f: "=COUNTIF(H2:H6,\"緊急発注\")" }
            ],
            [
              { v: "判定条件", ct: { t: "s" } },
              { v: "在庫<10且つ売れ筋=要発注", ct: { t: "s" } },
              { v: "季節商品OR特価商品=要注意", ct: { t: "s" } },
              { v: "緊急発注=在庫<10且つ(季節OR売れ筋)", ct: { t: "s" } },
              null,
              null,
              null,
              null
            ]
          ],
          examples: [
            'AND(B2<10,E2="売れ筋") - 在庫が10未満かつ売れ筋商品の場合TRUE',
            'OR(D2="季節商品",D2="特価商品") - 季節商品または特価商品の場合TRUE',
            'IF(AND(条件1,条件2),"結果1","結果2") - 複数条件がすべて満たされた場合の処理',
            'COUNTIF(F2:F6,"要発注") - 特定条件に一致するセルの数をカウント'
          ]
        }
      }
    ]
  },
  {
    id: 'finance',
    name: '経費計算',
    icon: '💰',
    description: '経費データの集計と分析',
    templates: [
      {
        id: 'expense-summary',
        title: '経費集計',
        description: 'SUMIF関数で部門別集計、IF関数で承認判定',
        category: 'finance',
        icon: '📊',
        functions: ['SUMIF', 'IF', 'COUNTIF'],
        prompt: '部門別の経費データを作成してください。SUMIF関数で部門ごとの経費合計を計算し、1万円以上の経費には「要承認」、未満には「承認済み」とIF関数で表示するスプレッドシートを作成してください。',
        difficulty: 'intermediate',
        tags: ['経費', '部門', '承認']
      }
    ]
  },
  {
    id: 'education',
    name: '成績管理',
    icon: '📈',
    description: '学生の成績データ管理',
    templates: [
      {
        id: 'grade-analysis',
        title: '成績分析',
        description: 'IF関数で合否判定、統計関数で分析',
        category: 'education',
        icon: '🎓',
        functions: ['IF', 'AVERAGE', 'MAX', 'MIN'],
        prompt: '学生の成績管理表を作成してください。各科目の点数から平均点をAVERAGE関数で計算し、60点以上は「合格」、未満は「不合格」とIF関数で判定し、MAX・MIN関数で最高点・最低点も表示するスプレッドシートを作成してください。',
        difficulty: 'beginner',
        tags: ['成績', '学生', '統計']
      }
    ]
  },
  {
    id: 'lookup',
    name: '検索関数',
    icon: '🔍',
    description: 'VLOOKUP・INDEX・MATCH等の検索関数',
    templates: [
      {
        id: 'vlookup-basic',
        title: 'VLOOKUP基本',
        description: 'VLOOKUP関数による商品マスター検索',
        category: 'lookup',
        icon: '🔎',
        functions: ['VLOOKUP'],
        prompt: '商品コードから商品名と価格を検索するVLOOKUP関数を使用したスプレッドシートを作成してください。商品マスターテーブルと検索テーブルを含む8行8列の構造で作成してください。',
        difficulty: 'intermediate',
        tags: ['VLOOKUP', '検索', '商品マスター']
      },
      {
        id: 'index-match',
        title: 'INDEX・MATCH組み合わせ',
        description: 'INDEX関数とMATCH関数の組み合わせ',
        category: 'lookup',
        icon: '🎯',
        functions: ['INDEX', 'MATCH'],
        prompt: 'INDEX関数とMATCH関数を組み合わせて、社員IDから部署名を検索するスプレッドシートを作成してください。VLOOKUPより柔軟な検索機能を示してください。',
        difficulty: 'advanced',
        tags: ['INDEX', 'MATCH', '社員管理'],
        fixedData: {
          function_name: 'INDEX & MATCH',
          description: 'INDEX関数とMATCH関数を組み合わせてVLOOKUPより柔軟な検索を実現します',
          syntax: 'INDEX(配列, MATCH(検索値, 検索配列, 0))',
          syntax_detail: 'INDEX(array, row_num) - 配列から指定した位置の値を取得 + MATCH(lookup_value, lookup_array, [match_type]) - 検索値の位置を返す',
          category: '検索関数',
          spreadsheet_data: [
            [
              { v: "社員ID", ct: { t: "s" }, bg: "#E3F2FD" },
              { v: "社員名", ct: { t: "s" }, bg: "#E3F2FD" },
              { v: "部署名", ct: { t: "s" }, bg: "#E3F2FD" },
              { v: "検索ID", ct: { t: "s" }, bg: "#FFF8E1" },
              { v: "検索結果", ct: { t: "s" }, bg: "#FFF8E1" },
              null, null, null
            ],
            [
              { v: 101, ct: { t: "n" } },
              { v: "田中", ct: { t: "s" } },
              { v: "営業部", ct: { t: "s" } },
              { v: 102, ct: { t: "n" } },
              { v: null, f: "=INDEX(C2:C6,MATCH(D2,A2:A6,0))", bg: "#FFE0B2", fc: "#D84315" },
              null, null, null
            ],
            [
              { v: 102, ct: { t: "n" } },
              { v: "佐藤", ct: { t: "s" } },
              { v: "開発部", ct: { t: "s" } },
              { v: 103, ct: { t: "n" } },
              { v: null, f: "=INDEX(C2:C6,MATCH(D3,A2:A6,0))", bg: "#FFE0B2", fc: "#D84315" },
              null, null, null
            ],
            [
              { v: 103, ct: { t: "n" } },
              { v: "鈴木", ct: { t: "s" } },
              { v: "人事部", ct: { t: "s" } },
              { v: 104, ct: { t: "n" } },
              { v: null, f: "=INDEX(C2:C6,MATCH(D4,A2:A6,0))", bg: "#FFE0B2", fc: "#D84315" },
              null, null, null
            ],
            [
              { v: 104, ct: { t: "n" } },
              { v: "山田", ct: { t: "s" } },
              { v: "総務部", ct: { t: "s" } },
              { v: 105, ct: { t: "n" } },
              { v: null, f: "=INDEX(C2:C6,MATCH(D5,A2:A6,0))", bg: "#FFE0B2", fc: "#D84315" },
              null, null, null
            ],
            [
              { v: 105, ct: { t: "n" } },
              { v: "伊藤", ct: { t: "s" } },
              { v: "マーケティング部", ct: { t: "s" } },
              { v: 101, ct: { t: "n" } },
              { v: null, f: "=INDEX(C2:C6,MATCH(D6,A2:A6,0))", bg: "#FFE0B2", fc: "#D84315" },
              null, null, null
            ],
            [null, null, null, null, null, null, null, null],
            [null, null, null, null, null, null, null, null]
          ],
          examples: ['=INDEX(C2:C6,MATCH(D2,A2:A6,0))', '=MATCH(102,A2:A6,0)', '=INDEX(B2:B6,MATCH(103,A2:A6,0))']
        }
      }
    ]
  },
  {
    id: 'math-stats',
    name: '数学・統計',
    icon: '📊',
    description: 'SUM・AVERAGE・MAX・MIN等の数学統計関数',
    templates: [
      {
        id: 'statistics-basic',
        title: '基本統計',
        description: 'AVERAGE・MAX・MIN・COUNTによる統計分析',
        category: 'math-stats',
        icon: '📈',
        functions: ['AVERAGE', 'MAX', 'MIN', 'COUNT'],
        prompt: '売上データからAVERAGE関数で平均売上、MAX関数で最高売上、MIN関数で最低売上、COUNT関数でデータ数を計算するスプレッドシートを作成してください。',
        difficulty: 'beginner',
        tags: ['統計', '平均', '最大', '最小']
      },
      {
        id: 'conditional-stats',
        title: '条件付き統計',
        description: 'SUMIF・COUNTIF・AVERAGEIFによる条件集計',
        category: 'math-stats',
        icon: '🎲',
        functions: ['SUMIF', 'COUNTIF', 'AVERAGEIF'],
        prompt: '部門別売上データでSUMIF関数による部門別合計、COUNTIF関数による部門別人数、AVERAGEIF関数による部門別平均を計算するスプレッドシートを作成してください。',
        difficulty: 'intermediate',
        tags: ['条件集計', 'SUMIF', 'COUNTIF']
      }
    ]
  },
  {
    id: 'date-time',
    name: '日付・時刻',
    icon: '📅',
    description: 'TODAY・DATE・YEAR等の日付時刻関数',
    templates: [
      {
        id: 'date-functions',
        title: '日付関数基本',
        description: 'TODAY・YEAR・MONTH・DAYによる日付処理',
        category: 'date-time',
        icon: '📆',
        functions: ['TODAY', 'YEAR', 'MONTH', 'DAY'],
        prompt: 'TODAY関数で今日の日付を取得し、YEAR・MONTH・DAY関数で年月日を個別に抽出し、社員の年齢計算を行うスプレッドシートを作成してください。',
        difficulty: 'beginner',
        tags: ['日付', '年齢計算', 'TODAY'],
        fixedData: {
          function_name: 'TODAY & YEAR & MONTH & DAY',
          description: 'TODAY関数で現在の日付を取得し、YEAR・MONTH・DAY関数で日付要素を抽出します',
          syntax: 'TODAY() + YEAR(日付) + MONTH(日付) + DAY(日付)',
          syntax_detail: 'TODAY() - 現在の日付を取得 + YEAR(date) - 日付から年を抽出 + MONTH(date) - 日付から月を抽出 + DAY(date) - 日付から日を抽出',
          category: '日付関数',
          spreadsheet_data: [
            [
              { v: "社員名", ct: { t: "s" }, bg: "#E3F2FD" },
              { v: "生年月日", ct: { t: "s" }, bg: "#E3F2FD" },
              { v: "年齢", ct: { t: "s" }, bg: "#E3F2FD" },
              { v: "今日の日付", ct: { t: "s" }, bg: "#E3F2FD" },
              { v: "年", ct: { t: "s" }, bg: "#E3F2FD" },
              { v: "月", ct: { t: "s" }, bg: "#E3F2FD" },
              { v: "日", ct: { t: "s" }, bg: "#E3F2FD" },
              null
            ],
            [
              { v: "田中", ct: { t: "s" } },
              { v: "1990/05/15", ct: { t: "s" } },
              { v: null, f: "=(YEAR(TODAY())-1990)", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=TODAY()", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=YEAR(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=MONTH(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=DAY(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              null
            ],
            [
              { v: "佐藤", ct: { t: "s" } },
              { v: "1985/11/20", ct: { t: "s" } },
              { v: null, f: "=(YEAR(TODAY())-1985)", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=TODAY()", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=YEAR(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=MONTH(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=DAY(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              null
            ],
            [
              { v: "鈴木", ct: { t: "s" } },
              { v: "1992/02/10", ct: { t: "s" } },
              { v: null, f: "=(YEAR(TODAY())-1992)", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=TODAY()", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=YEAR(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=MONTH(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=DAY(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              null
            ],
            [
              { v: "山田", ct: { t: "s" } },
              { v: "1995/08/30", ct: { t: "s" } },
              { v: null, f: "=(YEAR(TODAY())-1995)", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=TODAY()", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=YEAR(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=MONTH(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              { v: null, f: "=DAY(TODAY())", bg: "#FFE0B2", fc: "#D84315" },
              null
            ],
            [null, null, null, null, null, null, null, null],
            [null, null, null, null, null, null, null, null],
            [null, null, null, null, null, null, null, null]
          ],
          examples: ['=TODAY()', '=YEAR(TODAY())', '=MONTH(TODAY())', '=DAY(TODAY())', '=(YEAR(TODAY())-1990)']
        }
      },
      {
        id: 'date-calculations',
        title: '日付計算',
        description: 'DATEDIF・WORKDAY等による期間計算',
        category: 'date-time',
        icon: '⏰',
        functions: ['DATEDIF', 'WORKDAY', 'NETWORKDAYS'],
        prompt: 'プロジェクト管理表を作成してください。プロジェクト名、開始日、終了日、期間（日）、稼働日数、評価の列を含めてください。DATEDIF関数で期間（日数）を計算し、NETWORKDAYS関数で稼働日数（営業日数）を計算し、IF関数で期間に基づく評価（15日以上なら「長期」、未満なら「短期」）を表示してください。',
        difficulty: 'intermediate',
        tags: ['プロジェクト管理', '期間計算', '評価']
      }
    ]
  },
  {
    id: 'text-functions',
    name: '文字列関数',
    icon: '📝',
    description: 'CONCATENATE・LEFT・RIGHT等の文字列処理',
    templates: [
      {
        id: 'text-basic',
        title: '文字列基本',
        description: 'CONCATENATE・LEFT・RIGHT・MIDによる文字列操作',
        category: 'text-functions',
        icon: '✂️',
        functions: ['CONCATENATE', 'LEFT', 'RIGHT', 'MID'],
        prompt: '社員の姓名をCONCATENATE関数で結合し、LEFT関数で姓の最初の文字、RIGHT関数で名の最後の文字を抽出する文字列処理のスプレッドシートを作成してください。',
        difficulty: 'beginner',
        tags: ['文字列', '結合', '抽出']
      },
      {
        id: 'text-advanced',
        title: '高度な文字列処理',
        description: 'TRIM・UPPER・LOWER・LEN等の文字列関数',
        category: 'text-functions',
        icon: '🔤',
        functions: ['TRIM', 'UPPER', 'LOWER', 'LEN'],
        prompt: '顧客データの整理でTRIM関数で余分なスペース除去、UPPER/LOWER関数で大文字小文字変換、LEN関数で文字数カウントを行うデータクリーニング表を作成してください。',
        difficulty: 'intermediate',
        tags: ['データクリーニング', 'TRIM', '文字変換']
      }
    ]
  },
  {
    id: 'logical',
    name: '論理関数',
    icon: '⚡',
    description: 'IF・AND・OR・NOT等の論理処理',
    templates: [
      {
        id: 'logical-basic',
        title: 'IF関数基本',
        description: 'IF関数による条件分岐と判定',
        category: 'logical',
        icon: '🔀',
        functions: ['IF'],
        prompt: '学生の点数からIF関数で合否判定（60点以上で合格）を行い、さらにネストしたIF関数で優良可の評価を付けるスプレッドシートを作成してください。',
        difficulty: 'beginner',
        tags: ['条件分岐', '判定', 'IF関数']
      },
      {
        id: 'logical-advanced',
        title: '複合論理判定',
        description: 'AND・OR・NOT関数との組み合わせ',
        category: 'logical',
        icon: '🧠',
        functions: ['IF', 'AND', 'OR', 'NOT'],
        prompt: '商品の在庫管理でAND関数（在庫数<10かつ売れ筋商品）、OR関数（季節商品または特価商品）の条件でアラート表示するスプレッドシートを作成してください。',
        difficulty: 'advanced',
        tags: ['複合条件', 'AND', 'OR', '在庫管理']
      }
    ]
  },
  {
    id: 'advanced',
    name: '高度な関数',
    icon: '🚀',
    description: '複数関数の組み合わせや高度な機能',
    templates: [
      {
        id: 'pivot-simulation',
        title: 'ピボット風集計',
        description: 'SUMIFS・COUNTIFS等による多条件集計',
        category: 'advanced',
        icon: '📋',
        functions: ['SUMIFS', 'COUNTIFS', 'AVERAGEIFS'],
        prompt: '売上データから地域×商品カテゴリーのマトリックス集計をSUMIFS関数で作成し、COUNTIFS関数で件数、AVERAGEIFS関数で平均単価を計算するピボット風の集計表を作成してください。',
        difficulty: 'advanced',
        tags: ['多条件集計', 'SUMIFS', 'マトリックス']
      },
      {
        id: 'dashboard-kpi',
        title: 'KPIダッシュボード',
        description: '複数関数を組み合わせたKPI計算',
        category: 'advanced',
        icon: '📊',
        functions: ['IF', 'VLOOKUP', 'SUM', 'AVERAGE', 'COUNTIF'],
        prompt: '売上目標達成率、顧客満足度、商品回転率などのKPIを複数の関数（IF、VLOOKUP、SUM、AVERAGE、COUNTIF）を組み合わせて計算するダッシュボードを作成してください。',
        difficulty: 'advanced',
        tags: ['KPI', 'ダッシュボード', '複合関数']
      }
    ]
  }
];