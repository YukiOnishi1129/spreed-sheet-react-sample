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
        prompt: 'プロジェクト管理表を作成してください。プロジェクト名、開始日、終了日、期間（日）、稼働日数、評価の列を含めてください。期間と稼働日数は計算済みの数値で設定し、IF関数で期間に基づく評価（15日以上なら「長期」、未満なら「短期」）を表示してください。DATEDIF関数やNETWORKDAYS関数は使用せず、数値とIF関数、SUM関数のみを使用してください。',
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