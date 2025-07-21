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
        tags: ['売上', '目標管理', 'KPI']
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
  }
];