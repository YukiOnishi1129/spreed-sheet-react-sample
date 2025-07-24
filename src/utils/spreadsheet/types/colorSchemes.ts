// カテゴリ別の色定義
// UIの色分けに合わせて調整:
// 💰 財務関数 - 金色/黄色系
// 📊 数学・統計関数 - オレンジ系
// 🔍 検索・参照関数 - 青系  
// ⚡ 論理関数 - 緑系
// 📅 日付・時刻関数 - 紫系
// 📝 文字列関数 - ピンク系
// 🔧 その他/未対応 - グレー系

export const COLOR_SCHEMES = {
  // 📊 数学・統計関数 - オレンジ系
  math: { bg: "#FFF3E0", fc: "#E65100" },         // 数学・三角関数
  statistical: { bg: "#FFF3E0", fc: "#E65100" },  // 統計関数
  
  // 📝 文字列関数 - ピンク系
  text: { bg: "#FCE4EC", fc: "#C2185B" },         
  
  // 📅 日付・時刻関数 - 紫系
  datetime: { bg: "#F3E5F5", fc: "#7B1FA2" },     
  
  // ⚡ 論理関数 - 緑系
  logical: { bg: "#E8F5E8", fc: "#2E7D32" },      
  
  // 🔍 検索・参照関数 - 青系
  lookup: { bg: "#E3F2FD", fc: "#1976D2" },       
  
  // 💰 財務関数 - 金色/黄色系
  financial: { bg: "#FFFDE7", fc: "#F57F17" },    
  
  // 🔧 その他/未対応 - グレー系
  information: { bg: "#F5F5F5", fc: "#616161" },   // 情報関数
  database: { bg: "#F5F5F5", fc: "#616161" },     // データベース関数
  engineering: { bg: "#F5F5F5", fc: "#616161" },  // エンジニアリング関数
  cube: { bg: "#F5F5F5", fc: "#616161" },         // キューブ関数
  web: { bg: "#F5F5F5", fc: "#616161" },          // Web関数
  googlesheets: { bg: "#F5F5F5", fc: "#616161" }, // Google Sheets専用
  other: { bg: "#F5F5F5", fc: "#616161" }         // その他
};