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
  // 📊 数学・三角関数 - オレンジ系
  math: { bg: "#FFF3E0", fc: "#E65100" },         // 数学・三角関数
  
  // 📈 統計関数 - 赤系
  statistical: { bg: "#FFEBEE", fc: "#C62828" },  // 統計関数
  
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
  
  // 🔧 情報関数 - 水色系
  information: { bg: "#E0F2F1", fc: "#00695C" },   // 情報関数
  
  // 🗄️ データベース関数 - 茶色系
  database: { bg: "#EFEBE9", fc: "#5D4037" },     // データベース関数
  
  // ⚙️ エンジニアリング関数 - 深緑系
  engineering: { bg: "#E8F5E8", fc: "#388E3C" },  // エンジニアリング関数
  
  // 📊 キューブ関数 - インディゴ系
  cube: { bg: "#E8EAF6", fc: "#3F51B5" },         // キューブ関数
  
  // 🌐 Web関数 - シアン系
  web: { bg: "#E0F7FA", fc: "#0097A7" },          // Web関数
  
  // 📋 Google Sheets専用 - ライム系
  googlesheets: { bg: "#F1F8E9", fc: "#689F38" }, // Google Sheets専用
  
  // 🔧 その他 - グレー系
  other: { bg: "#F5F5F5", fc: "#616161" }         // その他
};