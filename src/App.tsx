import ChatGPTSpreadsheet from './components/ChatGPTSpreadsheet';
import './App.css';

function App() {
  return (
    <div className="text-center p-5 max-w-6xl mx-auto">
      <h1 className="text-3xl font-bold mb-6">ChatGPT連携 Excel関数学習デモ</h1>
      <ChatGPTSpreadsheet />
    </div>
  );
}

export default App
