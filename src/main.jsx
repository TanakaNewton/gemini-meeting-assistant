// src/main.jsx (変更なし、確認用)
import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import './index.css'
import { CssBaseline } from '@mui/material'; // Material UIのベースラインCSSを適用

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <CssBaseline /> {/* Material UIのスタイルリセットを追加 */}
    <App />
  </React.StrictMode>,
)