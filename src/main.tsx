import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.tsx'
import './index.css'

const root = document.getElementById('root')!

ReactDOM.createRoot(root).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)

// Mark root as ready after splash animation completes
// Total step delays: 100+300+580+820+650+480 ≈ 2930ms → add buffer = 3200ms
setTimeout(() => {
  root.classList.add('ready')
}, 3200)


// Use contextBridge
window.ipcRenderer.on('main-process-message', (_event, message) => {
  console.log(message)
})
