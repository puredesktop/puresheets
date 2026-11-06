import './heldInRoom'
import { createRoot } from 'react-dom/client'
import { App } from './App'
import { SheetsStyles } from './ui/styles'
import './compact-in-room.css'

createRoot(document.getElementById('root')!).render(
  <>
    <SheetsStyles />
    <App />
  </>,
)
