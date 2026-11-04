import './heldInRoom'
import { createRoot } from 'react-dom/client'
import { App } from './App'
import { WhiteNeutral } from './whiteNeutral'
import './compact-in-room.css'

createRoot(document.getElementById('root')!).render(
  <>
    <WhiteNeutral />
    <App />
  </>,
)
