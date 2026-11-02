// @vitest-environment happy-dom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import { useCompactInRoom } from './useCompactInRoom'

const adjacent = vi.hoisted(() => ({ current: null as { tabId: string; name: string } | null }))
vi.mock('@purescience/platform-ui/bridge/react/useAdjacentApp', () => ({
  useAdjacentApp: () => adjacent.current,
}))

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

function compactBeside(name: string | null): boolean {
  adjacent.current = name === null ? null : { tabId: 'tab-1', name, ...(name === 'Executive' ? { heldInRoom: true } : {}) }
  let result: boolean | undefined
  function Probe(): null {
    result = useCompactInRoom()
    return null
  }
  const root = createRoot(document.createElement('div'))
  act(() => root.render(<Probe />))
  act(() => root.unmount())
  return result as boolean
}

describe('useCompactInRoom', () => {
  it('is on beside the Executive app', () => {
    expect(compactBeside('Executive')).toBe(true)
  })

  it('is off with no neighbour', () => {
    expect(compactBeside(null)).toBe(false)
  })

  it.each(['Mail', 'Sheets', 'Executive team', 'executive', ''])(
    'is off beside %j',
    name => {
      expect(compactBeside(name)).toBe(false)
    },
  )
})
