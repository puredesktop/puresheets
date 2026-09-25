import { useEffect, useState } from 'react'
import { fetchPureSheetsSettings } from '../bridge/platformBridge'
import { createBlankWorkbook } from '../lib/workbookModel'
import type { PureSheetsDocument, PureSheetsSettings } from '../types'

export interface PureSheetsBoot {
  document: PureSheetsDocument
  settings: PureSheetsSettings
}

export function usePureSheetsBoot(ready: boolean): {
  boot: PureSheetsBoot | null
  bootError: Error | null
} {
  const [boot, setBoot] = useState<PureSheetsBoot | null>(null)
  const [bootError, setBootError] = useState<Error | null>(null)

  useEffect(() => {
    if (!ready) return
    let cancelled = false
    async function load(): Promise<void> {
      try {
        const settings = await fetchPureSheetsSettings()
        if (cancelled) return
        setBoot({
          document: createBlankWorkbook(),
          settings,
        })
        setBootError(null)
      } catch (error) {
        if (cancelled) return
        setBootError(error instanceof Error ? error : new Error(String(error)))
      }
    }
    void load()
    return () => {
      cancelled = true
    }
  }, [ready])

  return { boot, bootError }
}
