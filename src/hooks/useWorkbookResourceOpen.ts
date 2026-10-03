import { useEffect, useRef, useState } from 'react'

/** Consume only the successful matching host request; later requests may supersede a slow open. */
export function useWorkbookResourceOpen(
  ready: boolean,
  path: string | undefined,
  open: (path: string) => Promise<boolean>,
  clear: () => void,
) {
  const latest = useRef({ path, open, clear })
  latest.current = { path, open, clear }
  const attempted = useRef<string | undefined>(undefined)
  const generation = useRef(0)
  const [failedPath, setFailedPath] = useState<string>()
  const [settlement, setSettlement] = useState(0)
  const mounted = useRef(false)
  useEffect(() => {
    mounted.current = true
    return () => {
      mounted.current = false
      attempted.current = undefined
      generation.current++
    }
  }, [])
  useEffect(() => {
    if (!path) attempted.current = undefined
    if (!ready || !path || attempted.current === path) return
    attempted.current = path
    const attempt = ++generation.current
    // A separate lifetime effect handles disposal; busy changes during our own
    // open must not cancel the request or consume a newer host resource.
    void latest.current
      .open(path)
      .then(
        opened => {
          if (
            mounted.current &&
            generation.current === attempt &&
            latest.current.path === path
          ) {
            setFailedPath(opened ? undefined : path)
            if (opened) latest.current.clear()
          }
        },
        () => {
          if (
            mounted.current &&
            generation.current === attempt &&
            latest.current.path === path
          )
            setFailedPath(path)
        },
      )
      .finally(() => {
        if (mounted.current) setSettlement(value => value + 1)
      })
  }, [ready, path, settlement])
  return {
    failed: failedPath === path && !!path,
    retry: () => {
      attempted.current = undefined
      setFailedPath(undefined)
      setSettlement(value => value + 1)
    },
  }
}
