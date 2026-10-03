/** One workbook-owned boundary for open/new. Never install content before its old save settles. */
export function createWorkbookTransitions(
  flush: () => Promise<void>,
  onBusy: (busy: boolean) => void,
) {
  let generation = 0
  let busy = false
  return {
    get version() {
      return generation
    },
    get busy() {
      return busy
    },
    dispose() {
      generation++
      busy = false
    },
    async run<T>(
      load: () => Promise<T>,
      install: (value: T) => void,
    ): Promise<boolean> {
      const attempt = ++generation
      busy = true
      onBusy(true)
      try {
        await flush()
        if (attempt !== generation) return false
        const value = await load()
        if (attempt !== generation) return false
        install(value)
        return true
      } catch (error) {
        if (attempt !== generation) return false
        throw error
      } finally {
        if (attempt === generation) {
          busy = false
          onBusy(false)
        }
      }
    },
  }
}
