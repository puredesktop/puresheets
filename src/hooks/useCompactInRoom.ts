import { useAdjacentApp } from '@purescience/platform-ui/bridge/react/useAdjacentApp'

/**
 * Sheets is working beside the Executive team room in split view, so it shows
 * only the grid, formula bar and sheet tabs. 'Executive' is that app's
 * navigationLabel. Off in every other case, including beside any other app.
 */
export function useCompactInRoom(): boolean {
  return useAdjacentApp()?.heldInRoom === true
}
