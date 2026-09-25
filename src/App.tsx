import { useRef } from 'react'
import { AppFrame } from '@purescience/platform-bridge/components/AppFrame'
import { EmptyState } from '@purescience/platform-ui/components/common/feedback/EmptyState'
import { usePlatformBridge } from '@purescience/platform-ui/bridge/react/usePlatformBridge'
import { usePlatformViewportResource } from '@purescience/platform-ui/bridge/react/usePlatformViewportResource'
import { isStandaloneDevMode } from './bridge/platformBridge'
import { PureSheetsShell } from './components/PureSheetsShell'
import { usePureSheetsBoot } from './hooks/usePureSheetsBoot'

export function App(): React.ReactElement {
  const { error: bridgeError, ready, meta } = usePlatformBridge()
  const standalone = isStandaloneDevMode()
  const appReady = ready || standalone
  const { boot, bootError } = usePureSheetsBoot(appReady)
  const { resource, clearResource } = usePlatformViewportResource(
    ready && !standalone,
    meta,
  )
  // Once the platform has ever handed us an explicit resource, stop offering the
  // remembered boot file: after such a resource is applied, clearResource() nulls
  // `resource`, and re-offering the (now stale) boot file would reload it over
  // the file the user just opened.
  const platformResourceSeenRef = useRef(false)
  if (resource) platformResourceSeenRef.current = true

  if (bridgeError && !standalone) {
    return (
      <AppFrame>
        <EmptyState
          tone="error"
          title="Bridge unavailable"
          message={bridgeError.message}
        />
      </AppFrame>
    )
  }

  if (bootError) {
    return (
      <AppFrame>
        <EmptyState
          tone="error"
          title="PureSheets boot failed"
          message={bootError.message}
        />
      </AppFrame>
    )
  }

  if (!appReady || !boot) {
    return (
      <AppFrame>
        <EmptyState
          tone="neutral"
          title="PureSheets"
          message="Loading spreadsheet workspace..."
        />
      </AppFrame>
    )
  }

  // Self-reopen the last file through the same resource loader the shell uses
  // for explicit opens (it reads + recovers content before setting the save
  // target). We deliberately do NOT pass lastFilePath as `initialFilePath`:
  // that would make the blank boot document's save target the real file, so the
  // first autosave would overwrite it with an empty sheet. An explicit platform
  // resource always wins over the remembered file.
  const bootFileResource =
    platformResourceSeenRef.current || !boot.settings.lastFilePath
      ? null
      : { path: boot.settings.lastFilePath }
  const effectiveResource = resource ?? bootFileResource

  return (
    <PureSheetsShell
      initialDocument={boot.document}
      resource={effectiveResource}
      onResourceHandled={clearResource}
    />
  )
}
