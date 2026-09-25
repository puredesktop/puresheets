import { bridge } from '@purescience/platform-ui/bridge/client'
import { PLATFORM_BRIDGE_METHODS } from '@purescience/platform-ui/bridge/methods'
import { PURESHEETS_APP_SLUG } from '../constants'
import type { PureSheetsSettings } from '../types'

export { bridge }

export function isStandaloneDevMode(): boolean {
  return import.meta.env.DEV && window.parent === window
}

export async function fetchPureSheetsSettings(): Promise<PureSheetsSettings> {
  if (isStandaloneDevMode()) {
    const raw = window.localStorage.getItem('purescience:sheets:settings')
    return raw ? JSON.parse(raw) as PureSheetsSettings : {}
  }
  return bridge.call<PureSheetsSettings>(PLATFORM_BRIDGE_METHODS.SETTINGS_APP_GET, [
    PURESHEETS_APP_SLUG,
  ])
}

export async function updatePureSheetsSettings(
  patch: Partial<PureSheetsSettings>,
): Promise<void> {
  if (isStandaloneDevMode()) {
    const current = await fetchPureSheetsSettings()
    window.localStorage.setItem(
      'purescience:sheets:settings',
      JSON.stringify({ ...current, ...patch }),
    )
    return
  }
  await bridge.call(PLATFORM_BRIDGE_METHODS.SETTINGS_APP_UPDATE, [
    { appSlug: PURESHEETS_APP_SLUG, patch },
  ])
}

export async function readClipboard(): Promise<string> {
  if (isStandaloneDevMode()) {
    return navigator.clipboard?.readText() ?? ''
  }
  const result = await bridge.call<{ text: string }>(PLATFORM_BRIDGE_METHODS.CLIPBOARD_READ, [])
  return result.text ?? ''
}

export async function writeClipboard(text: string): Promise<void> {
  if (isStandaloneDevMode()) {
    await navigator.clipboard?.writeText(text)
    return
  }
  await bridge.call(PLATFORM_BRIDGE_METHODS.CLIPBOARD_WRITE, [text])
}

export async function readTextFile(path: string): Promise<string> {
  if (isStandaloneDevMode()) {
    const raw = window.localStorage.getItem(`purescience:sheets:file:${path}`)
    if (raw === null) throw new Error(`No standalone dev file saved for ${path}`)
    return raw
  }
  return bridge.call<string>(PLATFORM_BRIDGE_METHODS.FS_READ, [path])
}

/** Generic open-file dialog for imports (CSV, XLSX). */
export async function openImportFile(): Promise<string | null> {
  if (isStandaloneDevMode()) return null
  const result = await bridge.call<{ path?: string | null }>(
    PLATFORM_BRIDGE_METHODS.DIALOG_OPEN_FILE,
  )
  return typeof result.path === 'string' && result.path.trim() ? result.path : null
}

/** Write binary content (base64) to a path — used by XLSX export. */
export async function writeBinaryFile(
  path: string,
  base64: string,
): Promise<void> {
  if (isStandaloneDevMode()) {
    throw new Error('Standalone dev cannot write local binary files.')
  }
  await bridge.call(PLATFORM_BRIDGE_METHODS.FS_WRITE_BINARY, [
    { path, base64 },
  ])
}

export async function chooseSaveFilePath(opts?: {
  defaultName?: string
  filters?: { name: string; extensions: string[] }[]
}): Promise<string | null> {
  if (isStandaloneDevMode()) return null
  const result = await bridge.call<{ path?: string | null }>(
    PLATFORM_BRIDGE_METHODS.DIALOG_SAVE_FILE,
    [opts],
  )
  return typeof result.path === 'string' && result.path.trim() ? result.path : null
}

export async function openPath(path: string): Promise<void> {
  if (isStandaloneDevMode()) return
  await bridge.call(PLATFORM_BRIDGE_METHODS.OS_OPEN_PATH, [path])
}

export async function readBinaryFile(
  path: string,
  maxBytes = 4_000_000,
): Promise<{
  base64: string
  mimeType: string
  byteLength: number
  truncated: boolean
}> {
  if (isStandaloneDevMode()) {
    throw new Error('Standalone dev cannot read local binary files.')
  }
  const result = await bridge.call<{
    base64?: string
    mimeType?: string
    byteLength?: number
    truncated?: boolean
  }>(PLATFORM_BRIDGE_METHODS.FS_READ_BINARY, [{ path, maxBytes }])
  return {
    base64: typeof result.base64 === 'string' ? result.base64 : '',
    mimeType: typeof result.mimeType === 'string' ? result.mimeType : 'application/octet-stream',
    byteLength: typeof result.byteLength === 'number' ? result.byteLength : 0,
    truncated: Boolean(result.truncated),
  }
}

export async function writeTextFile(path: string, content: string): Promise<void> {
  if (isStandaloneDevMode()) {
    window.localStorage.setItem(`purescience:sheets:file:${path}`, content)
    return
  }
  await bridge.call(PLATFORM_BRIDGE_METHODS.FS_WRITE, [path, content])
}

// ---- Platform operations ledger ---------------------------------------------
// Suite-wide record of user/agent interactions in two lanes ('user'|'agent'),
// stored by the shell and rendered live by the PureAssistant tab. Building
// rule (see AGENTS.md "Operations ledger"): record every meaningful user or
// agent interaction this app performs, and when you find legacy activity/feed
// code duplicating this, tag it `DEPRECATED(operations-ledger)` for cleanup.
import { recordPlatformOperation as recordPlatformOperationBridge } from '@purescience/platform-ui/bridge/operations'
import type {
  PlatformOperation,
  PlatformOperationInput,
} from '@purescience/platform-ui/bridge/operations'

export type { PlatformOperation, PlatformOperationInput }

function operationsBridgeAvailable(): boolean {
  return !(import.meta.env.DEV && window.parent === window)
}

/** Record one interaction into the ledger (the shell pins appSlug to this app). */
export async function recordOperation(
  input: PlatformOperationInput,
): Promise<PlatformOperation | null> {
  if (!operationsBridgeAvailable()) return null
  return recordPlatformOperationBridge(input)
}
