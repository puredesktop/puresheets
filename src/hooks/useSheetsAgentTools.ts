import { useRef } from 'react'
import type { AgentToolInvokeContext } from '@purescience/platform-ui/bridge/react/usePlatformAgentTools'
import { usePlatformAgentTools } from '@purescience/platform-ui/bridge/react/usePlatformAgentTools'
import {
  AgentSheetsToolError,
  EXPOSED_TOOL_NAMES,
  PURESHEETS_AGENT_LOG_LABEL,
  type SheetsAgentToolContext,
} from '../agents/catalog'
import { buildPureSheetsHandlers } from '../agents/operationsRegistry'

export function useSheetsAgentTools(
  ready: boolean,
  context: SheetsAgentToolContext,
  mayChangeWorkbook: () => boolean = () => true,
): void {
  const contextRef = useRef(context)
  contextRef.current = context
  const mayChangeRef = useRef(mayChangeWorkbook)
  mayChangeRef.current = mayChangeWorkbook

  const registryHandlers = buildPureSheetsHandlers(() => contextRef.current, {
    beforeInvoke: guard => {
      if (
        guard.risk !== 'read' &&
        guard.name !== 'createWorkbook' &&
        !mayChangeRef.current()
      ) {
        throw new AgentSheetsToolError(
          'Wait for the workbook to finish opening, then read getSheetsContext again before changing it.',
        )
      }
    },
  })

  // Statically key-preserving registration so `scripts/check-manifests.mjs`
  // verifies every manifest tool has a live registered handler at build time.
  usePlatformAgentTools({
    ready,
    tools: EXPOSED_TOOL_NAMES,
    logLabel: PURESHEETS_AGENT_LOG_LABEL,
    errorType: AgentSheetsToolError,
    handlers: Object.fromEntries(
      EXPOSED_TOOL_NAMES.map(name => [
        name,
        async (invoke: AgentToolInvokeContext) =>
          registryHandlers[name]!(invoke),
      ]),
    ),
  })
}
