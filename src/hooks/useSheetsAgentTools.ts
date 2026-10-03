import { useRef } from 'react'
import { formatAgentToolJson } from '@purescience/platform-ui/bridge/agentToolHelpers'
import type { AgentToolHandler } from '@purescience/platform-ui/bridge/react/usePlatformAgentTools'
import { usePlatformAgentTools } from '@purescience/platform-ui/bridge/react/usePlatformAgentTools'
import {
  AgentSheetsToolError,
  PURESHEETS_AGENT_LOG_LABEL,
  PURESHEETS_AGENT_TOOL_NAMES,
  type SheetsAgentToolContext,
} from '../agents/catalog'
import {
  autoFitRowsHandler,
  addSheetHandler,
  deleteSheetHandler,
  formatRangeHandler,
  renameSheetHandler,
  applyCellChangesHandler,
  findInWorkbookHandler,
  getSheetsContextHandler,
  listSheetChangesHandler,
  addChartHandler,
  readRangeHandler,
  getExchangeRateHandler,
  setCellCommentHandler,
} from '../agents/handlers'

export function useSheetsAgentTools(
  ready: boolean,
  context: SheetsAgentToolContext,
  mayChangeWorkbook: () => boolean = () => true,
): void {
  const contextRef = useRef(context)
  contextRef.current = context
  const mayChangeRef = useRef(mayChangeWorkbook)
  mayChangeRef.current = mayChangeWorkbook

  const handlers: Record<string, AgentToolHandler> = {
    autoFitRows: async invoke =>
      autoFitRowsHandler(contextRef.current, invoke.arguments ?? {}),
    createWorkbook: async invoke => {
      const title = invoke.arguments?.title
      if (typeof title !== 'string' || !title.trim())
        throw new AgentSheetsToolError(
          'A non-empty workbook title is required.',
        )
      const create = contextRef.current.createWorkbook
      if (!create)
        throw new AgentSheetsToolError('Workbook creation is unavailable.')
      return { content: formatAgentToolJson(await create(title.trim())) }
    },
    completeWorkbook: async () => {
      const complete = contextRef.current.completeWorkbook
      if (!complete)
        throw new AgentSheetsToolError('Workbook saving is unavailable.')
      return { content: formatAgentToolJson(await complete()) }
    },
    getSheetsContext: async () => getSheetsContextHandler(contextRef.current),
    readRange: async invoke =>
      readRangeHandler(contextRef.current, invoke.arguments ?? {}),
    findInWorkbook: async invoke =>
      findInWorkbookHandler(contextRef.current, invoke.arguments ?? {}),
    applyCellChanges: async invoke =>
      applyCellChangesHandler(contextRef.current, invoke.arguments ?? {}),
    getExchangeRate: async invoke =>
      getExchangeRateHandler(contextRef.current, invoke.arguments ?? {}),
    addChart: async invoke =>
      addChartHandler(contextRef.current, invoke.arguments ?? {}),
    listSheetChanges: async invoke =>
      listSheetChangesHandler(contextRef.current, invoke.arguments ?? {}),
    addSheet: async invoke =>
      addSheetHandler(contextRef.current, invoke.arguments ?? {}),
    renameSheet: async invoke =>
      renameSheetHandler(contextRef.current, invoke.arguments ?? {}),
    deleteSheet: async invoke =>
      deleteSheetHandler(contextRef.current, invoke.arguments ?? {}),
    formatRange: async invoke =>
      formatRangeHandler(contextRef.current, invoke.arguments ?? {}),
    setCellComment: async invoke =>
      setCellCommentHandler(contextRef.current, invoke.arguments ?? {}),
  }
  usePlatformAgentTools({
    ready,
    tools: PURESHEETS_AGENT_TOOL_NAMES,
    logLabel: PURESHEETS_AGENT_LOG_LABEL,
    errorType: AgentSheetsToolError,
    handlers: Object.fromEntries(
      Object.entries(handlers).map(([name, handler]) => [
        name,
        async invoke => {
          if (
            ![
              'getSheetsContext',
              'readRange',
              'findInWorkbook',
              'listSheetChanges',
              'getExchangeRate',
              'createWorkbook',
            ].includes(name) &&
            !mayChangeRef.current()
          )
            throw new AgentSheetsToolError(
              'Wait for the workbook to finish opening, then read getSheetsContext again before changing it.',
            )
          return handler(invoke)
        },
      ]),
    ),
  })
}
