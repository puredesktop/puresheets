import {
  createAppOperationsRegistry,
  defineAppTool,
  type AppAtomicToolDefinition,
  type BuildHandlersOptions,
} from '@purescience/platform-ui/bridge/appFunctions'
import type { AgentToolHandler } from '@purescience/platform-ui/bridge/react/usePlatformAgentTools'
import { PURESHEETS_ATOMIC_TOOL_SCHEMAS } from './atomicToolSchemas'
import {
  AgentSheetsToolError,
  type SheetsAgentToolContext,
} from './catalog'
import { PURESHEETS_FUNCTIONS } from './functions'
import { PURESHEETS_ATOMIC_RUNNERS } from './handlers'

const atomicTools: AppAtomicToolDefinition<SheetsAgentToolContext>[] =
  PURESHEETS_ATOMIC_TOOL_SCHEMAS.map(schema =>
    defineAppTool<SheetsAgentToolContext>({
      ...schema,
      run: async (ctx, args) => {
        const runner = PURESHEETS_ATOMIC_RUNNERS[schema.name]
        if (!runner) {
          throw new AgentSheetsToolError(
            `Missing atomic tool handler for "${schema.name}".`,
          )
        }
        return runner(ctx, args)
      },
    }),
  )

/**
 * PureSheets' unified Operations Registry built on `@purescience/platform-ui/bridge/appFunctions`:
 *   - 5 composite `function` workflows (`fn_*`, always exposed)
 *   - 4 core atomic tools (~24% of 17, always exposed)
 *   - 3 universal escape-hatch meta-tools (always exposed)
 *   - 13 extended specialist atomic tools (hidden by default, 100% reachable on demand)
 */
export const PURESHEETS_OPERATIONS_REGISTRY =
  createAppOperationsRegistry<SheetsAgentToolContext>({
    appSlug: 'sheets',
    appName: 'PureSheets',
    functions: PURESHEETS_FUNCTIONS,
    tools: atomicTools,
  })

export function buildPureSheetsHandlers(
  getCtx: () => SheetsAgentToolContext,
  options: BuildHandlersOptions<SheetsAgentToolContext> = {},
): Record<string, AgentToolHandler> {
  return PURESHEETS_OPERATIONS_REGISTRY.buildHandlers(getCtx, {
    createError: message => new AgentSheetsToolError(message),
    ...options,
  })
}
