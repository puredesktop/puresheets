import { defineConfig } from 'vitest/config'
import react from '/Users/adamhyde/.puredev/suite/apps/puresheets/node_modules/@vitejs/plugin-react-swc/index.js'

export default defineConfig({
  plugins: [react()],
  resolve: {
    dedupe: ['react', 'react-dom', 'styled-components'],
    alias: [
      {
        find: /^@purescience\/platform-ui\/(.*)$/,
        replacement: '/Users/adamhyde/.puredev/suite/packages/ui/src/$1',
      },
      {
        find: /^@purescience\/platform-bridge\/(.*)$/,
        replacement: '/Users/adamhyde/.puredev/suite/packages/bridge/src/$1',
      },
    ],
  },
  esbuild: false,
  server: { fs: { allow: ['/Users/adamhyde/.puredev/suite', process.cwd()] } },
  test: { maxWorkers: 2 },
})
