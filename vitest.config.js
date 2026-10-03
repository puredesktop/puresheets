import { defineConfig } from 'vitest/config'
export default defineConfig({
  resolve: { dedupe: ['react', 'react-dom', 'styled-components'] },
  test: { maxWorkers: 2 },
})
