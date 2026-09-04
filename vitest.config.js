import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    globals: true,
    environment: 'node',
    include: ['server/tests/**/*.test.js'],
    exclude: ['server/tests/**/*.benchmark.js'],
    setupFiles: ['server/tests/helpers/env.js'],
    testTimeout: 30000,
    hookTimeout: 30000,
    teardownTimeout: 30000,
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html'],
      exclude: [
        'node_modules/**',
        'server/tests/**',
        '**/*.config.js',
        '**/*.benchmark.js'
      ]
    }
  }
});