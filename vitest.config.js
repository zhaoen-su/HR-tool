import { defineConfig } from 'vitest/config'

export default defineConfig({
    test: {
        // 讓你可以在測試檔案中直接使用 describe, it, expect 而不用每次 import
        globals: true,
        // 設定測試環境，GAS 邏輯大多是純 JS，用 node 即可
        environment: 'node',
        // 包含哪些檔案
        include: ['**/*.{test,spec}.{js,ts}'],
        setupFiles: ['./setup.js'],
    },
})