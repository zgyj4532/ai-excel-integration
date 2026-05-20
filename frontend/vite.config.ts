import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { resolve } from 'path'
import dotenv from 'dotenv'

// load repo root .env so frontend dev server proxies to correct backend port
dotenv.config({ path: resolve(__dirname, '..', '.env') })

export default defineConfig({
  plugins: [vue()],
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src'),
    }
  },
  server: {
    port: 5173,
    proxy: (function () {
      const serverPort = process.env.SERVER_PORT || '8080'
      const target = `http://localhost:${serverPort}`
      return {
        '/api': {
          target,
          changeOrigin: true,
        }
      }
    })()
  },
  // Set base to ensure built assets use the application context path
  base: process.env.VITE_BASE || process.env.SERVER_CONTEXT_PATH || '/InsightCloud/',
  build: {
    outDir: '../src/main/resources/static', // 关键：打包到后端静态资源目录
    emptyOutDir: true,
    chunkSizeWarningLimit: 6000,
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (!id.includes('node_modules')) return undefined

          if (id.includes('/react-dom/')) return 'vendor-react-dom'
          if (id.includes('/react/')) return 'vendor-react'
          if (id.includes('/scheduler/')) return 'vendor-scheduler'
          if (id.includes('react-is')) return 'vendor-react-is'
          if (id.includes('use-sync-external-store')) return 'vendor-use-sync-external-store'

          if (id.includes('@floating-ui/')) return 'vendor-floating-ui'
          if (id.includes('@radix-ui/')) return 'vendor-radix-ui'
          if (id.includes('rc-dropdown')) return 'vendor-rc-dropdown'
          if (id.includes('rc-menu')) return 'vendor-rc-menu'
          if (id.includes('rc-virtual-list')) return 'vendor-rc-virtual-list'
          if (id.includes('react-grid-layout')) return 'vendor-react-grid-layout'
          if (id.includes('react-transition-group')) return 'vendor-react-transition-group'
          if (id.includes('class-variance-authority')) return 'vendor-class-variance-authority'
          if (id.includes('tailwind-merge')) return 'vendor-tailwind-merge'
          if (id.includes('sonner')) return 'vendor-sonner'
          if (id.includes('dayjs')) return 'vendor-dayjs'
          if (id.includes('clsx')) return 'vendor-clsx'
          if (id.includes('cjk-regex')) return 'vendor-cjk-regex'
          if (id.includes('franc-min')) return 'vendor-franc-min'
          if (id.includes('opentype.js')) return 'vendor-opentype'

          if (id.includes('@univerjs/engine-render/lib/') && !id.includes('/index.js')) return 'vendor-univer-engine-render-locales'
          if (id.includes('@univerjs/design/lib/locale/')) return 'vendor-univer-design-locales'
          if (id.includes('@univerjs/sheets-ui/lib/locale/')) return 'vendor-univer-sheets-ui-locales'
          if (id.includes('@univerjs/docs-ui/lib/locale/')) return 'vendor-univer-docs-ui-locales'
          if (id.includes('@univerjs/sheets-formula-ui/lib/locale/')) return 'vendor-univer-sheets-formula-ui-locales'
          if (id.includes('@univerjs/sheets-numfmt-ui/lib/locale/')) return 'vendor-univer-sheets-numfmt-ui-locales'
          if (id.includes('@univerjs/icons')) return 'vendor-univer-icons'
          if (id.includes('@univerjs/telemetry')) return 'vendor-univer-telemetry'
          if (id.includes('@univerjs/ui/lib/es/facade.js')) return 'vendor-univer-ui-facade'
          if (id.includes('@univerjs/ui/lib/es/index.js')) return 'vendor-univer-ui-base'
          if (id.includes('@univerjs/docs-ui/lib/es/facade.js')) return 'vendor-univer-docs-ui-facade'
          if (id.includes('@univerjs/docs-ui/lib/es/index.js')) return 'vendor-univer-docs-ui-base'
          if (id.includes('@univerjs/design/lib/es/index.js')) return 'vendor-univer-design-index'
          if (id.includes('@univerjs/design/lib/es/locale/')) return 'vendor-univer-design-locale'

          if (id.includes('@univerjs/sheets-ui/lib/es/facade.js')) return 'vendor-univer-sheets-ui-facade'
          if (id.includes('@univerjs/sheets-ui/lib/es/index.js')) return 'vendor-univer-sheets-ui-index'
          if (id.includes('@univerjs/docs-ui/lib/es/locale/')) return 'vendor-univer-docs-ui-locale'
          if (id.includes('@univerjs/ui/lib/es/locale/')) return 'vendor-univer-ui-locale'

          if (id.includes('@univerjs/preset-sheets-core')) return 'vendor-univer-preset-sheets'
          if (id.includes('@univerjs/sheets-formula-ui')) return 'vendor-univer-sheets-formula-ui'
          if (id.includes('@univerjs/sheets-formula')) return 'vendor-univer-sheets-formula'
          if (id.includes('@univerjs/sheets-numfmt-ui')) return 'vendor-univer-sheets-numfmt-ui'
          if (id.includes('@univerjs/sheets-numfmt')) return 'vendor-univer-sheets-numfmt'
          if (id.includes('@univerjs/sheets-ui')) return 'vendor-univer-sheets-ui'
          if (id.includes('@univerjs/sheets-thread-comment')) return 'vendor-univer-sheets-thread-comment'
          if (id.includes('@univerjs/sheets-find-replace')) return 'vendor-univer-sheets-find-replace'
          if (id.includes('@univerjs/sheets-drawing-ui')) return 'vendor-univer-sheets-drawing-ui'
          if (id.includes('@univerjs/sheets-note')) return 'vendor-univer-sheets-note'
          if (id.includes('@univerjs/sheets-hyper-link')) return 'vendor-univer-sheets-hyper-link'
          if (id.includes('@univerjs/engine-formula')) return 'vendor-univer-engine-formula'
          if (id.includes('@univerjs/engine-render')) return 'vendor-univer-engine-render'
          if (id.includes('@univerjs/docs-ui')) return 'vendor-univer-docs-ui'
          if (id.includes('@univerjs/docs')) return 'vendor-univer-docs'
          if (id.includes('@univerjs/ui-adapter-vue3')) return 'vendor-univer-ui-adapter'
          if (id.includes('@univerjs/network')) return 'vendor-univer-network'
          if (id.includes('@univerjs/rpc')) return 'vendor-univer-rpc'
          if (id.includes('@univerjs/design')) return 'vendor-univer-design'
          if (id.includes('@univerjs/protocol')) return 'vendor-univer-protocol'
          if (id.includes('@univerjs/ui')) return 'vendor-univer-ui'
          if (id.includes('@univerjs/core')) return 'vendor-univer-core'
          if (id.includes('@univerjs/themes')) return 'vendor-univer-themes'
          if (id.includes('@univerjs/sheets')) return 'vendor-univer-sheets'

          if (id.includes('html2canvas')) return 'vendor-html2canvas'
          if (id.includes('jspdf')) return 'vendor-jspdf'
          if (id.includes('xlsx')) return 'vendor-xlsx'

          return undefined
        }
      }
    }
  }
})
