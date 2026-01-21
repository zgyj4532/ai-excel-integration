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
  build: {
    outDir: '../src/main/resources/static', // 关键：打包到后端静态资源目录
    emptyOutDir: true,
  }
})
