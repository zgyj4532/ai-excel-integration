import { createApp } from 'vue'
import App from './App.vue'
import { router } from './router'
import './darkstyle.css'
import './lightstyle.css'
import './styles.css'
import { i18n } from './i18n'
import '../node_modules/vue-element-plus-x/dist/index.css'

// 在应用挂载前恢复主题，避免闪烁
const savedTheme = localStorage.getItem('app-theme') || 'dark'
document.documentElement.setAttribute('data-theme', savedTheme)

const app = createApp(App)
app.use(router)
app.use(i18n)
app.mount('#app')
