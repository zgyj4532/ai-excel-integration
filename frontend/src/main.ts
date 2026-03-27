import { createApp } from 'vue'
import App from './App.vue'
import { router } from './router'
import './styles.css'
import { i18n } from './i18n'
import '../node_modules/vue-element-plus-x/dist/index.css'

const app = createApp(App)
app.use(router)
app.use(i18n)
app.mount('#app')
