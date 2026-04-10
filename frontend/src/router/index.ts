import { createRouter, createWebHistory } from 'vue-router'

const routes = [
  { path: '/', name: 'Landing', component: () => import('../pages/Landing.vue') },
  { path: '/workspace', name: 'Workspace', component: () => import('../pages/Workspace.vue') },
  { path: '/analysis', name: 'Analysis', component: () => import('../pages/Analysis.vue') },
  { path: '/templates', name: 'Templates', component: () => import('../pages/Templates.vue') },
  { path: '/settings', name: 'Settings', component: () => import('../pages/Settings.vue') }
]

export const router = createRouter({
  history: createWebHistory(),
  routes
})

export default router
