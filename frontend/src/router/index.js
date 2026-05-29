import { createRouter, createWebHistory } from 'vue-router'
import HomeView from '../views/HomeView.vue'

const routes = [
  { path: '/', name: 'home', component: HomeView },
  { path: '/upload', name: 'upload', component: () => import('../views/UploadView.vue') },
  { path: '/statistics', name: 'statistics', component: () => import('../views/StatisticsView.vue') },
  { path: '/voc', name: 'voc-list', component: () => import('../views/VocListView.vue') },
  { path: '/qdata', name: 'qdata', component: () => import('../views/QDataView.vue') },
  { path: '/settings', name: 'settings', component: () => import('../views/SettingsView.vue') },
  { path: '/launch-compare', name: 'launch-compare', component: () => import('../views/LaunchCompareView.vue') },
  { path: '/statistics/month/:month', name: 'month-detail', component: () => import('../views/MonthDetailView.vue') },
  { path: '/statistics/model/:name', name: 'model-detail', component: () => import('../views/ModelDetailView.vue') },
  { path: '/dev-issues', name: 'dev-issues', component: () => import('../views/DevIssuesView.vue') },
]

export default createRouter({
  history: createWebHistory(),
  routes,
})
