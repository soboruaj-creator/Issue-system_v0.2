<template>
  <nav class="navbar">
    <div class="navbar-brand">
      <RouterLink to="/">📊 VOC 관리 시스템</RouterLink>
    </div>
    <ul class="navbar-menu">
      <li><RouterLink to="/" :class="{ active: route.path === '/' }">대시보드</RouterLink></li>
      <li><RouterLink to="/upload" :class="{ active: route.path === '/upload' }">업로드</RouterLink></li>
      <li><RouterLink to="/statistics" :class="{ active: route.path === '/statistics' }">통계</RouterLink></li>
      <li><RouterLink to="/voc" :class="{ active: route.path === '/voc' }">M-issue</RouterLink></li>
      <li><RouterLink to="/qdata" :class="{ active: route.path === '/qdata' }">Q-data</RouterLink></li>
      <li><RouterLink to="/launch-compare" :class="{ active: route.path === '/launch-compare' }">모델별 비교 분석</RouterLink></li>
      <li><RouterLink to="/dev-issues" :class="{ active: route.path === '/dev-issues' }">개발이슈</RouterLink></li>
      <li><RouterLink to="/settings" :class="{ active: route.path === '/settings' }">설정</RouterLink></li>
    </ul>
  </nav>
</template>

<script setup>
import { useRoute } from 'vue-router'
const route = useRoute()
</script>

<style scoped>
.navbar {
  background: #1a237e;
  color: #fff;
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 24px;
  height: 56px;
  box-shadow: 0 2px 6px rgba(0,0,0,0.2);
}
.navbar-brand a {
  color: #fff;
  text-decoration: none;
  font-size: 1.1rem;
  font-weight: 700;
}
.navbar-menu {
  list-style: none;
  display: flex;
  gap: 4px;
}
.navbar-menu a {
  color: rgba(255,255,255,0.8);
  text-decoration: none;
  padding: 8px 14px;
  border-radius: 6px;
  font-size: 0.9rem;
  transition: background 0.2s;
}
.navbar-menu a:hover,
.navbar-menu a.active {
  background: rgba(255,255,255,0.15);
  color: #fff;
}
</style>
