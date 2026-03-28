<template>
  <div class="home">
    <h1 class="page-title">대시보드</h1>

    <div v-if="loading" class="loading">로딩 중...</div>

    <div v-else-if="data" class="cards">
      <div class="card stat-card">
        <div class="stat-label">전체 VOC</div>
        <div class="stat-value">{{ data.total_count?.toLocaleString() }}</div>
      </div>
      <div class="card stat-card">
        <div class="stat-label">{{ data.yesterday_date }} 신규</div>
        <div class="stat-value highlight">{{ data.daily_count }}</div>
      </div>
    </div>

    <div v-if="data?.top10_models?.length" class="card mt-20">
      <h2 class="section-title">전일 Top 10 모델</h2>
      <table class="table">
        <thead>
          <tr><th>순위</th><th>모델명</th><th>건수</th></tr>
        </thead>
        <tbody>
          <tr v-for="(item, idx) in data.top10_models" :key="item.model_name">
            <td>{{ idx + 1 }}</td>
            <td>{{ item.model_name }}</td>
            <td><span class="badge">{{ item.count }}</span></td>
          </tr>
        </tbody>
      </table>
    </div>

    <div v-if="!data && !loading" class="empty-state">
      <p>데이터가 없습니다. VOC 파일을 업로드해주세요.</p>
      <RouterLink to="/upload" class="btn-primary">업로드 바로가기</RouterLink>
    </div>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { getDashboard } from '../api'

const data = ref(null)
const loading = ref(false)

onMounted(async () => {
  loading.value = true
  try {
    const res = await getDashboard()
    data.value = res.data
  } catch (e) {
    console.error(e)
  } finally {
    loading.value = false
  }
})
</script>

<style scoped>
.page-title { font-size: 1.5rem; font-weight: 700; margin-bottom: 20px; color: #1a237e; }
.cards { display: flex; gap: 16px; flex-wrap: wrap; }
.card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }
.stat-card { min-width: 180px; }
.stat-label { font-size: 0.85rem; color: #666; margin-bottom: 8px; }
.stat-value { font-size: 2rem; font-weight: 700; color: #333; }
.stat-value.highlight { color: #1a237e; }
.mt-20 { margin-top: 20px; width: 100%; }
.section-title { font-size: 1rem; font-weight: 600; margin-bottom: 14px; color: #333; }
.table { width: 100%; border-collapse: collapse; }
.table th, .table td { padding: 10px 14px; text-align: left; border-bottom: 1px solid #eee; font-size: 0.9rem; }
.table th { background: #f8f9ff; font-weight: 600; color: #555; }
.badge { background: #e8eaf6; color: #1a237e; padding: 2px 10px; border-radius: 12px; font-weight: 600; }
.loading { text-align: center; padding: 60px; color: #888; }
.empty-state { text-align: center; padding: 60px; color: #888; }
.btn-primary { display: inline-block; margin-top: 12px; padding: 10px 24px; background: #1a237e; color: #fff; border-radius: 6px; text-decoration: none; }
</style>
