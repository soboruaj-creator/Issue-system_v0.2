<template>
  <div class="home">
    <h1 class="page-title">대시보드</h1>

    <div v-if="loading" class="loading">로딩 중...</div>

    <template v-else-if="data">

      <!-- 상단 요약 카드 -->
      <div class="stat-row">
        <div class="card stat-card">
          <div class="stat-label">Members issue <span class="stat-date">{{ data.yesterday }}</span></div>
          <div class="stat-value voc">{{ data.members_issue.yesterday_count.toLocaleString() }}<span class="stat-unit">건</span></div>
        </div>
        <div class="card stat-card">
          <div class="stat-label">Q-data <span class="stat-date">{{ data.yesterday }}</span></div>
          <div class="stat-value qdata">{{ data.qdata.yesterday_count.toLocaleString() }}<span class="stat-unit">건</span></div>
        </div>
        <div class="card stat-card upload-card">
          <div class="stat-label">마지막 업로드</div>
          <div class="upload-row">
            <span class="upload-type voc">Members</span>
            <span class="upload-time">{{ formatUpload(data.upload_status.members_issue) }}</span>
          </div>
          <div class="upload-row">
            <span class="upload-type qdata">Q-data</span>
            <span class="upload-time">{{ formatUpload(data.upload_status.qdata) }}</span>
          </div>
        </div>
      </div>

      <!-- 2열: Members issue 리스트 | Q-data Top5 -->
      <div class="two-col">

        <!-- Members issue 어제 인입 리스트 -->
        <div class="card">
          <h2 class="section-title">Members issue 어제 인입 <span class="count-badge voc">{{ data.members_issue.yesterday_count }}건</span></h2>
          <div v-if="!data.members_issue.entries.length" class="no-data">어제 인입 건 없음</div>
          <div v-else class="entry-list">
            <div v-for="(e, idx) in data.members_issue.entries" :key="idx" class="entry-item">
              <span class="entry-model">{{ e.model_name }}</span>
              <span class="entry-summary">{{ e.summary || '(내용 없음)' }}</span>
            </div>
          </div>
        </div>

        <!-- Q-data Top5 + WoW -->
        <div class="col-right">
          <div class="card">
            <h2 class="section-title">Q-data Top 5 <span class="stat-date">{{ data.yesterday }}</span></h2>
            <div v-if="!data.qdata.top5_yesterday.length" class="no-data">데이터 없음</div>
            <table v-else class="table">
              <thead><tr><th>순위</th><th>모델명</th><th>건수</th></tr></thead>
              <tbody>
                <tr v-for="(item, idx) in data.qdata.top5_yesterday" :key="item.model_name">
                  <td>{{ idx + 1 }}</td>
                  <td>{{ item.model_name }}</td>
                  <td><span class="badge qdata">{{ item.count }}</span></td>
                </tr>
              </tbody>
            </table>
          </div>
          <div class="card mt-12">
            <h2 class="section-title">Q-data 지난주 대비 증가 Top 5</h2>
            <div v-if="!data.qdata.top5_wow.length" class="no-data">비교 데이터 없음</div>
            <table v-else class="table">
              <thead><tr><th>모델명</th><th>이번 7일</th><th>이전 7일</th><th>증가</th></tr></thead>
              <tbody>
                <tr v-for="item in data.qdata.top5_wow" :key="item.model_name">
                  <td>{{ item.model_name }}</td>
                  <td>{{ item.this_week }}</td>
                  <td>{{ item.last_week }}</td>
                  <td><span class="badge up">+{{ item.diff }}</span></td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- 4주 추이 차트 -->
      <div class="card mt-16">
        <h2 class="section-title">최근 4주 추이</h2>
        <div class="chart-wrap">
          <Bar v-if="trendChartData" :data="trendChartData" :options="barOptions" />
        </div>
      </div>

      <!-- 이상 탐지 -->
      <div v-if="data.members_issue.anomaly.length" class="card mt-16">
        <h2 class="section-title">이상 탐지 <span class="anomaly-badge">7일 평균 대비 2배↑</span></h2>
        <table class="table">
          <thead><tr><th>모델명</th><th>어제 건수</th><th>7일 평균</th><th>배율</th></tr></thead>
          <tbody>
            <tr v-for="item in data.members_issue.anomaly" :key="item.model_name" class="anomaly-row">
              <td>{{ item.model_name }}</td>
              <td><span class="badge voc">{{ item.yesterday_count }}</span></td>
              <td>{{ item.avg_7d }}</td>
              <td><span v-if="item.ratio" class="badge warn">×{{ item.ratio }}</span><span v-else class="badge warn">신규</span></td>
            </tr>
          </tbody>
        </table>
      </div>

    </template>

    <div v-else-if="!loading" class="empty-state">
      <p>데이터가 없습니다. VOC 파일을 업로드해주세요.</p>
      <RouterLink to="/upload" class="btn-primary">업로드 바로가기</RouterLink>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'
import { Bar } from 'vue-chartjs'
import {
  Chart as ChartJS, CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend,
} from 'chart.js'
import { getDashboard } from '../api'

ChartJS.register(CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend)

const data = ref(null)
const loading = ref(false)

function formatUpload(dt) {
  if (!dt) return '없음'
  return dt.slice(0, 16)
}

const trendChartData = computed(() => {
  const trend = data.value?.weekly_trend
  if (!trend?.length) return null
  return {
    labels: trend.map(t => t.label),
    datasets: [
      {
        label: 'Members issue',
        data: trend.map(t => t.members_issue),
        backgroundColor: 'rgba(26,35,126,0.75)',
        borderRadius: 4,
      },
      {
        label: 'Q-data',
        data: trend.map(t => t.qdata),
        backgroundColor: 'rgba(230,74,25,0.75)',
        borderRadius: 4,
      },
    ],
  }
})

const barOptions = {
  responsive: true,
  plugins: { legend: { display: true, position: 'top' } },
  scales: { x: { grid: { display: false } } },
}

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
.loading { text-align: center; padding: 60px; color: #888; }

.stat-row { display: flex; gap: 14px; flex-wrap: wrap; margin-bottom: 16px; }
.card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }
.stat-card { min-width: 180px; flex: 1; }
.stat-label { font-size: 0.82rem; color: #888; margin-bottom: 8px; }
.stat-date { font-weight: 600; color: #555; }
.stat-value { font-size: 2rem; font-weight: 800; line-height: 1; }
.stat-value.voc { color: #1a237e; }
.stat-value.qdata { color: #bf360c; }
.stat-unit { font-size: 1rem; font-weight: 400; margin-left: 2px; }
.upload-card { min-width: 220px; }
.upload-row { display: flex; align-items: center; gap: 8px; margin-top: 8px; }
.upload-type { font-size: 0.75rem; font-weight: 700; padding: 2px 7px; border-radius: 10px; }
.upload-type.voc { background: #e8eaf6; color: #1a237e; }
.upload-type.qdata { background: #fbe9e7; color: #bf360c; }
.upload-time { font-size: 0.85rem; color: #555; }

.two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 16px; }
@media (max-width: 960px) { .two-col { grid-template-columns: 1fr; } }
.col-right { display: flex; flex-direction: column; gap: 0; }

.section-title { font-size: 0.95rem; font-weight: 700; margin-bottom: 12px; color: #333; display: flex; align-items: center; gap: 8px; }
.count-badge { font-size: 0.78rem; padding: 2px 8px; border-radius: 10px; font-weight: 600; }
.count-badge.voc { background: #e8eaf6; color: #1a237e; }

.entry-list { max-height: 420px; overflow-y: auto; display: flex; flex-direction: column; gap: 6px; }
.entry-item { display: flex; gap: 8px; align-items: flex-start; padding: 7px 8px; border-radius: 6px; background: #f8f9ff; font-size: 0.83rem; }
.entry-model { font-weight: 700; color: #1a237e; white-space: nowrap; min-width: 72px; }
.entry-summary { color: #444; line-height: 1.4; overflow: hidden; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
.no-data { color: #aaa; font-size: 0.85rem; padding: 12px 0; }

.table { width: 100%; border-collapse: collapse; }
.table th, .table td { padding: 8px 10px; text-align: left; border-bottom: 1px solid #eee; font-size: 0.85rem; }
.table th { background: #f8f9ff; font-weight: 600; color: #666; }
.badge { padding: 2px 9px; border-radius: 10px; font-weight: 700; font-size: 0.82rem; }
.badge.voc { background: #e8eaf6; color: #1a237e; }
.badge.qdata { background: #fbe9e7; color: #bf360c; }
.badge.up { background: #e8f5e9; color: #2e7d32; }
.badge.warn { background: #fff3e0; color: #e65100; }

.chart-wrap { height: 200px; }
.mt-12 { margin-top: 12px; }
.mt-16 { margin-top: 16px; }

.anomaly-badge { font-size: 0.75rem; padding: 2px 8px; background: #fff3e0; color: #e65100; border-radius: 10px; font-weight: 600; }
.anomaly-row { background: #fffde7; }

.empty-state { text-align: center; padding: 60px; color: #888; }
.btn-primary { display: inline-block; margin-top: 12px; padding: 10px 24px; background: #1a237e; color: #fff; border-radius: 6px; text-decoration: none; }
</style>
