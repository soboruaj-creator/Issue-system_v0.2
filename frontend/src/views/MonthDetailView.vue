<template>
  <div class="month-detail-view">
    <!-- 헤더 -->
    <div class="header-row">
      <button class="btn-back" @click="window.close()">✕ 닫기</button>
      <div class="month-nav">
        <button class="btn-nav" @click="goMonth(-1)">◀</button>
        <h1 class="page-title">{{ month }} 월별 상세</h1>
        <button class="btn-nav" @click="goMonth(1)">▶</button>
      </div>
      <span></span>
    </div>

    <div v-if="loading" class="loading">불러오는 중...</div>
    <div v-else-if="!data" class="no-data">데이터가 없습니다.</div>

    <template v-else>
      <!-- 요약 카드 -->
      <div class="summary-cards">
        <div class="summary-card voc-card">
          <div class="card-label">Members issue</div>
          <div class="card-value">{{ data.voc.total.toLocaleString() }}</div>
          <div class="card-diff" :class="diffClass(data.voc.total - data.voc.prev_total)">
            <span>{{ diffLabel(data.voc.total, data.voc.prev_total) }}</span>
            <span class="prev-val">전월 {{ data.voc.prev_total.toLocaleString() }}</span>
          </div>
        </div>
        <div class="summary-card qdata-card">
          <div class="card-label">Q-data</div>
          <div class="card-value">{{ data.qdata.total.toLocaleString() }}</div>
          <div class="card-diff" :class="diffClass(data.qdata.total - data.qdata.prev_total)">
            <span>{{ diffLabel(data.qdata.total, data.qdata.prev_total) }}</span>
            <span class="prev-val">전월 {{ data.qdata.prev_total.toLocaleString() }}</span>
          </div>
        </div>
      </div>

      <!-- 3개월 연속 Top5 -->
      <div v-if="data.streak_models.length" class="card streak-card">
        <div class="section-title">🔥 3개월 연속 Top5 모델</div>
        <div class="streak-chips">
          <span v-for="m in data.streak_models" :key="m" class="streak-chip">{{ m }}</span>
        </div>
      </div>

      <div class="two-col">
        <!-- VOC 모델별 -->
        <div class="card">
          <div class="section-title">Members issue 모델별</div>
          <div class="table-scroll">
            <table class="table">
              <thead>
                <tr>
                  <th>#</th>
                  <th>모델</th>
                  <th>건수</th>
                  <th>전월</th>
                  <th>증감</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="(row, idx) in data.voc.models" :key="row.name" :class="{ 'row-zero': row.count === 0 }">
                  <td class="rank">{{ idx + 1 }}</td>
                  <td class="name-cell">
                    {{ row.name }}
                    <span v-if="row.is_new" class="badge new-badge">NEW</span>
                    <span v-if="row.is_top5_streak" class="badge streak-badge">🔥</span>
                  </td>
                  <td class="count">{{ row.count }}</td>
                  <td class="prev">{{ row.prev_count }}</td>
                  <td :class="diffClass(row.diff)">
                    {{ row.diff > 0 ? '+' + row.diff : row.diff }}
                    <span v-if="row.diff_pct !== null" class="pct">({{ row.diff > 0 ? '+' : '' }}{{ row.diff_pct }}%)</span>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <!-- Q-data 모델별 -->
        <div class="card">
          <div class="section-title">Q-data 모델별</div>
          <div class="table-scroll">
            <table class="table">
              <thead>
                <tr>
                  <th>#</th>
                  <th>모델</th>
                  <th>건수</th>
                  <th>전월</th>
                  <th>증감</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="(row, idx) in data.qdata.models" :key="row.name" :class="{ 'row-zero': row.count === 0 }">
                  <td class="rank">{{ idx + 1 }}</td>
                  <td class="name-cell">{{ row.name }}</td>
                  <td class="count">{{ row.count }}</td>
                  <td class="prev">{{ row.prev_count }}</td>
                  <td :class="diffClass(row.diff)">
                    {{ row.diff > 0 ? '+' + row.diff : row.diff }}
                    <span v-if="row.diff_pct !== null" class="pct">({{ row.diff > 0 ? '+' : '' }}{{ row.diff_pct }}%)</span>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- 칩셋 분포 -->
      <div class="two-col">
        <div class="card">
          <div class="section-title">VOC 칩셋 분포</div>
          <div class="chart-wrap-sm">
            <Doughnut v-if="chipsetChartData" :data="chipsetChartData" :options="doughnutOptions" />
          </div>
          <div class="table-scroll mt-8">
            <table class="table">
              <thead><tr><th>칩셋</th><th>건수</th><th>비율</th></tr></thead>
              <tbody>
                <tr v-for="c in data.voc.chipsets" :key="c.chipset">
                  <td>{{ c.chipset }}</td>
                  <td>{{ c.count }}</td>
                  <td>{{ chipsetPct(c.count) }}%</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <!-- 주차별 분포 -->
        <div class="card">
          <div class="section-title">주차별 VOC / Q-data 분포</div>
          <div class="chart-wrap-sm">
            <Bar v-if="weeklyChartData" :data="weeklyChartData" :options="barOptions" />
          </div>
        </div>
      </div>
    </template>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { Doughnut, Bar } from 'vue-chartjs'
import {
  Chart as ChartJS, ArcElement, CategoryScale, LinearScale,
  BarElement, Title, Tooltip, Legend,
} from 'chart.js'
import { getMonthDetail } from '../api'

ChartJS.register(ArcElement, CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend)

const route = useRoute()
const router = useRouter()

const month = ref(route.params.month)
const data = ref(null)
const loading = ref(false)

async function load() {
  loading.value = true
  data.value = null
  try {
    const res = await getMonthDetail(month.value)
    data.value = res.data
  } catch (e) {
    console.error(e)
  } finally {
    loading.value = false
  }
}

function goMonth(delta) {
  const [y, m] = month.value.split('-').map(Number)
  const d = new Date(y, m - 1 + delta, 1)
  const ny = d.getFullYear()
  const nm = String(d.getMonth() + 1).padStart(2, '0')
  month.value = `${ny}-${nm}`
  router.replace(`/statistics/month/${month.value}`)
}

function diffClass(diff) {
  if (diff > 0) return 'diff-up'
  if (diff < 0) return 'diff-down'
  return 'diff-zero'
}

function diffLabel(cur, prev) {
  const d = cur - prev
  if (d === 0) return '±0'
  const sign = d > 0 ? '+' : ''
  const pct = prev > 0 ? ` (${sign}${Math.round(d / prev * 100)}%)` : ''
  return `${sign}${d}${pct}`
}

const chipsetTotal = computed(() =>
  data.value ? data.value.voc.chipsets.reduce((s, c) => s + c.count, 0) : 0
)

function chipsetPct(count) {
  return chipsetTotal.value ? Math.round(count / chipsetTotal.value * 100) : 0
}

const PALETTE = [
  '#3f51b5','#e91e63','#ff9800','#4caf50','#00bcd4',
  '#9c27b0','#f44336','#2196f3','#8bc34a','#ff5722',
]

const chipsetChartData = computed(() => {
  if (!data.value?.voc.chipsets.length) return null
  return {
    labels: data.value.voc.chipsets.map(c => c.chipset),
    datasets: [{
      data: data.value.voc.chipsets.map(c => c.count),
      backgroundColor: data.value.voc.chipsets.map((_, i) => PALETTE[i % PALETTE.length]),
    }],
  }
})

const weeklyChartData = computed(() => {
  if (!data.value) return null
  const labels = ['1주', '2주', '3주', '4주', '5주']
  const voc_by_week = Object.fromEntries(data.value.voc.weekly.map(w => [w.week, w.count]))
  const qdata_by_week = Object.fromEntries(data.value.qdata.weekly.map(w => [w.week, w.count]))
  return {
    labels,
    datasets: [
      {
        label: 'VOC',
        data: [1,2,3,4,5].map(w => voc_by_week[w] || 0),
        backgroundColor: 'rgba(63,81,181,0.7)',
      },
      {
        label: 'Q-data',
        data: [1,2,3,4,5].map(w => qdata_by_week[w] || 0),
        backgroundColor: 'rgba(255,152,0,0.7)',
      },
    ],
  }
})

const doughnutOptions = {
  responsive: true,
  plugins: { legend: { position: 'right', labels: { font: { size: 11 } } } },
}

const barOptions = {
  responsive: true,
  plugins: { legend: { position: 'top' } },
  scales: { x: {}, y: { beginAtZero: true } },
}

onMounted(load)
watch(() => route.params.month, (val) => {
  if (val) { month.value = val; load() }
})
</script>

<style scoped>
.month-detail-view { padding: 4px 0; }
.header-row { display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; }
.btn-back { padding: 7px 16px; background: #e8eaf6; color: #1a237e; border: none; border-radius: 6px; cursor: pointer; font-size: 0.9rem; font-weight: 600; }
.btn-back:hover { background: #c5cae9; }
.month-nav { display: flex; align-items: center; gap: 12px; }
.btn-nav { padding: 6px 12px; background: #f5f5f5; border: 1px solid #ddd; border-radius: 6px; cursor: pointer; font-size: 0.9rem; }
.btn-nav:hover { background: #e8eaf6; }
.page-title { font-size: 1.3rem; font-weight: 700; color: #1a237e; margin: 0; }

.summary-cards { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 16px; }
.summary-card { background: #fff; border-radius: 10px; padding: 18px 24px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }
.voc-card { border-left: 4px solid #3f51b5; }
.qdata-card { border-left: 4px solid #ff9800; }
.card-label { font-size: 0.85rem; color: #888; margin-bottom: 6px; }
.card-value { font-size: 2rem; font-weight: 700; color: #1a237e; }
.card-diff { font-size: 0.88rem; margin-top: 4px; display: flex; gap: 10px; align-items: center; }
.prev-val { color: #aaa; }

.streak-card { background: #fff8e1; border-left: 4px solid #ffa000; }
.streak-chips { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 8px; }
.streak-chip { padding: 4px 14px; background: #fff3e0; border: 1.5px solid #ffa000; border-radius: 16px; font-size: 0.88rem; font-weight: 600; color: #e65100; }

.two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 16px; }
.card { background: #fff; border-radius: 10px; padding: 18px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); margin-bottom: 16px; }
.section-title { font-size: 0.95rem; font-weight: 600; color: #333; margin-bottom: 12px; }

.table-scroll { overflow-x: auto; }
.table { width: 100%; border-collapse: collapse; font-size: 0.83rem; }
.table th, .table td { padding: 7px 10px; border-bottom: 1px solid #f0f0f0; text-align: center; white-space: nowrap; }
.table th { background: #f8f9ff; font-weight: 600; color: #555; }
.rank { color: #aaa; }
.name-cell { text-align: left; max-width: 160px; overflow: hidden; text-overflow: ellipsis; }
.count { font-weight: 600; color: #1a237e; }
.prev { color: #888; }
.pct { font-size: 0.78rem; color: inherit; opacity: 0.75; }
.row-zero { opacity: 0.5; }

.diff-up { color: #b71c1c; font-weight: 600; }
.diff-down { color: #1b5e20; font-weight: 600; }
.diff-zero { color: #888; }

.badge { font-size: 0.7rem; padding: 1px 6px; border-radius: 8px; margin-left: 4px; vertical-align: middle; }
.new-badge { background: #e8f5e9; color: #2e7d32; border: 1px solid #a5d6a7; font-weight: 700; }
.streak-badge { background: transparent; }

.chart-wrap-sm { height: 220px; }
.mt-8 { margin-top: 8px; }

.loading { text-align: center; padding: 60px; color: #888; font-size: 1rem; }
.no-data { text-align: center; padding: 60px; color: #aaa; }

@media (max-width: 768px) {
  .two-col { grid-template-columns: 1fr; }
  .summary-cards { grid-template-columns: 1fr; }
}
</style>
