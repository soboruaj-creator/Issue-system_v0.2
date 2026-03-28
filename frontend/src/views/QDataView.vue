<template>
  <div class="qdata-view">
    <h1 class="page-title">Q-data 관리</h1>

    <!-- 탭 -->
    <div class="tabs">
      <button :class="['tab-btn', { active: activeTab === 'list' }]" @click="activeTab='list'">데이터 목록</button>
      <button :class="['tab-btn', { active: activeTab === 'stats' }]" @click="activeTab='stats'">통계</button>
    </div>

    <!-- 데이터 목록 -->
    <template v-if="activeTab === 'list'">
      <div class="card filter-bar">
        <input v-model="modelFilter" placeholder="모델명 필터" class="filter-input" @keyup.enter="fetchList" />
        <input type="date" v-model="startDate" />
        <input type="date" v-model="endDate" />
        <button class="btn-primary" @click="fetchList">검색</button>
        <button class="btn-secondary" @click="exportData">엑셀 다운로드</button>
        <button class="btn-danger" @click="checkDuplicates">중복 확인</button>
      </div>

      <div v-if="duplicateInfo" class="card dup-card">
        <span>중복 데이터: <strong>{{ duplicateInfo.duplicate_count }}</strong>건</span>
        <button v-if="duplicateInfo.duplicate_count > 0" class="btn-danger" @click="removeDuplicates">중복 제거</button>
      </div>

      <div v-if="loading" class="loading">로딩 중...</div>
      <div v-else class="card table-card">
        <div class="table-header">
          <span class="total-count">총 {{ total.toLocaleString() }}건</span>
        </div>
        <div class="table-wrap">
          <table class="table">
            <thead>
              <tr>
                <th>서비스일</th><th>처리유형</th><th>수리명</th><th>수리상세</th>
                <th>모델명</th><th>시리얼번호</th><th>SW 이전</th><th>SW 이후</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="item in list" :key="item._id">
                <td>{{ item.service_date }}</td>
                <td>{{ item.process_type || '-' }}</td>
                <td>{{ item.repair_name || '-' }}</td>
                <td class="content-cell">{{ truncate(item.repair_detail, 40) }}</td>
                <td>{{ item.model_name || '-' }}</td>
                <td><code>{{ item.serial_number || '-' }}</code></td>
                <td>{{ item.sw_before || '-' }}</td>
                <td>{{ item.sw_after || '-' }}</td>
              </tr>
              <tr v-if="!list.length">
                <td colspan="8" class="empty">데이터가 없습니다.</td>
              </tr>
            </tbody>
          </table>
        </div>
        <div class="pagination">
          <button @click="changePage(page-1)" :disabled="page<=1">이전</button>
          <span>{{ page }} / {{ totalPages }}</span>
          <button @click="changePage(page+1)" :disabled="page>=totalPages">다음</button>
        </div>
      </div>
    </template>

    <!-- 통계 -->
    <template v-if="activeTab === 'stats'">
      <div class="card filter-bar">
        <input type="date" v-model="statsStart" />
        <input type="date" v-model="statsEnd" />
        <button class="btn-primary" @click="loadStats">조회</button>
      </div>

      <div class="stats-grid">
        <div class="card">
          <h2 class="section-title">모델별 건수</h2>
          <div class="chart-wrap">
            <Bar v-if="modelChartData" :data="modelChartData" :options="barOpts" />
          </div>
          <table class="table mt-12">
            <thead><tr><th>순위</th><th>모델</th><th>건수</th></tr></thead>
            <tbody>
              <tr v-for="(item, idx) in modelStats" :key="item.model_name">
                <td>{{ idx+1 }}</td><td>{{ item.model_name }}</td>
                <td><span class="badge">{{ item.count }}</span></td>
              </tr>
            </tbody>
          </table>
        </div>

        <div class="card">
          <h2 class="section-title">월별 추이</h2>
          <div class="chart-wrap">
            <Line v-if="monthlyChartData" :data="monthlyChartData" :options="lineOpts" />
          </div>
        </div>
      </div>
    </template>
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'
import { Bar, Line } from 'vue-chartjs'
import { Chart as ChartJS, CategoryScale, LinearScale, BarElement, LineElement, PointElement, Title, Tooltip, Legend } from 'chart.js'
import * as api from '../api'

ChartJS.register(CategoryScale, LinearScale, BarElement, LineElement, PointElement, Title, Tooltip, Legend)

const activeTab = ref('list')
const list = ref([])
const total = ref(0)
const page = ref(1)
const pageSize = 20
const loading = ref(false)
const modelFilter = ref('')
const startDate = ref('')
const endDate = ref('')
const duplicateInfo = ref(null)

const totalPages = computed(() => Math.max(1, Math.ceil(total.value / pageSize)))

// Stats
const modelStats = ref([])
const monthlyStats = ref([])
const statsStart = ref('')
const statsEnd = ref('')

const COLORS = ['#1a237e','#283593','#3949ab','#5c6bc0','#7986cb','#9fa8da','#c5cae9','#e8eaf6','#b71c1c','#c62828']

const modelChartData = computed(() => modelStats.value.length ? {
  labels: modelStats.value.slice(0, 12).map(i => i.model_name),
  datasets: [{ label: '건수', data: modelStats.value.slice(0, 12).map(i => i.count), backgroundColor: COLORS, borderRadius: 4 }]
} : null)

const monthlyChartData = computed(() => monthlyStats.value.length ? {
  labels: monthlyStats.value.map(i => i.month),
  datasets: [{ label: '월별', data: monthlyStats.value.map(i => i.count), borderColor: '#1a237e', backgroundColor: 'rgba(26,35,126,0.1)', tension: 0.4, fill: true }]
} : null)

const barOpts = { responsive: true, plugins: { legend: { display: false } } }
const lineOpts = { responsive: true, plugins: { legend: { display: false } } }

async function fetchList() {
  loading.value = true
  try {
    const params = { page: page.value, size: pageSize }
    if (modelFilter.value) params.model_name = modelFilter.value
    if (startDate.value) params.start_date = startDate.value
    if (endDate.value) params.end_date = endDate.value
    const res = await api.listQData(params)
    list.value = res.data.items
    total.value = res.data.total
  } catch (e) { console.error(e) } finally { loading.value = false }
}

function changePage(p) {
  if (p < 1 || p > totalPages.value) return
  page.value = p
  fetchList()
}

async function checkDuplicates() {
  try {
    const res = await api.checkQDataDuplicates()
    duplicateInfo.value = res.data
  } catch (e) { console.error(e) }
}

async function removeDuplicates() {
  if (!confirm('중복 데이터를 제거하시겠습니까?')) return
  try {
    await api.removeQDataDuplicates()
    duplicateInfo.value = null
    fetchList()
  } catch (e) { console.error(e) }
}

function exportData() { api.exportQDataExcel({ start_date: startDate.value, end_date: endDate.value }) }

async function loadStats() {
  try {
    const params = {}
    if (statsStart.value) params.start_date = statsStart.value
    if (statsEnd.value) params.end_date = statsEnd.value
    const [m, mo] = await Promise.all([api.getQDataModelStats(params), api.getQDataMonthlyStats(params)])
    modelStats.value = m.data
    monthlyStats.value = mo.data
  } catch (e) { console.error(e) }
}

function truncate(str, len) { return str && str.length > len ? str.slice(0, len) + '…' : (str || '-') }

onMounted(fetchList)
</script>

<style scoped>
.page-title { font-size: 1.5rem; font-weight: 700; margin-bottom: 20px; color: #1a237e; }
.tabs { display: flex; gap: 8px; margin-bottom: 16px; }
.tab-btn { padding: 8px 18px; border: 1px solid #c5cae9; border-radius: 20px; background: #fff; cursor: pointer; font-size: 0.9rem; }
.tab-btn.active { background: #1a237e; color: #fff; border-color: #1a237e; }
.card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); margin-bottom: 16px; }
.filter-bar { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
.filter-input { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; width: 150px; }
.filter-bar input[type=date] { padding: 7px 10px; border: 1px solid #ddd; border-radius: 6px; }
.btn-primary { padding: 8px 20px; background: #1a237e; color: #fff; border: none; border-radius: 6px; cursor: pointer; }
.btn-secondary { padding: 8px 20px; background: #fff; color: #1a237e; border: 1px solid #1a237e; border-radius: 6px; cursor: pointer; }
.btn-danger { padding: 8px 20px; background: #fff; color: #b71c1c; border: 1px solid #b71c1c; border-radius: 6px; cursor: pointer; }
.dup-card { display: flex; align-items: center; gap: 16px; background: #fff8e1; }
.table-header { margin-bottom: 10px; }
.total-count { font-size: 0.9rem; color: #666; }
.table-wrap { overflow-x: auto; }
.table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
.table th { background: #f8f9ff; padding: 10px 12px; text-align: left; font-weight: 600; color: #555; border-bottom: 2px solid #e8eaf6; white-space: nowrap; }
.table td { padding: 8px 12px; border-bottom: 1px solid #f0f0f0; }
.content-cell { max-width: 200px; }
.badge { background: #e8eaf6; color: #1a237e; padding: 2px 10px; border-radius: 12px; font-weight: 600; }
.empty { text-align: center; color: #aaa; padding: 40px; }
.pagination { display: flex; align-items: center; justify-content: center; gap: 16px; margin-top: 16px; }
.pagination button { padding: 6px 16px; border: 1px solid #c5cae9; border-radius: 6px; background: #fff; cursor: pointer; }
.pagination button:disabled { opacity: 0.4; cursor: not-allowed; }
.loading { text-align: center; padding: 60px; color: #888; }
.stats-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
.section-title { font-size: 1rem; font-weight: 600; margin-bottom: 14px; }
.chart-wrap { max-height: 260px; }
.mt-12 { margin-top: 12px; }
</style>
