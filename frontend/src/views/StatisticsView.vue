<template>
  <div class="stats-view">
    <h1 class="page-title">통계</h1>

    <!-- 탭 -->
    <div class="tabs">
      <button v-for="tab in tabs" :key="tab.key"
              :class="['tab-btn', { active: activeTab === tab.key }]"
              @click="activeTab = tab.key">
        {{ tab.label }}
      </button>
    </div>

    <!-- 날짜 필터 -->
    <div class="card filter-bar">
      <label>시작일 <input type="date" v-model="startDate" /></label>
      <label>종료일 <input type="date" v-model="endDate" /></label>
      <button class="btn-primary" @click="loadStats">조회</button>
      <button class="btn-secondary" @click="exportData">엑셀 다운로드</button>
    </div>

    <div v-if="loading" class="loading">로딩 중...</div>

    <!-- 모델별: VOC + Q-data 나란히 -->
    <div v-if="activeTab === 'model' && !loading" class="side-by-side">
      <div class="card">
        <h2 class="section-title">Members issue 모델별</h2>
        <div class="chart-wrap">
          <Bar v-if="modelChartData" :data="modelChartData" :options="barOptions" />
        </div>
        <table class="table mt-16">
          <thead><tr><th>순위</th><th>모델명</th><th>건수</th><th>비율</th></tr></thead>
          <tbody>
            <tr v-for="(item, idx) in modelStats" :key="item.model_name">
              <td>{{ idx + 1 }}</td>
              <td class="model-link" @click="openModelDetail(item.model_name)">{{ item.model_name }}</td>
              <td><span class="badge voc">{{ item.count }}</span></td>
              <td>{{ totalModel ? (item.count / totalModel * 100).toFixed(1) : 0 }}%</td>
            </tr>
          </tbody>
        </table>
      </div>
      <div class="card">
        <h2 class="section-title">Q-data 모델별</h2>
        <div class="chart-wrap">
          <Bar v-if="qdataModelChartData" :data="qdataModelChartData" :options="barOptions" />
        </div>
        <table class="table mt-16">
          <thead><tr><th>순위</th><th>모델명</th><th>PPM</th><th>건수</th><th>비율</th></tr></thead>
          <tbody>
            <tr v-for="(item, idx) in qdataModelStats" :key="item.model_name">
              <td>{{ idx + 1 }}</td>
              <td>{{ item.model_name }}</td>
              <td>{{ item.ppm != null ? item.ppm.toFixed(1) : '-' }}</td>
              <td><span class="badge qdata">{{ item.count }}</span></td>
              <td>{{ totalQdataModel ? (item.count / totalQdataModel * 100).toFixed(1) : 0 }}%</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- 월별: 통합 차트 + 모델 선택 -->
    <div v-if="activeTab === 'monthly' && !loading">
      <!-- 모델 선택 시: 개발이슈 현황 카드 -->
      <template v-if="selectedModel">
        <div class="card">
          <div class="section-title">개발 이슈 현황</div>
          <div class="dev-type-cards">
            <div class="dev-type-card dev-card">
              <div class="dev-card-title">개발이슈 (자체)</div>
              <div class="dev-stats-row">
                <div class="dev-stat-item"><span class="dev-stat-label">Total</span><span class="dev-stat-val">{{ selectedModelDevStats?.stats_dev?.total || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Open</span><span class="badge dev-open">{{ selectedModelDevStats?.stats_dev?.open || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Resolve</span><span class="badge dev-resolve">{{ selectedModelDevStats?.stats_dev?.resolve || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Pending</span><span class="badge dev-pending">{{ selectedModelDevStats?.stats_dev?.pending || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Close</span><span class="badge dev-close">{{ selectedModelDevStats?.stats_dev?.close || 0 }}</span></div>
              </div>
            </div>
            <div class="dev-type-card ut-card">
              <div class="dev-card-title">UT 이슈</div>
              <div class="dev-stats-row">
                <div class="dev-stat-item"><span class="dev-stat-label">Total</span><span class="dev-stat-val">{{ selectedModelDevStats?.stats_ut?.total || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Open</span><span class="badge dev-open">{{ selectedModelDevStats?.stats_ut?.open || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Resolve</span><span class="badge dev-resolve">{{ selectedModelDevStats?.stats_ut?.resolve || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Pending</span><span class="badge dev-pending">{{ selectedModelDevStats?.stats_ut?.pending || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Close</span><span class="badge dev-close">{{ selectedModelDevStats?.stats_ut?.close || 0 }}</span></div>
              </div>
            </div>
          </div>
          <div v-if="selectedModelDevIssues.length" class="issue-list">
            <div v-for="iss in selectedModelDevIssues" :key="iss.case_code"
                 class="issue-item" :class="{ 'pending-row': iss.is_pending }">
              <div class="issue-row">
                <span class="badge" :class="'dev-status-' + iss.status">{{ iss.is_pending && iss.status === 'close' ? 'close/Pending' : iss.status }}</span>
                <span class="issue-title">{{ iss.title }}</span>
                <span class="issue-code">{{ iss.case_code }}</span>
              </div>
            </div>
          </div>
        </div>

        <!-- Members issue 현황 카드 -->
        <div class="card">
          <div class="section-title">Members issue 현황</div>
          <div class="dev-type-cards">
            <div class="dev-type-card voc-card">
              <div class="dev-card-title">Members issue (멤버스)</div>
              <div class="dev-stats-row">
                <div class="dev-stat-item"><span class="dev-stat-label">Total</span><span class="dev-stat-val">{{ selectedModelVocStats?.total || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Open</span><span class="badge dev-open">{{ selectedModelVocStats?.open || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Resolve</span><span class="badge dev-resolve">{{ selectedModelVocStats?.resolve || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Pending</span><span class="badge dev-pending">{{ selectedModelVocStats?.pending || 0 }}</span></div>
                <div class="dev-stat-item"><span class="dev-stat-label">Close</span><span class="badge dev-close">{{ selectedModelVocStats?.close || 0 }}</span></div>
              </div>
            </div>
          </div>
          <div v-if="selectedModelVocIssues.length" class="issue-list">
            <div v-for="iss in selectedModelVocIssues" :key="iss.case_code"
                 class="issue-item" :class="{ 'pending-row': iss.is_pending }">
              <div class="issue-row">
                <span class="badge" :class="'voc-status-' + (iss.status || 'open')">{{ iss.is_pending && iss.status === 'close' ? 'close/Pending' : (iss.status || 'open') }}</span>
                <span class="issue-title">{{ iss.summary || iss.title }}</span>
                <span class="issue-code">{{ iss.case_code }}</span>
              </div>
              <div v-if="iss.is_pending && iss.pending_memo" class="pending-detail">
                <span class="pending-memo">📝 {{ iss.pending_memo }}</span>
              </div>
            </div>
          </div>
        </div>
      </template>

      <!-- 월별 추이 차트 + 건수 테이블 -->
      <div class="card">
        <div class="monthly-header">
          <h2 class="section-title">월별 VOC / Q-data 추이</h2>
          <div class="model-filter">
            <label>모델 선택</label>
            <select v-model="selectedModel" @change="loadStats">
              <option value="">전체</option>
              <option v-for="m in allModels" :key="m" :value="m">{{ m }}</option>
            </select>
          </div>
        </div>
        <div class="chart-wrap">
          <Line v-if="combinedMonthlyChartData" :data="combinedMonthlyChartData" :options="lineOptionsLegend" />
        </div>
        <table class="table mt-16">
          <thead>
            <tr>
              <th>월</th>
              <th><span class="legend-dot voc"></span> VOC 건수</th>
              <th><span class="legend-dot qdata"></span> Q-data 건수</th>
              <th>메모</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="item in [...combinedMonthlyData].reverse()" :key="item.month"
                class="clickable-row" @click="goToMonthDetail(item.month)">
              <td>{{ item.month }}</td>
              <td><span class="badge voc">{{ item.voc_count }}</span></td>
              <td><span class="badge qdata">{{ item.qdata_count }}</span></td>
              <td @click.stop>
                <span v-if="!editingMemo[item.month]" class="memo-text"
                      @click="startEdit(item.month, item.memo)">
                  {{ item.memo || '+ 메모 추가' }}
                </span>
                <span v-else class="memo-edit">
                  <input v-model="memoInputs[item.month]"
                         @keyup.enter="saveMemo('monthly', item.month)"
                         @keyup.esc="cancelEdit(item.month)" placeholder="메모 입력" />
                  <button @click="saveMemo('monthly', item.month)">저장</button>
                </span>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- 주별 -->
    <div v-if="activeTab === 'weekly' && !loading" class="card">
      <h2 class="section-title">주별 Members issue / Q-data 추이</h2>
      <div class="chart-wrap">
        <Line v-if="weeklyChartData" :data="weeklyChartData" :options="lineOptionsLegend" />
      </div>
      <table class="table mt-16">
        <thead>
          <tr>
            <th>주차</th>
            <th><span class="legend-dot voc"></span> Members issue</th>
            <th><span class="legend-dot qdata"></span> Q-data</th>
            <th>메모</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="item in [...combinedWeeklyData].reverse()" :key="item.week">
            <td>{{ item.week }}</td>
            <td><span class="badge voc">{{ item.voc_count }}</span></td>
            <td><span class="badge qdata">{{ item.qdata_count }}</span></td>
            <td>
              <span v-if="!editingMemo[item.week]" class="memo-text" @click="startEdit(item.week, item.memo)">
                {{ item.memo || '+ 메모 추가' }}
              </span>
              <span v-else class="memo-edit">
                <input v-model="memoInputs[item.week]" @keyup.enter="saveMemo('weekly', item.week)"
                       @keyup.esc="cancelEdit(item.week)" placeholder="메모 입력" />
                <button @click="saveMemo('weekly', item.week)">저장</button>
              </span>
            </td>
          </tr>
        </tbody>
      </table>
    </div>

    <!-- 칩셋별 -->
    <div v-if="activeTab === 'chipset' && !loading" class="card">
      <h2 class="section-title">칩셋별 Members issue / Q-data</h2>
      <div class="chart-wrap">
        <Bar v-if="chipsetChartData" :data="chipsetChartData" :options="chipsetBarOptions" />
      </div>
      <table class="table mt-16">
        <thead>
          <tr>
            <th>순위</th>
            <th>칩셋</th>
            <th><span class="legend-dot voc"></span> Members issue</th>
            <th><span class="legend-dot qdata"></span> Q-data</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(item, idx) in chipsetStats" :key="item.chipset">
            <td>{{ idx + 1 }}</td>
            <td>{{ item.chipset }}</td>
            <td><span class="badge voc">{{ item.voc_count }}</span></td>
            <td><span class="badge qdata">{{ item.qdata_count }}</span></td>
          </tr>
        </tbody>
      </table>
    </div>

    <!-- 앱별 -->
    <div v-if="activeTab === 'app' && !loading" class="card">
      <h2 class="section-title">3rd Party 앱별 VOC 건수</h2>
      <div class="chart-wrap">
        <Bar v-if="appChartData" :data="appChartData" :options="barOptions" />
      </div>
      <table class="table mt-16">
        <thead><tr><th>순위</th><th>앱명</th><th>건수</th></tr></thead>
        <tbody>
          <tr v-for="(item, idx) in appStats" :key="item.app_name">
            <td>{{ idx + 1 }}</td>
            <td>{{ item.app_name }}</td>
            <td><span class="badge voc">{{ item.count }}</span></td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, watch, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { Bar, Line, Doughnut } from 'vue-chartjs'
import {
  Chart as ChartJS, CategoryScale, LinearScale, BarElement,
  LineElement, PointElement, ArcElement, Title, Tooltip, Legend,
} from 'chart.js'
import * as api from '../api'

ChartJS.register(CategoryScale, LinearScale, BarElement, LineElement, PointElement, ArcElement, Title, Tooltip, Legend)

const router = useRouter()
function goToMonthDetail(month) {
  window.open(`/statistics/month/${month}`, '_blank')
}
function openModelDetail(name) {
  window.open('/statistics/model/' + encodeURIComponent(name), '_blank')
}

const tabs = [
  { key: 'model', label: '모델별' },
  { key: 'monthly', label: '월별' },
  { key: 'weekly', label: '주별' },
  { key: 'chipset', label: '칩셋별' },
  { key: 'app', label: '앱별' },
]
const activeTab = ref('model')
const startDate = ref('')
const endDate = ref('')
const loading = ref(false)

// VOC 데이터
const modelStats = ref([])
const monthlyStats = ref([])
const weeklyStats = ref([])
const chipsetStats = ref([])
const appStats = ref([])

// Q-data 데이터
const qdataModelStats = ref([])
const qdataMonthlyStats = ref([])
const qdataWeeklyStats = ref([])

// 월별 모델 선택
const selectedModel = ref('')
const allModels = ref([])
const selectedModelDevStats = ref(null)
const selectedModelDevIssues = ref([])
const selectedModelVocStats = ref(null)
const selectedModelVocIssues = ref([])

const editingMemo = ref({})
const memoInputs = ref({})

const totalModel = computed(() => modelStats.value.reduce((s, i) => s + i.count, 0))
const totalQdataModel = computed(() => qdataModelStats.value.reduce((s, i) => s + i.count, 0))

const VOC_COLORS = ['#1a237e','#283593','#3949ab','#5c6bc0','#7986cb','#9fa8da','#c5cae9']
const QDATA_COLORS = ['#bf360c','#d84315','#e64a19','#f4511e','#ff5722','#ff7043','#ff8a65']

const modelChartData = computed(() => modelStats.value.length ? {
  labels: modelStats.value.slice(0, 15).map(i => i.model_name),
  datasets: [{ label: 'VOC 건수', data: modelStats.value.slice(0, 15).map(i => i.count),
    backgroundColor: VOC_COLORS, borderRadius: 4 }],
} : null)

const qdataModelChartData = computed(() => qdataModelStats.value.length ? {
  labels: qdataModelStats.value.slice(0, 15).map(i => i.model_name),
  datasets: [{ label: 'Q-data 건수', data: qdataModelStats.value.slice(0, 15).map(i => i.count),
    backgroundColor: QDATA_COLORS, borderRadius: 4 }],
} : null)

// 월별 통합 데이터 (VOC + Q-data 병합)
const combinedMonthlyData = computed(() => {
  const vocByMonth = Object.fromEntries(monthlyStats.value.map(i => [i.month, { count: i.count, memo: i.memo || '' }]))
  const qdataByMonth = Object.fromEntries(qdataMonthlyStats.value.map(i => [i.month, i.count]))
  const allMonths = [...new Set([...Object.keys(vocByMonth), ...Object.keys(qdataByMonth)])].sort()
  return allMonths.map(month => ({
    month,
    voc_count: vocByMonth[month]?.count || 0,
    qdata_count: qdataByMonth[month] || 0,
    memo: vocByMonth[month]?.memo || '',
  }))
})

const combinedMonthlyChartData = computed(() => {
  const data = combinedMonthlyData.value
  if (!data.length) return null
  return {
    labels: data.map(i => i.month),
    datasets: [
      {
        label: 'Members issue',
        data: data.map(i => i.voc_count),
        borderColor: '#1a237e', backgroundColor: 'rgba(26,35,126,0.1)',
        tension: 0.4, fill: false, pointRadius: 4,
      },
      {
        label: 'Q-data',
        data: data.map(i => i.qdata_count),
        borderColor: '#e64a19', backgroundColor: 'rgba(230,74,25,0.1)',
        tension: 0.4, fill: false, pointRadius: 4,
      },
    ],
  }
})

const combinedWeeklyData = computed(() => {
  const vocByWeek = Object.fromEntries(weeklyStats.value.map(i => [i.week, { count: i.count, memo: i.memo || '' }]))
  const qdataByWeek = Object.fromEntries(qdataWeeklyStats.value.map(i => [i.week, i.count]))
  const allWeeks = [...new Set([...Object.keys(vocByWeek), ...Object.keys(qdataByWeek)])].sort()
  return allWeeks.map(week => ({
    week,
    voc_count: vocByWeek[week]?.count || 0,
    qdata_count: qdataByWeek[week] || 0,
    memo: vocByWeek[week]?.memo || '',
  }))
})

const weeklyChartData = computed(() => combinedWeeklyData.value.length ? {
  labels: combinedWeeklyData.value.map(i => i.week),
  datasets: [
    { label: 'Members issue', data: combinedWeeklyData.value.map(i => i.voc_count),
      borderColor: '#1a237e', backgroundColor: 'rgba(26,35,126,0.1)', tension: 0.4, fill: false, pointRadius: 4 },
    { label: 'Q-data', data: combinedWeeklyData.value.map(i => i.qdata_count),
      borderColor: '#e64a19', backgroundColor: 'rgba(230,74,25,0.1)', tension: 0.4, fill: false, pointRadius: 4 },
  ],
} : null)

const chipsetChartData = computed(() => chipsetStats.value.length ? {
  labels: chipsetStats.value.slice(0, 10).map(i => i.chipset),
  datasets: [
    { label: 'Members issue', data: chipsetStats.value.slice(0, 10).map(i => i.voc_count),
      backgroundColor: 'rgba(26,35,126,0.7)', borderRadius: 4 },
    { label: 'Q-data', data: chipsetStats.value.slice(0, 10).map(i => i.qdata_count),
      backgroundColor: 'rgba(230,74,25,0.7)', borderRadius: 4 },
  ],
} : null)

const appChartData = computed(() => appStats.value.length ? {
  labels: appStats.value.map(i => i.app_name),
  datasets: [{ label: '건수', data: appStats.value.map(i => i.count),
    backgroundColor: VOC_COLORS, borderRadius: 4 }],
} : null)

const barOptions = { responsive: true, plugins: { legend: { display: false } } }
const lineOptions = { responsive: true, plugins: { legend: { display: false } } }
const chipsetBarOptions = { responsive: true, plugins: { legend: { display: true, position: 'top' } } }
const lineOptionsLegend = computed(() => ({
  responsive: true,
  plugins: {
    legend: { display: true, position: 'top' },
    tooltip: {
      callbacks: {
        footer(items) {
          const idx = items[0]?.dataIndex
          const memo = combinedMonthlyData.value[idx]?.memo
          return memo ? `📝 ${memo}` : ''
        },
      },
    },
  },
}))
const doughnutOptions = { responsive: true, plugins: { legend: { position: 'right' } } }

async function loadAllModels() {
  try {
    const res = await api.getEffectiveModels()
    allModels.value = res.data
  } catch (e) {
    console.error(e)
  }
}

async function loadStats() {
  loading.value = true
  const params = {}
  if (startDate.value) params.start_date = startDate.value
  if (endDate.value) params.end_date = endDate.value

  try {
    const tab = activeTab.value

    if (tab === 'model') {
      const [vocRes, qdataRes] = await Promise.all([
        api.getModelStats(params),
        api.getQDataModelStats(params),
      ])
      modelStats.value = vocRes.data
      qdataModelStats.value = qdataRes.data

    } else if (tab === 'monthly') {
      if (!allModels.value.length) await loadAllModels()

      if (selectedModel.value) {
        const [statsRes, devRes, vocRes] = await Promise.allSettled([
          api.getEffectiveNameMonthly(selectedModel.value, params),
          api.getModelDevIssues(selectedModel.value),
          api.getVocModelActive(selectedModel.value),
        ])
        if (statsRes.status === 'fulfilled') {
          monthlyStats.value = statsRes.value.data.voc_monthly || []
          qdataMonthlyStats.value = statsRes.value.data.qdata_monthly || []
        }
        if (devRes.status === 'fulfilled') {
          selectedModelDevStats.value = devRes.value.data
          selectedModelDevIssues.value = devRes.value.data.issues || []
        }
        if (vocRes.status === 'fulfilled') {
          selectedModelVocStats.value = vocRes.value.data.stats || null
          selectedModelVocIssues.value = vocRes.value.data.issues || []
        }
      } else {
        selectedModelDevStats.value = null
        selectedModelDevIssues.value = []
        selectedModelVocStats.value = null
        selectedModelVocIssues.value = []
        const [vocRes, qdataRes] = await Promise.all([
          api.getMonthlyStats(params),
          api.getQDataMonthlyStats(params),
        ])
        monthlyStats.value = vocRes.data
        qdataMonthlyStats.value = qdataRes.data
      }

    } else if (tab === 'weekly') {
      const [vocResult, qdataResult] = await Promise.allSettled([
        api.getWeeklyStats(params),
        api.getQDataWeeklyStats(params),
      ])
      if (vocResult.status === 'fulfilled') weeklyStats.value = vocResult.value.data
      if (qdataResult.status === 'fulfilled') qdataWeeklyStats.value = qdataResult.value.data
    } else if (tab === 'chipset') {
      chipsetStats.value = (await api.getChipsetStats(params)).data
    } else if (tab === 'app') {
      appStats.value = (await api.getAppStats(params)).data
    }
  } catch (e) {
    console.error(e)
  } finally {
    loading.value = false
  }
}

watch(activeTab, () => loadStats())
onMounted(() => loadStats())

function exportData() {
  api.exportVocExcel({ start_date: startDate.value, end_date: endDate.value })
}

function startEdit(key, current) {
  editingMemo.value[key] = true
  memoInputs.value[key] = current || ''
}
function cancelEdit(key) {
  editingMemo.value[key] = false
}
async function saveMemo(type, key) {
  const memo = memoInputs.value[key]
  try {
    if (type === 'monthly') await api.saveMonthlyMemo(key, memo)
    else if (type === 'weekly') await api.saveWeeklyMemo(key, memo)
    editingMemo.value[key] = false
    await loadStats()
  } catch (e) { console.error(e) }
}
</script>

<style scoped>
.page-title { font-size: 1.5rem; font-weight: 700; margin-bottom: 20px; color: #1a237e; }
.tabs { display: flex; gap: 8px; margin-bottom: 16px; flex-wrap: wrap; }
.tab-btn { padding: 8px 18px; border: 1px solid #c5cae9; border-radius: 20px; background: #fff; cursor: pointer; font-size: 0.9rem; transition: all 0.2s; }
.tab-btn.active { background: #1a237e; color: #fff; border-color: #1a237e; }
.card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); margin-bottom: 16px; }
.side-by-side { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px; }
@media (max-width: 900px) { .side-by-side { grid-template-columns: 1fr; } }
.filter-bar { display: flex; align-items: center; gap: 16px; flex-wrap: wrap; }
.filter-bar label { display: flex; align-items: center; gap: 8px; font-size: 0.9rem; }
.filter-bar input[type=date] { padding: 6px 10px; border: 1px solid #ddd; border-radius: 6px; }
.btn-primary { padding: 8px 20px; background: #1a237e; color: #fff; border: none; border-radius: 6px; cursor: pointer; }
.btn-secondary { padding: 8px 20px; background: #fff; color: #1a237e; border: 1px solid #1a237e; border-radius: 6px; cursor: pointer; }
.section-title { font-size: 1rem; font-weight: 600; margin-bottom: 14px; color: #333; }
.chart-wrap { max-height: 450px; display: flex; justify-content: center; }
.table { width: 100%; border-collapse: collapse; }
.table th, .table td { padding: 10px 14px; text-align: left; border-bottom: 1px solid #eee; font-size: 0.88rem; }
.table th { background: #f8f9ff; font-weight: 600; color: #555; }
.badge { padding: 2px 10px; border-radius: 12px; font-weight: 600; font-size: 0.85rem; }
.badge.voc { background: #e8eaf6; color: #1a237e; }
.badge.qdata { background: #fbe9e7; color: #bf360c; }
.mt-16 { margin-top: 16px; }
.loading { text-align: center; padding: 40px; color: #888; }
.memo-text { cursor: pointer; color: #888; font-style: italic; }
.memo-text:hover { color: #1a237e; text-decoration: underline; }
.memo-edit { display: flex; gap: 6px; align-items: center; }
.memo-edit input { padding: 4px 8px; border: 1px solid #c5cae9; border-radius: 4px; font-size: 0.85rem; }
.memo-edit button { padding: 4px 10px; background: #1a237e; color: #fff; border: none; border-radius: 4px; cursor: pointer; font-size: 0.8rem; }
.clickable-row { cursor: pointer; }
.clickable-row:hover { background: #f0f2ff; }
.monthly-header { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; margin-bottom: 14px; }
.monthly-header .section-title { margin-bottom: 0; }
.model-filter { display: flex; align-items: center; gap: 8px; font-size: 0.9rem; }
.model-filter select { padding: 6px 10px; border: 1px solid #c5cae9; border-radius: 6px; font-size: 0.9rem; min-width: 160px; }
.legend-dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 4px; vertical-align: middle; }
.legend-dot.voc { background: #1a237e; }
.legend-dot.qdata { background: #e64a19; }

/* 모델 링크 */
.model-link { cursor: pointer; color: #1a237e; }
.model-link:hover { text-decoration: underline; }

/* 개발/Members이슈 현황 카드 */
.dev-type-cards { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 4px; }
@media (max-width: 600px) { .dev-type-cards { grid-template-columns: 1fr; } }
.dev-type-card { border-radius: 8px; padding: 14px 16px; }
.dev-card { background: #f1f8e9; border-left: 4px solid #43a047; }
.ut-card { background: #fce4ec; border-left: 4px solid #e53935; }
.voc-card { background: #e8eaf6; border-left: 4px solid #3f51b5; }
.dev-card-title { font-size: 0.88rem; font-weight: 700; margin-bottom: 10px; color: #333; }
.dev-stats-row { display: flex; gap: 14px; flex-wrap: wrap; }
.dev-stat-item { display: flex; flex-direction: column; align-items: center; gap: 4px; }
.dev-stat-label { font-size: 0.75rem; color: #888; font-weight: 500; }
.dev-stat-val { font-size: 1.1rem; font-weight: 700; color: #333; }
.badge.dev-open { background: #e3f2fd; color: #1565c0; padding: 3px 10px; border-radius: 10px; font-weight: 700; font-size: 0.82rem; }
.badge.dev-resolve { background: #e8f5e9; color: #2e7d32; padding: 3px 10px; border-radius: 10px; font-weight: 700; font-size: 0.82rem; }
.badge.dev-pending { background: #fff3e0; color: #e65100; padding: 3px 10px; border-radius: 10px; font-weight: 700; font-size: 0.82rem; }
.badge.dev-close { background: #f5f5f5; color: #757575; padding: 3px 10px; border-radius: 10px; font-weight: 700; font-size: 0.82rem; }
.badge.dev-status-open { background: #e3f2fd; color: #1565c0; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.dev-status-resolve { background: #e8f5e9; color: #2e7d32; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.dev-status-close { background: #f5f5f5; color: #757575; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.voc-status-open { background: #e3f2fd; color: #1565c0; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.voc-status-resolve { background: #e8f5e9; color: #2e7d32; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.voc-status-close { background: #f5f5f5; color: #757575; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }

/* 이슈 리스트 */
.issue-list { display: flex; flex-direction: column; gap: 6px; margin-top: 12px; }
.issue-item { border-radius: 7px; background: #fafafa; border-left: 3px solid #e0e0e0; overflow: hidden; }
.issue-item.pending-row { background: #fffde7; border-left-color: #ffa000; }
.issue-row { display: flex; align-items: center; gap: 10px; padding: 8px 12px; }
.issue-title { flex: 1; font-size: 0.85rem; color: #333; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.issue-code { font-size: 0.78rem; color: #999; white-space: nowrap; font-family: monospace; }
.pending-detail { padding: 4px 12px 8px 12px; }
.pending-memo { font-size: 0.82rem; color: #795548; }
</style>
