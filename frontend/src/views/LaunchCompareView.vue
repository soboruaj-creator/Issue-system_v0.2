<template>
  <div class="launch-view">
    <h1 class="page-title">모델별 비교 분석</h1>
    <p class="page-desc">모델 출시일을 기준으로 개통일(출시 후 경과일) 단위로 Members issue/Q-data 건수를 비교합니다.</p>

    <!-- 모델 선택 -->
    <div class="card">
      <h2 class="section-title">비교 모델 선택</h2>
      <p class="hint">첫 번째 모델(기준 모델)의 현재 개통일 수를 기준으로 비교합니다. 마케팅명으로 묶인 모델은 건수가 합산됩니다.</p>

      <div class="model-rows">
        <div v-for="(sel, idx) in selectedModels" :key="idx" class="model-row">
          <span class="model-label" :style="{ color: MODEL_COLORS[idx] }">
            모델 {{ idx === 0 ? 'A (기준)' : String.fromCharCode(65 + idx) }}
          </span>
          <div class="search-wrap">
            <input
              type="text"
              class="model-input"
              :value="searchQueries[idx] !== null ? searchQueries[idx] : sel"
              @input="onInput(idx, $event.target.value)"
              @focus="openDropdown(idx)"
              @blur="onBlur(idx)"
              placeholder="마케팅명 또는 모델명 검색..."
              autocomplete="off"
            />
            <div v-if="openIdx === idx && filteredModels(idx).length" class="dropdown">
              <div
                v-for="m in filteredModels(idx)"
                :key="m.effective_name"
                class="dropdown-item"
                @mousedown.prevent="selectModel(idx, m)"
              >
                <div class="d-row">
                  <span class="d-name">{{ m.effective_name }}</span>
                  <span class="d-date">출시 {{ m.launch_date }}</span>
                </div>
                <div v-if="m.model_names.length > 1" class="d-subs">
                  {{ m.model_names.join(', ') }}
                </div>
              </div>
            </div>
          </div>
          <button v-if="idx > 1" class="btn-remove" @click="removeModel(idx)">✕</button>
        </div>
      </div>

      <div class="btn-row">
        <button class="btn-add" @click="addModel" :disabled="selectedModels.length >= 5">
          + 모델 추가
        </button>
        <button class="btn-primary" @click="compare" :disabled="!canCompare || loading">
          {{ loading ? '분석 중...' : '비교 분석' }}
        </button>
      </div>

      <div v-if="launchModels.length === 0 && !loadingModels" class="no-data">
        출시일이 등록된 모델이 없습니다. 업로드 메뉴에서 출시일을 먼저 등록해주세요.
      </div>
    </div>

    <!-- 결과 -->
    <div v-if="compareResult.length && !loading">

      <!-- 기준 모델 정보 -->
      <div class="info-bar">
        <span v-for="r in compareResult" :key="r.model_name" class="info-chip"
              :style="{ borderColor: MODEL_COLORS[compareResult.indexOf(r)] }">
          <span class="dot" :style="{ background: MODEL_COLORS[compareResult.indexOf(r)] }"></span>
          <strong>{{ r.display_name }}</strong>
          <span v-if="r.model_names && r.model_names.length > 1" class="sub-models">({{ r.model_names.join(', ') }})</span>
          출시일 {{ r.launch_date }} · 개통 {{ r.max_days }}일차
          <span class="total-counts">
            <span class="tc voc">Members issue {{ totalVoc(r) }}건</span>
            <span class="tc qdata">Q-data {{ totalQdata(r) }}건</span>
            <span v-if="r.latest_ppm != null" class="tc ppm">PPM {{ r.latest_ppm.toFixed(1) }}</span>
          </span>
        </span>
      </div>

      <!-- Members issue 차트 -->
      <div class="card">
        <h2 class="section-title">Members issue - 개통일별 추이</h2>
        <div class="chart-wrap-lg">
          <Line v-if="vocChartData" :data="vocChartData" :options="lineOptions" />
        </div>
      </div>

      <!-- Q-data 차트 -->
      <div class="card">
        <h2 class="section-title">Q-data - 개통일별 추이</h2>
        <div class="chart-wrap-lg">
          <Line v-if="qdataChartData" :data="qdataChartData" :options="lineOptions" />
        </div>
      </div>

      <!-- 데이터 테이블 -->
      <div class="card">
        <h2 class="section-title">상세 데이터</h2>
        <div class="table-scroll">
          <table class="table">
            <thead>
              <tr>
                <th>개통일</th>
                <th>날짜 (기준)</th>
                <template v-for="r in compareResult" :key="r.model_name">
                  <th>
                    <span class="legend-dot" :style="{ background: MODEL_COLORS[compareResult.indexOf(r)] }"></span>
                    {{ r.display_name }} VOC
                  </th>
                  <th>
                    <span class="legend-dot" :style="{ background: MODEL_COLORS[compareResult.indexOf(r)] }"></span>
                    {{ r.display_name }} Q-data
                  </th>
                  <th>{{ r.display_name }} PPM</th>
                </template>
              </tr>
            </thead>
            <tbody>
              <tr v-for="row in tableData" :key="row.day">
                <td>{{ row.day }}일</td>
                <td class="date-col">{{ row.ref_date }}</td>
                <template v-for="r in compareResult" :key="r.model_name">
                  <td>
                    <span v-if="row.data[r.model_name]" class="badge voc">
                      {{ row.data[r.model_name].voc_count }}
                    </span>
                    <span v-else class="na">-</span>
                  </td>
                  <td>
                    <span v-if="row.data[r.model_name]" class="badge qdata">
                      {{ row.data[r.model_name].qdata_count }}
                    </span>
                    <span v-else class="na">-</span>
                  </td>
                  <td>
                    <span v-if="row.data[r.model_name]?.ppm != null" class="badge ppm">
                      {{ row.data[r.model_name].ppm.toFixed(1) }}
                    </span>
                    <span v-else class="na">-</span>
                  </td>
                </template>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <div v-if="loading" class="loading">분석 중...</div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'
import { Line } from 'vue-chartjs'
import {
  Chart as ChartJS, CategoryScale, LinearScale, LineElement,
  PointElement, Title, Tooltip, Legend,
} from 'chart.js'
import { getLaunchModels, compareLaunchModels } from '../api'

ChartJS.register(CategoryScale, LinearScale, LineElement, PointElement, Title, Tooltip, Legend)

const MODEL_COLORS = ['#1a237e', '#1b5e20', '#b71c1c', '#f57f17', '#4a148c']
const MODEL_COLORS_ALPHA = ['rgba(26,35,126,0.15)', 'rgba(27,94,32,0.15)', 'rgba(183,28,28,0.15)', 'rgba(245,127,23,0.15)', 'rgba(74,20,140,0.15)']

// launchModels: [{effective_name, model_names[], launch_date}]
const launchModels = ref([])
const loadingModels = ref(false)
// selectedModels stores effective_name
const selectedModels = ref(['', ''])
// searchQueries[idx]: null = show selected value, string = active search
const searchQueries = ref([null, null])
const openIdx = ref(null)
const compareResult = ref([])
const loading = ref(false)

const canCompare = computed(() =>
  selectedModels.value.filter(m => m).length >= 2
)

function filteredModels(idx) {
  const q = (searchQueries.value[idx] || '').toLowerCase()
  if (!q) return launchModels.value
  return launchModels.value.filter(m =>
    m.effective_name.toLowerCase().includes(q) ||
    m.model_names.some(n => n.toLowerCase().includes(q))
  )
}

function onInput(idx, val) {
  searchQueries.value[idx] = val
  selectedModels.value[idx] = ''
  openIdx.value = idx
}

function openDropdown(idx) {
  searchQueries.value[idx] = ''
  openIdx.value = idx
}

function onBlur(idx) {
  setTimeout(() => {
    if (openIdx.value === idx) openIdx.value = null
    if (!selectedModels.value[idx]) searchQueries.value[idx] = null
    else searchQueries.value[idx] = null  // show selected value
  }, 200)
}

function selectModel(idx, m) {
  selectedModels.value[idx] = m.effective_name
  searchQueries.value[idx] = null
  openIdx.value = null
}

async function loadLaunchModels() {
  loadingModels.value = true
  try {
    const res = await getLaunchModels()
    launchModels.value = res.data  // [{effective_name, model_names, launch_date}]
  } catch (e) {
    console.error(e)
  } finally {
    loadingModels.value = false
  }
}

function addModel() {
  if (selectedModels.value.length < 5) {
    selectedModels.value.push('')
    searchQueries.value.push(null)
  }
}

function removeModel(idx) {
  selectedModels.value.splice(idx, 1)
  searchQueries.value.splice(idx, 1)
}

function totalVoc(r) {
  return r.daily_data.reduce((s, d) => s + d.voc_count, 0).toLocaleString()
}

function totalQdata(r) {
  return r.daily_data.reduce((s, d) => s + d.qdata_count, 0).toLocaleString()
}

async function compare() {
  const models = selectedModels.value.filter(m => m)
  if (models.length < 2) return
  loading.value = true
  compareResult.value = []
  try {
    const res = await compareLaunchModels(models)
    compareResult.value = res.data.filter(r => !r.error)
    const errors = res.data.filter(r => r.error)
    if (errors.length) {
      alert(errors.map(e => `${e.display_name}: ${e.error}`).join('\n'))
    }
  } catch (e) {
    console.error(e)
  } finally {
    loading.value = false
  }
}

const vocChartData = computed(() => {
  if (!compareResult.value.length) return null
  const refData = compareResult.value[0].daily_data
  return {
    labels: refData.map(d => `${d.day}일`),
    datasets: compareResult.value.map((r, idx) => ({
      label: r.display_name,
      data: r.daily_data.map(d => d.voc_count),
      borderColor: MODEL_COLORS[idx],
      backgroundColor: MODEL_COLORS_ALPHA[idx],
      tension: 0.3, fill: false, pointRadius: 2,
    })),
  }
})

const qdataChartData = computed(() => {
  if (!compareResult.value.length) return null
  const refData = compareResult.value[0].daily_data
  return {
    labels: refData.map(d => `${d.day}일`),
    datasets: compareResult.value.map((r, idx) => ({
      label: r.display_name,
      data: r.daily_data.map(d => d.qdata_count),
      borderColor: MODEL_COLORS[idx],
      backgroundColor: MODEL_COLORS_ALPHA[idx],
      tension: 0.3, fill: false, pointRadius: 2,
      borderDash: [5, 3],
    })),
  }
})

const tableData = computed(() => {
  if (!compareResult.value.length) return []
  const refResult = compareResult.value[0]
  return refResult.daily_data.map(d => ({
    day: d.day,
    ref_date: d.date,
    data: Object.fromEntries(
      compareResult.value.map(r => [
        r.model_name,
        r.daily_data.find(dd => dd.day === d.day) || null,
      ])
    ),
  }))
})

const lineOptions = {
  responsive: true,
  plugins: { legend: { display: true, position: 'top' } },
  scales: { x: { ticks: { maxTicksLimit: 30 } } },
}

onMounted(() => loadLaunchModels())
</script>

<style scoped>
.page-title { font-size: 1.5rem; font-weight: 700; margin-bottom: 6px; color: #1a237e; }
.page-desc { font-size: 0.9rem; color: #888; margin-bottom: 20px; }
.card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); margin-bottom: 16px; }
.section-title { font-size: 1rem; font-weight: 600; margin-bottom: 14px; color: #333; }
.hint { font-size: 0.85rem; color: #888; margin-bottom: 14px; }
.model-rows { display: flex; flex-direction: column; gap: 10px; margin-bottom: 16px; }
.model-row { display: flex; align-items: center; gap: 10px; }
.model-label { font-size: 0.85rem; font-weight: 600; min-width: 90px; }
.search-wrap { position: relative; flex: 1; max-width: 420px; }
.model-input { width: 100%; padding: 8px 10px; border: 1px solid #c5cae9; border-radius: 6px; font-size: 0.9rem; box-sizing: border-box; }
.model-input:focus { outline: none; border-color: #3f51b5; box-shadow: 0 0 0 2px rgba(63,81,181,0.15); }
.dropdown { position: absolute; top: calc(100% + 4px); left: 0; right: 0; background: #fff; border: 1px solid #c5cae9; border-radius: 6px; box-shadow: 0 4px 12px rgba(0,0,0,0.12); z-index: 100; max-height: 260px; overflow-y: auto; }
.dropdown-item { padding: 8px 12px; cursor: pointer; }
.dropdown-item:hover { background: #e8eaf6; }
.d-row { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
.d-name { font-weight: 600; color: #1a237e; font-size: 0.9rem; }
.d-date { color: #aaa; font-size: 0.8rem; white-space: nowrap; }
.d-subs { font-size: 0.78rem; color: #888; margin-top: 2px; }
.btn-remove { padding: 4px 10px; background: #fce4ec; color: #b71c1c; border: none; border-radius: 6px; cursor: pointer; font-size: 0.85rem; }
.btn-row { display: flex; gap: 10px; align-items: center; }
.btn-add { padding: 8px 16px; background: #e8eaf6; color: #1a237e; border: 1px solid #c5cae9; border-radius: 6px; cursor: pointer; font-size: 0.9rem; }
.btn-add:disabled { opacity: 0.5; cursor: not-allowed; }
.btn-primary { padding: 8px 24px; background: #1a237e; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 0.9rem; }
.btn-primary:disabled { opacity: 0.5; cursor: not-allowed; }
.no-data { margin-top: 12px; padding: 16px; background: #fff9c4; border-radius: 6px; font-size: 0.88rem; color: #795548; }
.info-bar { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 16px; }
.info-chip { display: flex; align-items: center; gap: 6px; padding: 6px 14px; background: #fff; border: 2px solid #ccc; border-radius: 20px; font-size: 0.85rem; flex-wrap: wrap; }
.sub-models { font-size: 0.78rem; color: #888; }
.total-counts { display: flex; gap: 6px; margin-left: 4px; }
.tc { font-size: 0.78rem; padding: 2px 8px; border-radius: 10px; font-weight: 600; }
.tc.voc { background: #e8eaf6; color: #1a237e; }
.tc.qdata { background: #fff3e0; color: #e65100; }
.tc.ppm { background: #e8f5e9; color: #2e7d32; }
.badge.ppm { background: #e8f5e9; color: #2e7d32; }
.dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
.chart-wrap-lg { height: 300px; }
.table-scroll { overflow-x: auto; }
.table { width: 100%; border-collapse: collapse; min-width: 600px; }
.table th, .table td { padding: 8px 12px; text-align: center; border-bottom: 1px solid #eee; font-size: 0.85rem; white-space: nowrap; }
.table th { background: #f8f9ff; font-weight: 600; color: #555; }
.date-col { color: #888; font-size: 0.8rem; }
.badge { padding: 2px 8px; border-radius: 10px; font-weight: 600; font-size: 0.82rem; }
.badge.voc { background: #e8eaf6; color: #1a237e; }
.badge.qdata { background: #fbe9e7; color: #bf360c; }
.na { color: #ccc; }
.loading { text-align: center; padding: 40px; color: #888; }
.legend-dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 4px; vertical-align: middle; }
</style>
