<template>
  <div class="launch-view">
    <h1 class="page-title">개통일별 추이 비교 분석</h1>
    <p class="page-desc">모델 출시일을 기준으로 개통일(출시 후 경과일) 단위로 VOC/Q-data 건수를 비교합니다.</p>

    <!-- 모델 선택 -->
    <div class="card">
      <h2 class="section-title">비교 모델 선택</h2>
      <p class="hint">첫 번째 모델(기준 모델)의 현재 개통일 수를 기준으로 비교합니다.</p>

      <div class="model-rows">
        <div v-for="(sel, idx) in selectedModels" :key="idx" class="model-row">
          <span class="model-label" :style="{ color: MODEL_COLORS[idx] }">
            모델 {{ idx === 0 ? 'A (기준)' : String.fromCharCode(65 + idx) }}
          </span>
          <select v-model="selectedModels[idx]" class="model-select">
            <option value="">-- 모델 선택 --</option>
            <option v-for="m in launchModels" :key="m.model_name" :value="m.model_name">
              {{ m.model_name }} (출시일: {{ m.launch_date }})
            </option>
          </select>
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
          <strong>{{ r.model_name }}</strong>
          출시일 {{ r.launch_date }} · 개통 {{ r.max_days }}일차
        </span>
      </div>

      <!-- VOC 차트 -->
      <div class="card">
        <h2 class="section-title">사내 VOC - 개통일별 추이</h2>
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
                    {{ r.model_name }} VOC
                  </th>
                  <th>
                    <span class="legend-dot" :style="{ background: MODEL_COLORS[compareResult.indexOf(r)] }"></span>
                    {{ r.model_name }} Q-data
                  </th>
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

const launchModels = ref([])
const loadingModels = ref(false)
const selectedModels = ref(['', ''])
const compareResult = ref([])
const loading = ref(false)

const canCompare = computed(() =>
  selectedModels.value.filter(m => m).length >= 2
)

async function loadLaunchModels() {
  loadingModels.value = true
  try {
    const res = await getLaunchModels()
    launchModels.value = res.data
  } catch (e) {
    console.error(e)
  } finally {
    loadingModels.value = false
  }
}

function addModel() {
  if (selectedModels.value.length < 5) selectedModels.value.push('')
}

function removeModel(idx) {
  selectedModels.value.splice(idx, 1)
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
      alert(errors.map(e => `${e.model_name}: ${e.error}`).join('\n'))
    }
  } catch (e) {
    console.error(e)
  } finally {
    loading.value = false
  }
}

// 차트 데이터 (VOC)
const vocChartData = computed(() => {
  if (!compareResult.value.length) return null
  const refData = compareResult.value[0].daily_data
  return {
    labels: refData.map(d => `${d.day}일`),
    datasets: compareResult.value.map((r, idx) => ({
      label: r.model_name,
      data: r.daily_data.map(d => d.voc_count),
      borderColor: MODEL_COLORS[idx],
      backgroundColor: MODEL_COLORS_ALPHA[idx],
      tension: 0.3, fill: false, pointRadius: 2,
    })),
  }
})

// 차트 데이터 (Q-data)
const qdataChartData = computed(() => {
  if (!compareResult.value.length) return null
  const refData = compareResult.value[0].daily_data
  return {
    labels: refData.map(d => `${d.day}일`),
    datasets: compareResult.value.map((r, idx) => ({
      label: r.model_name,
      data: r.daily_data.map(d => d.qdata_count),
      borderColor: MODEL_COLORS[idx],
      backgroundColor: MODEL_COLORS_ALPHA[idx],
      tension: 0.3, fill: false, pointRadius: 2,
      borderDash: [5, 3],
    })),
  }
})

// 테이블 데이터
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
.model-select { flex: 1; padding: 8px 10px; border: 1px solid #c5cae9; border-radius: 6px; font-size: 0.9rem; max-width: 400px; }
.btn-remove { padding: 4px 10px; background: #fce4ec; color: #b71c1c; border: none; border-radius: 6px; cursor: pointer; font-size: 0.85rem; }
.btn-row { display: flex; gap: 10px; align-items: center; }
.btn-add { padding: 8px 16px; background: #e8eaf6; color: #1a237e; border: 1px solid #c5cae9; border-radius: 6px; cursor: pointer; font-size: 0.9rem; }
.btn-add:disabled { opacity: 0.5; cursor: not-allowed; }
.btn-primary { padding: 8px 24px; background: #1a237e; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 0.9rem; }
.btn-primary:disabled { opacity: 0.5; cursor: not-allowed; }
.no-data { margin-top: 12px; padding: 16px; background: #fff9c4; border-radius: 6px; font-size: 0.88rem; color: #795548; }
.info-bar { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 16px; }
.info-chip { display: flex; align-items: center; gap: 6px; padding: 6px 14px; background: #fff; border: 2px solid #ccc; border-radius: 20px; font-size: 0.85rem; }
.dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; }
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
