<template>
  <div class="model-detail-view">
    <!-- 헤더 -->
    <div class="header-row">
      <button class="btn-close" @click="window.close()">✕ 닫기</button>
      <h1 class="page-title">{{ modelName }}</h1>
      <span></span>
    </div>

    <div v-if="loading" class="loading">불러오는 중...</div>

    <template v-else>
      <!-- 월별 추이 차트 -->
      <div class="card">
        <div class="section-title">월별 인입 추이 (Members issue + Q-data)</div>
        <div class="chart-wrap">
          <Line v-if="chartData" :data="chartData" :options="chartOptions" />
          <div v-else class="no-chart">데이터 없음</div>
        </div>
      </div>

      <!-- 월별 테이블 -->
      <div class="card">
        <div class="section-title">월별 상세</div>
        <div class="table-scroll">
          <table class="table">
            <thead>
              <tr>
                <th>월</th>
                <th><span class="dot voc-dot"></span> Members issue</th>
                <th><span class="dot qdata-dot"></span> Q-data</th>
                <th>합계</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="row in monthlyRows" :key="row.month">
                <td>{{ row.month }}</td>
                <td><span class="badge voc">{{ row.voc }}</span></td>
                <td><span class="badge qdata">{{ row.qdata }}</span></td>
                <td><strong>{{ row.voc + row.qdata }}</strong></td>
              </tr>
              <tr v-if="!monthlyRows.length">
                <td colspan="4" class="empty">데이터 없음</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- 모델 이벤트 로그 -->
      <div class="card">
        <div class="log-header">
          <div class="section-title">모델 이벤트 로그</div>
          <button class="btn-add-note" @click="startAdd">+ 추가</button>
        </div>

        <!-- 추가 폼 -->
        <div v-if="addMode" class="note-form">
          <input type="date" v-model="form.date" class="input-date" />
          <input type="text" v-model="form.content" class="input-content" placeholder="내용 입력 (예: OTA v2.0 배포, 배터리 이슈 접수 시작 등)" @keyup.enter="submitAdd" />
          <button class="btn-save" @click="submitAdd" :disabled="saving">저장</button>
          <button class="btn-cancel" @click="cancelAdd">취소</button>
        </div>

        <!-- 로그 목록 -->
        <div v-if="notes.length" class="notes-list">
          <div v-for="note in notes" :key="note.id" class="note-item">
            <template v-if="editingId !== note.id">
              <div class="note-date">{{ note.date }}</div>
              <div class="note-content">{{ note.content }}</div>
              <div class="note-actions">
                <button class="btn-edit" @click="startEdit(note)">수정</button>
                <button class="btn-delete" @click="deleteNote(note.id)">삭제</button>
              </div>
            </template>
            <template v-else>
              <input type="date" v-model="editForm.date" class="input-date" />
              <input type="text" v-model="editForm.content" class="input-content" @keyup.enter="submitEdit(note.id)" />
              <div class="note-actions">
                <button class="btn-save" @click="submitEdit(note.id)" :disabled="saving">저장</button>
                <button class="btn-cancel" @click="cancelEdit">취소</button>
              </div>
            </template>
          </div>
        </div>
        <div v-else-if="!addMode" class="empty-notes">
          아직 등록된 이벤트가 없습니다. '+ 추가' 버튼으로 기록을 시작하세요.
        </div>
      </div>
    </template>
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'
import { useRoute } from 'vue-router'
import { Line } from 'vue-chartjs'
import {
  Chart as ChartJS, CategoryScale, LinearScale, LineElement,
  PointElement, Title, Tooltip, Legend,
} from 'chart.js'
import { getEffectiveNameMonthly, getModelNotes, createModelNote, updateModelNote, deleteModelNote } from '../api'

ChartJS.register(CategoryScale, LinearScale, LineElement, PointElement, Title, Tooltip, Legend)

const route = useRoute()
const modelName = decodeURIComponent(route.params.name)

const loading = ref(true)
const vocMonthly = ref([])   // [{month, count}]
const qdataMonthly = ref([]) // [{month, count}]
const notes = ref([])

// ── 차트 ────────────────────────────────────────────────────────────────────

const monthlyRows = computed(() => {
  const months = new Set([
    ...vocMonthly.value.map(d => d.month),
    ...qdataMonthly.value.map(d => d.month),
  ])
  const vocMap = Object.fromEntries(vocMonthly.value.map(d => [d.month, d.count]))
  const qdataMap = Object.fromEntries(qdataMonthly.value.map(d => [d.month, d.count]))
  return [...months].sort().reverse().map(m => ({
    month: m,
    voc: vocMap[m] || 0,
    qdata: qdataMap[m] || 0,
  }))
})

const chartData = computed(() => {
  const rows = [...monthlyRows.value].reverse() // chronological order for chart
  if (!rows.length) return null
  return {
    labels: rows.map(r => r.month),
    datasets: [
      {
        label: 'Members issue',
        data: rows.map(r => r.voc),
        borderColor: '#3f51b5',
        backgroundColor: 'rgba(63,81,181,0.1)',
        tension: 0.3,
        fill: true,
        pointRadius: 3,
      },
      {
        label: 'Q-data',
        data: rows.map(r => r.qdata),
        borderColor: '#ff9800',
        backgroundColor: 'rgba(255,152,0,0.1)',
        tension: 0.3,
        fill: true,
        pointRadius: 3,
      },
    ],
  }
})

const chartOptions = {
  responsive: true,
  plugins: { legend: { position: 'top' } },
  scales: { y: { beginAtZero: true } },
}

// ── 데이터 로드 ──────────────────────────────────────────────────────────────

async function load() {
  loading.value = true
  try {
    const [statsRes, notesRes] = await Promise.all([
      getEffectiveNameMonthly(modelName),
      getModelNotes(modelName),
    ])
    vocMonthly.value = statsRes.data.voc_monthly || []
    qdataMonthly.value = statsRes.data.qdata_monthly || []
    notes.value = notesRes.data
  } catch (e) {
    console.error(e)
  } finally {
    loading.value = false
  }
}

// ── 노트 CRUD ────────────────────────────────────────────────────────────────

const addMode = ref(false)
const form = ref({ date: '', content: '' })
const editingId = ref(null)
const editForm = ref({ date: '', content: '' })
const saving = ref(false)

function today() {
  return new Date().toISOString().slice(0, 10)
}

function startAdd() {
  addMode.value = true
  form.value = { date: today(), content: '' }
  editingId.value = null
}

function cancelAdd() {
  addMode.value = false
}

async function submitAdd() {
  if (!form.value.date || !form.value.content.trim()) return
  saving.value = true
  try {
    const res = await createModelNote(modelName, form.value)
    notes.value.unshift(res.data)
    addMode.value = false
  } catch (e) {
    console.error(e)
  } finally {
    saving.value = false
  }
}

function startEdit(note) {
  editingId.value = note.id
  editForm.value = { date: note.date, content: note.content }
  addMode.value = false
}

function cancelEdit() {
  editingId.value = null
}

async function submitEdit(noteId) {
  if (!editForm.value.date || !editForm.value.content.trim()) return
  saving.value = true
  try {
    const res = await updateModelNote(noteId, editForm.value)
    const idx = notes.value.findIndex(n => n.id === noteId)
    if (idx !== -1) notes.value[idx] = res.data
    editingId.value = null
    // re-sort by date desc
    notes.value.sort((a, b) => b.date.localeCompare(a.date))
  } catch (e) {
    console.error(e)
  } finally {
    saving.value = false
  }
}

async function deleteNote(noteId) {
  if (!confirm('삭제하시겠습니까?')) return
  try {
    await deleteModelNote(noteId)
    notes.value = notes.value.filter(n => n.id !== noteId)
  } catch (e) {
    console.error(e)
  }
}

onMounted(load)
</script>

<style scoped>
.model-detail-view { padding: 4px 0; max-width: 900px; margin: 0 auto; }

.header-row { display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; }
.btn-close { padding: 7px 16px; background: #e8eaf6; color: #1a237e; border: none; border-radius: 6px; cursor: pointer; font-size: 0.9rem; font-weight: 600; }
.btn-close:hover { background: #c5cae9; }
.page-title { font-size: 1.4rem; font-weight: 700; color: #1a237e; margin: 0; }

.card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); margin-bottom: 16px; }
.section-title { font-size: 0.95rem; font-weight: 600; color: #333; margin-bottom: 14px; }

.chart-wrap { height: 260px; }
.no-chart { text-align: center; padding: 40px; color: #aaa; }

.table-scroll { overflow-x: auto; }
.table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
.table th, .table td { padding: 8px 12px; border-bottom: 1px solid #f0f0f0; text-align: center; }
.table th { background: #f8f9ff; font-weight: 600; color: #555; }
.dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 4px; vertical-align: middle; }
.voc-dot { background: #3f51b5; }
.qdata-dot { background: #ff9800; }
.badge { padding: 2px 8px; border-radius: 10px; font-weight: 600; font-size: 0.82rem; }
.badge.voc { background: #e8eaf6; color: #1a237e; }
.badge.qdata { background: #fff3e0; color: #e65100; }
.empty { color: #aaa; padding: 20px; }

/* 이벤트 로그 */
.log-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; }
.log-header .section-title { margin-bottom: 0; }
.btn-add-note { padding: 6px 16px; background: #1a237e; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 0.88rem; }
.btn-add-note:hover { background: #283593; }

.note-form { display: flex; gap: 8px; align-items: center; background: #f8f9ff; border-radius: 8px; padding: 12px; margin-bottom: 14px; flex-wrap: wrap; }
.input-date { padding: 7px 10px; border: 1px solid #c5cae9; border-radius: 6px; font-size: 0.88rem; width: 140px; }
.input-content { flex: 1; min-width: 200px; padding: 7px 10px; border: 1px solid #c5cae9; border-radius: 6px; font-size: 0.88rem; }
.btn-save { padding: 7px 16px; background: #1a237e; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 0.88rem; }
.btn-save:disabled { opacity: 0.5; cursor: not-allowed; }
.btn-cancel { padding: 7px 12px; background: #f5f5f5; color: #555; border: 1px solid #ddd; border-radius: 6px; cursor: pointer; font-size: 0.88rem; }

.notes-list { display: flex; flex-direction: column; gap: 6px; }
.note-item { display: flex; align-items: center; gap: 12px; padding: 10px 14px; background: #fafafa; border-radius: 8px; border-left: 3px solid #3f51b5; flex-wrap: wrap; }
.note-date { font-size: 0.85rem; font-weight: 600; color: #3f51b5; min-width: 90px; }
.note-content { flex: 1; font-size: 0.88rem; color: #333; min-width: 180px; }
.note-actions { display: flex; gap: 6px; margin-left: auto; }
.btn-edit { padding: 4px 10px; background: #e8eaf6; color: #1a237e; border: none; border-radius: 5px; cursor: pointer; font-size: 0.8rem; }
.btn-delete { padding: 4px 10px; background: #fce4ec; color: #b71c1c; border: none; border-radius: 5px; cursor: pointer; font-size: 0.8rem; }

.empty-notes { text-align: center; padding: 24px; color: #bbb; font-size: 0.88rem; }
.loading { text-align: center; padding: 60px; color: #888; }
</style>
