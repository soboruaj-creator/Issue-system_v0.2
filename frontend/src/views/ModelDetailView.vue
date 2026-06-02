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
      <!-- 추이 차트 (월별 or 주차별) -->
      <div class="card">
        <div class="section-title">
          {{ showWeekly ? '주차별 인입 추이 (개발이슈 + UT)' : '월별 인입 추이 (Members issue + Q-data + 개발이슈 + UT)' }}
        </div>
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
                <th><span class="dot dev-dot"></span> 개발이슈</th>
                <th><span class="dot ut-dot"></span> UT</th>
                <th>합계</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="row in monthlyRows" :key="row.month">
                <td>{{ row.month }}</td>
                <td><span class="badge voc">{{ row.voc }}</span></td>
                <td><span class="badge qdata">{{ row.qdata }}</span></td>
                <td><span class="badge dev">{{ row.dev }}</span></td>
                <td><span class="badge ut">{{ row.ut }}</span></td>
                <td><strong>{{ row.voc + row.qdata + row.dev + row.ut }}</strong></td>
              </tr>
              <tr v-if="!monthlyRows.length">
                <td colspan="6" class="empty">데이터 없음</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- 개발 이슈 현황 -->
      <div class="card">
        <div class="section-title">개발 이슈 현황</div>
        <div v-if="devLoading" class="no-chart">로딩 중...</div>
        <div v-else-if="devIssueError" class="no-chart">개발 이슈 데이터 없음</div>
        <template v-else>
          <!-- 자체이슈 / UT이슈 카드 -->
          <div class="dev-type-cards">
            <!-- 개발이슈(자체) -->
            <div class="dev-type-card dev-card">
              <div class="dev-card-title">개발이슈 (자체)</div>
              <div class="dev-stats-row">
                <div class="dev-stat-item">
                  <span class="dev-stat-label">Total</span>
                  <span class="dev-stat-val">{{ statsDev.total }}</span>
                </div>
                <div class="dev-stat-item">
                  <span class="dev-stat-label">Open</span>
                  <span class="badge dev-open">{{ statsDev.open }}</span>
                </div>
                <div class="dev-stat-item">
                  <span class="dev-stat-label">Resolve</span>
                  <span class="badge dev-resolve">{{ statsDev.resolve }}</span>
                </div>
                <div class="dev-stat-item">
                  <span class="dev-stat-label">Pending</span>
                  <span class="badge dev-pending">{{ statsDev.pending }}</span>
                </div>
                <div class="dev-stat-item">
                  <span class="dev-stat-label">Close</span>
                  <span class="badge dev-close">{{ statsDev.close }}</span>
                </div>
              </div>
            </div>
            <!-- UT이슈 -->
            <div class="dev-type-card ut-card">
              <div class="dev-card-title">UT 이슈</div>
              <div class="dev-stats-row">
                <div class="dev-stat-item">
                  <span class="dev-stat-label">Total</span>
                  <span class="dev-stat-val">{{ statsUt.total }}</span>
                </div>
                <div class="dev-stat-item">
                  <span class="dev-stat-label">Open</span>
                  <span class="badge dev-open">{{ statsUt.open }}</span>
                </div>
                <div class="dev-stat-item">
                  <span class="dev-stat-label">Resolve</span>
                  <span class="badge dev-resolve">{{ statsUt.resolve }}</span>
                </div>
                <div class="dev-stat-item">
                  <span class="dev-stat-label">Pending</span>
                  <span class="badge dev-pending">{{ statsUt.pending }}</span>
                </div>
                <div class="dev-stat-item">
                  <span class="dev-stat-label">Close</span>
                  <span class="badge dev-close">{{ statsUt.close }}</span>
                </div>
              </div>
            </div>
          </div>

          <!-- 이슈 리스트 (close 제외, pending+close 포함) -->
          <div v-if="!devIssues.length" class="no-chart">활성 이슈 없음</div>
          <div v-else class="issue-list">
            <div v-for="iss in devIssues" :key="iss.case_code"
                 class="issue-item" :class="{ 'pending-row': iss.is_pending }">
              <div class="issue-row">
                <span class="badge" :class="'dev-status-' + iss.status">{{ iss.is_pending && iss.status === 'close' ? 'close/Pending' : iss.status }}</span>
                <span class="issue-title">{{ iss.title }}</span>
                <span class="issue-code">{{ iss.case_code }}</span>
              </div>
              <div v-if="iss.is_pending" class="pending-detail">
                <span v-if="iss.pending_memo" class="pending-memo">📝 {{ iss.pending_memo }}</span>
                <div v-if="iss.pending_attachments?.length" class="attach-list">
                  <a v-for="att in iss.pending_attachments" :key="att.stored_name"
                     :href="getDevIssueAttachmentUrl(iss.case_code, att.stored_name)"
                     class="attach-link" download>
                    📎 {{ att.filename }}
                  </a>
                </div>
              </div>
            </div>
          </div>
        </template>
      </div>

      <!-- Members issue 현황 -->
      <div class="card">
        <div class="section-title">Members issue 현황</div>
        <div v-if="!vocIssues.length" class="no-chart">활성 이슈 없음</div>
        <div v-else class="issue-list">
          <div v-for="iss in vocIssues" :key="iss.case_code"
               class="issue-item" :class="{ 'pending-row': iss.is_pending }">
            <div class="issue-row">
              <span class="badge" :class="'voc-status-' + (iss.status || 'open')">{{ iss.is_pending && iss.status === 'close' ? 'close/Pending' : (iss.status || 'open') }}</span>
              <span class="issue-title">{{ iss.title }}</span>
              <span class="issue-code">{{ iss.case_code }}</span>
            </div>
            <div v-if="iss.is_pending && iss.pending_memo" class="pending-detail">
              <span class="pending-memo">📝 {{ iss.pending_memo }}</span>
            </div>
          </div>
        </div>
      </div>

      <!-- 모델 이벤트 로그 -->
      <div class="card">
        <div class="log-header">
          <div class="section-title">모델 이벤트 로그</div>
          <button class="btn-add-note" @click="startAdd">+ 추가</button>
        </div>

        <div v-if="addMode" class="note-form">
          <input type="date" v-model="form.date" class="input-date" />
          <input type="text" v-model="form.content" class="input-content" placeholder="내용 입력 (예: OTA v2.0 배포, 배터리 이슈 접수 시작 등)" @keyup.enter="submitAdd" />
          <button class="btn-save" @click="submitAdd" :disabled="saving">저장</button>
          <button class="btn-cancel" @click="cancelAdd">취소</button>
        </div>

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
import {
  getEffectiveNameMonthly, getModelNotes, createModelNote, updateModelNote, deleteModelNote,
  getModelDevIssues, getModelDevIssuesMonthly, getModelDevIssuesWeekly, getDevIssueAttachmentUrl,
  getVocModelActive,
} from '../api'

ChartJS.register(CategoryScale, LinearScale, LineElement, PointElement, Title, Tooltip, Legend)

const route = useRoute()
const modelName = decodeURIComponent(route.params.name)

const loading = ref(true)
const vocMonthly = ref([])
const qdataMonthly = ref([])
const devMonthly = ref([])
const devWeekly = ref([])
const notes = ref([])
const devIssues = ref([])
const statsDev = ref({ total: 0, open: 0, resolve: 0, close: 0, pending: 0 })
const statsUt = ref({ total: 0, open: 0, resolve: 0, close: 0, pending: 0 })
const devLoading = ref(true)
const devIssueError = ref(false)
const vocIssues = ref([])

const showWeekly = computed(() => vocMonthly.value.length === 0 && qdataMonthly.value.length === 0)

// ── 월별 합산 ────────────────────────────────────────────────────────────────

const monthlyRows = computed(() => {
  const months = new Set([
    ...vocMonthly.value.map(d => d.month),
    ...qdataMonthly.value.map(d => d.month),
    ...devMonthly.value.map(d => d.month),
  ])
  const vocMap = Object.fromEntries(vocMonthly.value.map(d => [d.month, d.count]))
  const qdataMap = Object.fromEntries(qdataMonthly.value.map(d => [d.month, d.count]))
  const devMap = Object.fromEntries(devMonthly.value.map(d => [d.month, d.dev_count]))
  const utMap = Object.fromEntries(devMonthly.value.map(d => [d.month, d.ut_count]))
  return [...months].sort().reverse().map(m => ({
    month: m,
    voc: vocMap[m] || 0,
    qdata: qdataMap[m] || 0,
    dev: devMap[m] || 0,
    ut: utMap[m] || 0,
  }))
})

const chartData = computed(() => {
  if (showWeekly.value) {
    const rows = [...devWeekly.value]
    if (!rows.length) return null
    return {
      labels: rows.map(r => r.week),
      datasets: [
        {
          label: '개발이슈',
          data: rows.map(r => r.dev_count),
          borderColor: '#43a047',
          backgroundColor: 'rgba(67,160,71,0.08)',
          tension: 0.3,
          fill: true,
          pointRadius: 3,
        },
        {
          label: 'UT',
          data: rows.map(r => r.ut_count),
          borderColor: '#e53935',
          backgroundColor: 'rgba(229,57,53,0.08)',
          tension: 0.3,
          fill: true,
          pointRadius: 3,
        },
      ],
    }
  }
  const rows = [...monthlyRows.value].reverse()
  if (!rows.length) return null
  return {
    labels: rows.map(r => r.month),
    datasets: [
      {
        label: 'Members issue',
        data: rows.map(r => r.voc),
        borderColor: '#3f51b5',
        backgroundColor: 'rgba(63,81,181,0.08)',
        tension: 0.3,
        fill: true,
        pointRadius: 3,
      },
      {
        label: 'Q-data',
        data: rows.map(r => r.qdata),
        borderColor: '#ff9800',
        backgroundColor: 'rgba(255,152,0,0.08)',
        tension: 0.3,
        fill: true,
        pointRadius: 3,
      },
      {
        label: '개발이슈',
        data: rows.map(r => r.dev),
        borderColor: '#43a047',
        backgroundColor: 'rgba(67,160,71,0.08)',
        tension: 0.3,
        fill: true,
        pointRadius: 3,
      },
      {
        label: 'UT',
        data: rows.map(r => r.ut),
        borderColor: '#e53935',
        backgroundColor: 'rgba(229,57,53,0.08)',
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
  devLoading.value = true
  const [statsRes, notesRes, devRes, devMonthlyRes, devWeeklyRes, vocRes] = await Promise.allSettled([
    getEffectiveNameMonthly(modelName),
    getModelNotes(modelName),
    getModelDevIssues(modelName),
    getModelDevIssuesMonthly(modelName),
    getModelDevIssuesWeekly(modelName),
    getVocModelActive(modelName),
  ])
  if (statsRes.status === 'fulfilled') {
    vocMonthly.value = statsRes.value.data.voc_monthly || []
    qdataMonthly.value = statsRes.value.data.qdata_monthly || []
  }
  if (notesRes.status === 'fulfilled') notes.value = notesRes.value.data
  if (devRes.status === 'fulfilled') {
    statsDev.value = devRes.value.data.stats_dev || statsDev.value
    statsUt.value = devRes.value.data.stats_ut || statsUt.value
    devIssues.value = devRes.value.data.issues || []
    devIssueError.value = false
  } else {
    devIssueError.value = true
  }
  if (devMonthlyRes.status === 'fulfilled') devMonthly.value = devMonthlyRes.value.data || []
  if (devWeeklyRes.status === 'fulfilled') devWeekly.value = devWeeklyRes.value.data || []
  if (vocRes.status === 'fulfilled') vocIssues.value = vocRes.value.data || []
  loading.value = false
  devLoading.value = false
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
.model-detail-view { padding: 4px 0; max-width: 960px; margin: 0 auto; }

.header-row { display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; }
.btn-close { padding: 7px 16px; background: #e8eaf6; color: #1a237e; border: none; border-radius: 6px; cursor: pointer; font-size: 0.9rem; font-weight: 600; }
.btn-close:hover { background: #c5cae9; }
.page-title { font-size: 1.4rem; font-weight: 700; color: #1a237e; margin: 0; }

.card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); margin-bottom: 16px; }
.section-title { font-size: 0.95rem; font-weight: 600; color: #333; margin-bottom: 14px; }

.chart-wrap { height: 280px; }
.no-chart { text-align: center; padding: 40px; color: #aaa; }

.table-scroll { overflow-x: auto; }
.table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
.table th, .table td { padding: 8px 12px; border-bottom: 1px solid #f0f0f0; text-align: center; }
.table th { background: #f8f9ff; font-weight: 600; color: #555; }
.dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 4px; vertical-align: middle; }
.voc-dot { background: #3f51b5; }
.qdata-dot { background: #ff9800; }
.dev-dot { background: #43a047; }
.ut-dot { background: #e53935; }
.badge { padding: 2px 8px; border-radius: 10px; font-weight: 600; font-size: 0.82rem; }
.badge.voc { background: #e8eaf6; color: #1a237e; }
.badge.qdata { background: #fff3e0; color: #e65100; }
.badge.dev { background: #e8f5e9; color: #2e7d32; }
.badge.ut { background: #fce4ec; color: #c62828; }
.empty { color: #aaa; padding: 20px; }

/* 개발이슈 현황 카드 */
.dev-type-cards { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 4px; }
@media (max-width: 600px) { .dev-type-cards { grid-template-columns: 1fr; } }
.dev-type-card { border-radius: 8px; padding: 14px 16px; }
.dev-card { background: #f1f8e9; border-left: 4px solid #43a047; }
.ut-card { background: #fce4ec; border-left: 4px solid #e53935; }
.dev-card-title { font-size: 0.88rem; font-weight: 700; margin-bottom: 10px; color: #333; }
.dev-stats-row { display: flex; gap: 14px; flex-wrap: wrap; }
.dev-stat-item { display: flex; flex-direction: column; align-items: center; gap: 4px; }
.dev-stat-label { font-size: 0.75rem; color: #888; font-weight: 500; }
.dev-stat-val { font-size: 1.1rem; font-weight: 700; color: #333; }
.badge.dev-open { background: #e3f2fd; color: #1565c0; padding: 3px 10px; border-radius: 10px; font-weight: 700; font-size: 0.82rem; }
.badge.dev-resolve { background: #e8f5e9; color: #2e7d32; padding: 3px 10px; border-radius: 10px; font-weight: 700; font-size: 0.82rem; }
.badge.dev-pending { background: #fff3e0; color: #e65100; padding: 3px 10px; border-radius: 10px; font-weight: 700; font-size: 0.82rem; }
.badge.dev-close { background: #f5f5f5; color: #757575; padding: 3px 10px; border-radius: 10px; font-weight: 700; font-size: 0.82rem; }
.badge.dev-ut { background: #fce4ec; color: #c62828; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.dev-self { background: #f3e5f5; color: #6a1b9a; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.dev-none { background: #f5f5f5; color: #bbb; padding: 2px 8px; border-radius: 10px; font-size: 0.78rem; }
.badge.dev-status-open { background: #e3f2fd; color: #1565c0; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.dev-status-resolve { background: #e8f5e9; color: #2e7d32; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.dev-status-close { background: #f5f5f5; color: #757575; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.voc-status-open { background: #e3f2fd; color: #1565c0; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.voc-status-resolve { background: #e8f5e9; color: #2e7d32; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.badge.voc-status-close { background: #f5f5f5; color: #757575; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.78rem; }
.issue-list { display: flex; flex-direction: column; gap: 6px; margin-top: 12px; }
.issue-item { border-radius: 7px; background: #fafafa; border-left: 3px solid #e0e0e0; overflow: hidden; }
.issue-item.pending-row { background: #fffde7; border-left-color: #ffa000; }
.issue-row { display: flex; align-items: center; gap: 10px; padding: 8px 12px; }
.issue-title { flex: 1; font-size: 0.85rem; color: #333; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.issue-code { font-size: 0.78rem; color: #999; white-space: nowrap; font-family: monospace; }
.pending-detail { padding: 4px 12px 8px 12px; display: flex; flex-direction: column; gap: 4px; }
.pending-memo { font-size: 0.82rem; color: #795548; }
.attach-list { display: flex; flex-wrap: wrap; gap: 6px; }
.attach-link { font-size: 0.8rem; color: #1565c0; text-decoration: none; background: #e3f2fd; padding: 2px 8px; border-radius: 4px; }
.attach-link:hover { text-decoration: underline; }

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
