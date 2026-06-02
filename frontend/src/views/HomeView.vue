<template>
  <div class="home">
    <h1 class="page-title">대시보드</h1>

    <div v-if="loading" class="loading">로딩 중...</div>

    <template v-else-if="data">
      <div class="home-grid">

        <!-- 좌: 개발 이슈 현황 -->
        <div class="dev-col">
          <div class="card dev-issues-panel">
            <h2 class="section-title">개발 이슈 현황 <span class="count-badge dev">{{ devIssueList.length }}건</span></h2>
            <div v-if="devLoading" class="no-data">로딩 중...</div>
            <div v-else-if="!devIssueList.length" class="no-data">활성 이슈 없음</div>
            <div v-else class="dev-entry-list">

              <!-- 모델별 아코디언 (non-pending) -->
              <div v-for="group in devIssuesByModel" :key="group.model" class="dev-model-group">
                <div class="dev-model-header" @click="toggleModel(group.model)">
                  <span class="dev-collapse-icon">{{ openedModels.has(group.model) ? '▾' : '▸' }}</span>
                  <span class="dev-model-name-text" @click.stop="openModelDetail(group.model)">{{ group.model }}</span>
                  <span class="dev-badge-wrap">
                    <span v-if="group.openCount" class="badge dev-s-open">Open {{ group.openCount }}</span>
                    <span v-if="group.resolveCount" class="badge dev-s-resolve">Resolve {{ group.resolveCount }}</span>
                  </span>
                </div>
                <div v-show="openedModels.has(group.model)" class="dev-issue-items">
                  <div v-for="(iss, idx) in group.issues" :key="idx" class="dev-entry-item">
                    <span class="badge" :class="'dev-s-' + iss.status">{{ iss.status }}</span>
                    <span class="dev-entry-title">{{ iss.title || '(제목 없음)' }}</span>
                  </div>
                </div>
              </div>

              <!-- Pending 섹션 -->
              <div v-if="pendingIssues.length" class="dev-pending-section">
                <div class="dev-pending-header">
                  <span>⏸ Pending</span>
                  <span class="badge dev-s-pending">{{ pendingIssues.length }}건</span>
                </div>
                <div v-for="(iss, idx) in pendingIssues" :key="'p'+idx"
                     class="dev-entry-item pending-item" @click="openModelDetail(iss.model_name)">
                  <span class="dev-pending-model">{{ iss.model_name }}</span>
                  <span class="dev-entry-title">{{ iss.title || '(제목 없음)' }}</span>
                </div>
              </div>

            </div>
          </div>
        </div>

        <!-- 우: 기존 대시보드 -->
        <div class="main-col">

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

              <!-- 수정완료 섹션 -->
              <div class="fix-section">
                <h3 class="fix-title">🔧 수정완료 (Resolve) <span class="count-badge fix">{{ resolvedWithFix.length }}건</span></h3>
                <div v-if="fixLoading" class="no-data">로딩 중...</div>
                <div v-else-if="!resolvedWithFix.length" class="no-data">해당 이슈 없음</div>
                <div v-else class="fix-list">
                  <div v-for="(item, idx) in resolvedWithFix" :key="idx" class="fix-item">
                    <span class="fix-model">{{ item.model_name }}</span>
                    <span class="fix-summary">{{ item.title || '(내용 없음)' }}</span>
                    <span class="fix-date">{{ (item.created_date || '').slice(0, 7) }}</span>
                  </div>
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

        </div><!-- /main-col -->
      </div><!-- /home-grid -->
    </template>

    <div v-else-if="!loading" class="empty-state">
      <p>데이터가 없습니다. VOC 파일을 업로드해주세요.</p>
      <RouterLink to="/upload" class="btn-primary">업로드 바로가기</RouterLink>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, watch, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { Bar } from 'vue-chartjs'
import {
  Chart as ChartJS, CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend,
} from 'chart.js'
import { getDashboard, getDevIssueActiveList, getResolvedWithFix } from '../api'

ChartJS.register(CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend)

const router = useRouter()
const data = ref(null)
const loading = ref(false)
const devIssueList = ref([])
const devLoading = ref(false)
const resolvedWithFix = ref([])
const fixLoading = ref(false)
const openedModels = ref(new Set())

const pendingIssues = computed(() =>
  devIssueList.value.filter(i => i.is_pending)
)

const devIssuesByModel = computed(() => {
  const groups = {}
  for (const iss of devIssueList.value) {
    if (iss.is_pending) continue
    if (!iss.title) continue
    if (!groups[iss.model_name]) groups[iss.model_name] = []
    groups[iss.model_name].push(iss)
  }
  return Object.entries(groups)
    .map(([model, issues]) => ({
      model,
      issues,
      openCount: issues.filter(i => i.status === 'open').length,
      resolveCount: issues.filter(i => i.status === 'resolve').length,
    }))
    .sort((a, b) => b.openCount - a.openCount || a.model.localeCompare(b.model))
})

watch(devIssuesByModel, (groups) => {
  const s = new Set()
  groups.forEach(g => { if (g.openCount > 0) s.add(g.model) })
  openedModels.value = s
}, { immediate: true })

function toggleModel(model) {
  const s = new Set(openedModels.value)
  s.has(model) ? s.delete(model) : s.add(model)
  openedModels.value = s
}

function formatUpload(dt) {
  if (!dt) return '없음'
  return dt.slice(0, 16)
}

function openModelDetail(name) {
  const url = router.resolve({ name: 'model-detail', params: { name } }).href
  window.open(url, '_blank')
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
  devLoading.value = true
  fixLoading.value = true
  const [dashRes, devRes, fixRes] = await Promise.allSettled([
    getDashboard(), getDevIssueActiveList(), getResolvedWithFix()
  ])
  if (dashRes.status === 'fulfilled') data.value = dashRes.value.data
  if (devRes.status === 'fulfilled') devIssueList.value = devRes.value.data
  if (fixRes.status === 'fulfilled') resolvedWithFix.value = fixRes.value.data
  loading.value = false
  devLoading.value = false
  fixLoading.value = false
})
</script>

<style scoped>
.page-title { font-size: 1.5rem; font-weight: 700; margin-bottom: 20px; color: #1a237e; }
.loading { text-align: center; padding: 60px; color: #888; }

.home-grid { display: grid; grid-template-columns: 380px 1fr; gap: 16px; align-items: start; }
@media (max-width: 1200px) { .home-grid { grid-template-columns: 1fr; } }

.dev-col { min-width: 0; }
.main-col { min-width: 0; }

/* Dev issues panel */
.dev-issues-panel { height: fit-content; }
.count-badge.dev { background: #e8eaf6; color: #1a237e; }
.dev-entry-list { display: flex; flex-direction: column; gap: 6px; }

.dev-model-group { border-radius: 8px; background: #f8f9ff; overflow: hidden; border: 1px solid #e8eaf6; }
.dev-model-header { display: flex; align-items: center; gap: 6px; padding: 7px 10px; background: #e8eaf6; cursor: pointer; user-select: none; }
.dev-model-header:hover { background: #c5cae9; }
.dev-collapse-icon { font-size: 0.8rem; color: #666; width: 12px; flex-shrink: 0; }
.dev-model-name-text { font-weight: 700; font-size: 0.86rem; color: #1a237e; flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.dev-model-name-text:hover { text-decoration: underline; }
.dev-badge-wrap { display: flex; gap: 4px; flex-shrink: 0; }

.dev-issue-items { display: flex; flex-direction: column; }
.dev-entry-item { display: flex; align-items: flex-start; gap: 6px; padding: 5px 10px; border-bottom: 1px solid #eef0fb; }
.dev-entry-item:last-child { border-bottom: none; }
.dev-entry-title { font-size: 0.81rem; color: #444; line-height: 1.4; }

.dev-pending-section { border-radius: 8px; border: 1px solid #ffe0b2; overflow: hidden; margin-top: 4px; }
.dev-pending-header { display: flex; align-items: center; justify-content: space-between; padding: 7px 10px; background: #fff3e0; font-weight: 700; font-size: 0.86rem; color: #e65100; }
.pending-item { background: #fffde7; cursor: pointer; }
.pending-item:hover { background: #fff9c4; }
.dev-pending-model { font-size: 0.78rem; font-weight: 700; color: #e65100; white-space: nowrap; flex-shrink: 0; }

.badge.dev-s-open { background: #e3f2fd; color: #1565c0; font-size: 0.74rem; padding: 1px 6px; white-space: nowrap; }
.badge.dev-s-resolve { background: #e8f5e9; color: #2e7d32; font-size: 0.74rem; padding: 1px 6px; white-space: nowrap; }
.badge.dev-s-close { background: #f5f5f5; color: #888; font-size: 0.74rem; padding: 1px 6px; white-space: nowrap; }
.badge.dev-s-pending { background: #fff3e0; color: #e65100; font-size: 0.74rem; padding: 1px 6px; white-space: nowrap; }

/* Stat row */
.stat-row { display: flex; gap: 14px; flex-wrap: wrap; margin-bottom: 16px; }
.card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }
.stat-card { min-width: 150px; flex: 1; }
.stat-label { font-size: 0.82rem; color: #888; margin-bottom: 8px; }
.stat-date { font-weight: 600; color: #555; }
.stat-value { font-size: 2rem; font-weight: 800; line-height: 1; }
.stat-value.voc { color: #1a237e; }
.stat-value.qdata { color: #bf360c; }
.stat-unit { font-size: 1rem; font-weight: 400; margin-left: 2px; }
.upload-card { min-width: 180px; }
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

.entry-list { max-height: 280px; overflow-y: auto; display: flex; flex-direction: column; gap: 6px; }
.entry-item { display: flex; gap: 8px; align-items: flex-start; padding: 7px 8px; border-radius: 6px; background: #f8f9ff; font-size: 0.83rem; }
.entry-model { font-weight: 700; color: #1a237e; white-space: nowrap; min-width: 72px; }
.entry-summary { color: #444; line-height: 1.4; overflow: hidden; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
.no-data { color: #aaa; font-size: 0.85rem; padding: 12px 0; }

.fix-section { margin-top: 14px; border-top: 1px solid #eee; padding-top: 12px; }
.fix-title { font-size: 0.88rem; font-weight: 700; color: #1b5e20; margin-bottom: 8px; display: flex; align-items: center; gap: 6px; }
.count-badge.fix { background: #e8f5e9; color: #2e7d32; }
.fix-list { max-height: 220px; overflow-y: auto; display: flex; flex-direction: column; gap: 5px; }
.fix-item { display: flex; gap: 8px; align-items: flex-start; padding: 6px 8px; border-radius: 6px; background: #f1f8e9; border-left: 3px solid #66bb6a; font-size: 0.82rem; }
.fix-model { font-weight: 700; color: #2e7d32; white-space: nowrap; min-width: 72px; }
.fix-summary { flex: 1; color: #333; line-height: 1.4; overflow: hidden; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
.fix-date { font-size: 0.76rem; color: #888; white-space: nowrap; }

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
