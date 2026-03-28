<template>
  <div class="voc-list">
    <h1 class="page-title">VOC 목록</h1>

    <!-- 검색/필터 -->
    <div class="card filter-bar">
      <input v-model="search" placeholder="사례코드 / 제목 / 문제 검색" @keyup.enter="fetchList" class="search-input" />
      <input v-model="modelFilter" placeholder="모델명 필터" class="filter-input" />
      <input type="date" v-model="startDate" />
      <input type="date" v-model="endDate" />
      <button class="btn-primary" @click="fetchList">검색</button>
      <button class="btn-secondary" @click="exportExcel">엑셀 다운로드</button>
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
              <th>사례코드</th><th>모델명</th><th>칩셋</th><th>이슈유형</th>
              <th>문제</th><th>OS</th><th>등록일</th><th>액션</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="voc in list" :key="voc.case_code" @click="selectVoc(voc)" class="clickable-row">
              <td><code class="case-code">{{ voc.case_code }}</code></td>
              <td>{{ voc.model_name || '-' }}</td>
              <td><span v-if="voc.chipset" class="chip-tag">{{ voc.chipset }}</span><span v-else class="no-data">-</span></td>
              <td>
                <span :class="['issue-badge', voc.issue_type === '내부이슈' ? 'internal' : 'external']">
                  {{ voc.issue_type || '-' }}
                </span>
              </td>
              <td class="problem-cell">{{ truncate(voc.problem, 60) }}</td>
              <td>{{ voc.os_version || '-' }}</td>
              <td>{{ voc.created_date || '-' }}</td>
              <td><button class="btn-detail" @click.stop="selectVoc(voc)">상세</button></td>
            </tr>
            <tr v-if="!list.length">
              <td colspan="8" class="empty">데이터가 없습니다.</td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- 페이지네이션 -->
      <div class="pagination">
        <button @click="changePage(page - 1)" :disabled="page <= 1">이전</button>
        <span>{{ page }} / {{ totalPages }}</span>
        <button @click="changePage(page + 1)" :disabled="page >= totalPages">다음</button>
      </div>
    </div>

    <!-- 상세 모달 -->
    <VocDetailModal v-if="selectedVoc" :voc="selectedVoc" @close="selectedVoc = null" />
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'
import { listVoc } from '../api'
import { exportVocExcel } from '../api'
import VocDetailModal from '../components/VocDetailModal.vue'

const list = ref([])
const total = ref(0)
const page = ref(1)
const pageSize = 20
const loading = ref(false)
const search = ref('')
const modelFilter = ref('')
const startDate = ref('')
const endDate = ref('')
const selectedVoc = ref(null)

const totalPages = computed(() => Math.max(1, Math.ceil(total.value / pageSize)))

async function fetchList() {
  loading.value = true
  try {
    const params = { page: page.value, size: pageSize }
    if (search.value) params.search = search.value
    if (modelFilter.value) params.model_name = modelFilter.value
    if (startDate.value) params.start_date = startDate.value
    if (endDate.value) params.end_date = endDate.value
    const res = await listVoc(params)
    list.value = res.data.items
    total.value = res.data.total
  } catch (e) {
    console.error(e)
  } finally {
    loading.value = false
  }
}

function changePage(p) {
  if (p < 1 || p > totalPages.value) return
  page.value = p
  fetchList()
}

function selectVoc(voc) { selectedVoc.value = voc }

function truncate(str, len) {
  if (!str) return '-'
  return str.length > len ? str.slice(0, len) + '…' : str
}

function exportExcel() {
  exportVocExcel({ start_date: startDate.value, end_date: endDate.value, model_name: modelFilter.value })
}

onMounted(fetchList)
</script>

<style scoped>
.page-title { font-size: 1.5rem; font-weight: 700; margin-bottom: 20px; color: #1a237e; }
.card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); margin-bottom: 16px; }
.filter-bar { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
.search-input { flex: 1; min-width: 200px; padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; }
.filter-input { width: 140px; padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; }
.filter-bar input[type=date] { padding: 7px 10px; border: 1px solid #ddd; border-radius: 6px; }
.btn-primary { padding: 8px 20px; background: #1a237e; color: #fff; border: none; border-radius: 6px; cursor: pointer; white-space: nowrap; }
.btn-secondary { padding: 8px 20px; background: #fff; color: #1a237e; border: 1px solid #1a237e; border-radius: 6px; cursor: pointer; white-space: nowrap; }
.table-header { margin-bottom: 12px; }
.total-count { font-size: 0.9rem; color: #666; }
.table-wrap { overflow-x: auto; }
.table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
.table th { background: #f8f9ff; padding: 10px 12px; text-align: left; font-weight: 600; color: #555; border-bottom: 2px solid #e8eaf6; white-space: nowrap; }
.table td { padding: 9px 12px; border-bottom: 1px solid #f0f0f0; }
.clickable-row { cursor: pointer; transition: background 0.15s; }
.clickable-row:hover { background: #f8f9ff; }
.case-code { background: #e8eaf6; padding: 2px 6px; border-radius: 4px; font-size: 0.8rem; }
.chip-tag { background: #e3f2fd; color: #1565c0; padding: 2px 8px; border-radius: 10px; font-size: 0.78rem; }
.no-data { color: #bbb; }
.issue-badge { padding: 2px 8px; border-radius: 10px; font-size: 0.78rem; }
.issue-badge.internal { background: #e8eaf6; color: #3949ab; }
.issue-badge.external { background: #fff3e0; color: #e65100; }
.problem-cell { max-width: 280px; color: #555; }
.btn-detail { padding: 4px 12px; background: #e8eaf6; color: #1a237e; border: none; border-radius: 4px; cursor: pointer; font-size: 0.8rem; }
.btn-detail:hover { background: #c5cae9; }
.empty { text-align: center; color: #aaa; padding: 40px; }
.pagination { display: flex; align-items: center; justify-content: center; gap: 16px; margin-top: 16px; }
.pagination button { padding: 6px 16px; border: 1px solid #c5cae9; border-radius: 6px; background: #fff; cursor: pointer; }
.pagination button:disabled { opacity: 0.4; cursor: not-allowed; }
.loading { text-align: center; padding: 60px; color: #888; }
</style>
