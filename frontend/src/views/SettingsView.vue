<template>
  <div class="settings-view">
    <h1 class="page-title">설정</h1>

    <!-- 탭 -->
    <div class="tabs">
      <button v-for="t in tabs" :key="t.key" :class="['tab-btn', { active: activeTab === t.key }]" @click="activeTab=t.key">
        {{ t.label }}
      </button>
    </div>

    <!-- 칩셋 매핑 관리 -->
    <div v-if="activeTab === 'chipset'" class="card">
      <h2 class="section-title">칩셋 미매핑 모델</h2>
      <div v-if="unmappedLoading" class="loading-sm">로딩 중...</div>
      <div v-else>
        <p v-if="!unmapped.length" class="empty-sm">미매핑 모델이 없습니다.</p>
        <table v-else class="table">
          <thead><tr><th>모델명</th><th>VOC 건수</th><th>칩셋 매핑</th></tr></thead>
          <tbody>
            <tr v-for="item in unmapped" :key="item.model_name">
              <td>{{ item.model_name }}</td>
              <td><span class="badge">{{ item.count }}</span></td>
              <td>
                <div class="inline-form">
                  <input v-model="chipsetInputs[item.model_name]" placeholder="칩셋명 입력" class="sm-input" />
                  <button class="btn-sm-primary" @click="addMapping(item.model_name)">등록</button>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div class="divider"></div>
      <h2 class="section-title">칩셋 병합 / 이름 변경</h2>
      <div class="form-row">
        <label>소스 칩셋 <input v-model="mergeSource" placeholder="기존 칩셋명" class="form-input" /></label>
        <label>타겟 칩셋 <input v-model="mergeTarget" placeholder="병합할 칩셋명" class="form-input" /></label>
        <button class="btn-primary" @click="mergeChipsets">병합</button>
      </div>
      <div class="form-row mt-8">
        <label>변경 전 <input v-model="renameOld" placeholder="기존 칩셋명" class="form-input" /></label>
        <label>변경 후 <input v-model="renameNew" placeholder="새 칩셋명" class="form-input" /></label>
        <button class="btn-primary" @click="renameChipset">이름 변경</button>
      </div>
      <div v-if="chipsetMsg" class="msg" :class="chipsetMsgType">{{ chipsetMsg }}</div>
    </div>

    <!-- 모델 업데이트 -->
    <div v-if="activeTab === 'model'" class="card">
      <h2 class="section-title">모델 데이터 업데이트</h2>
      <div class="action-list">
        <div class="action-item">
          <div>
            <strong>워치 모델명 업데이트</strong>
            <p>타이틀에서 워치 모델명을 추출하여 VOC 데이터를 갱신합니다.</p>
          </div>
          <button class="btn-primary" @click="updateWatchModels" :disabled="actionLoading">
            {{ actionLoading ? '처리 중...' : '실행' }}
          </button>
        </div>
        <div class="action-item">
          <div>
            <strong>칩셋 매핑 적용</strong>
            <p>등록된 칩셋 매핑 정보를 VOC 데이터에 일괄 적용합니다.</p>
          </div>
          <button class="btn-primary" @click="updateMapping" :disabled="actionLoading">
            {{ actionLoading ? '처리 중...' : '실행' }}
          </button>
        </div>
      </div>
      <div v-if="modelMsg" class="msg" :class="modelMsgType">{{ modelMsg }}</div>
    </div>

    <!-- 데이터 초기화 -->
    <div v-if="activeTab === 'reset'" class="card">
      <h2 class="section-title danger-title">데이터 초기화</h2>
      <p class="warn-text">⚠️ 아래 작업은 되돌릴 수 없습니다. 신중하게 진행하세요.</p>
      <div class="action-list">
        <div class="action-item">
          <div>
            <strong>VOC 데이터 초기화</strong>
            <p>업로드된 모든 VOC 데이터를 삭제합니다.</p>
          </div>
          <button class="btn-danger" @click="resetVoc">초기화</button>
        </div>
        <div class="action-item">
          <div>
            <strong>Q-data 초기화</strong>
            <p>업로드된 모든 Q-data를 삭제합니다.</p>
          </div>
          <button class="btn-danger" @click="resetQData">초기화</button>
        </div>
      </div>
      <div v-if="resetMsg" class="msg" :class="resetMsgType">{{ resetMsg }}</div>
    </div>

    <!-- 메모 백업/복원 -->
    <div v-if="activeTab === 'memo'" class="card">
      <h2 class="section-title">메모 백업 / 복원</h2>
      <div class="action-list">
        <div class="action-item">
          <div>
            <strong>메모 백업</strong>
            <p>월별/주별/모델별 메모를 JSON 파일로 다운로드합니다.</p>
          </div>
          <button class="btn-primary" @click="backupMemos">백업 다운로드</button>
        </div>
        <div class="action-item">
          <div>
            <strong>메모 복원</strong>
            <p>백업 JSON 파일에서 메모를 복원합니다.</p>
          </div>
          <div class="inline-form">
            <input type="file" accept=".json" @change="handleRestoreFile" ref="restoreInput" class="hidden-input" />
            <button class="btn-secondary" @click="$refs.restoreInput.click()">파일 선택</button>
            <button v-if="restoreData" class="btn-primary" @click="restoreMemos">복원 실행</button>
          </div>
        </div>
      </div>
      <div v-if="memoMsg" class="msg success">{{ memoMsg }}</div>
    </div>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import * as api from '../api'
import axios from 'axios'

const tabs = [
  { key: 'chipset', label: '칩셋 관리' },
  { key: 'model', label: '모델 업데이트' },
  { key: 'reset', label: '데이터 초기화' },
  { key: 'memo', label: '메모 백업/복원' },
]
const activeTab = ref('chipset')

// 칩셋
const unmapped = ref([])
const unmappedLoading = ref(false)
const chipsetInputs = ref({})
const mergeSource = ref('')
const mergeTarget = ref('')
const renameOld = ref('')
const renameNew = ref('')
const chipsetMsg = ref('')
const chipsetMsgType = ref('success')

// 모델
const actionLoading = ref(false)
const modelMsg = ref('')
const modelMsgType = ref('success')

// 초기화
const resetMsg = ref('')
const resetMsgType = ref('success')

// 메모
const memoMsg = ref('')
const restoreData = ref(null)

async function fetchUnmapped() {
  unmappedLoading.value = true
  try {
    const res = await api.getUnmappedModels()
    unmapped.value = res.data
  } catch (e) { console.error(e) } finally { unmappedLoading.value = false }
}

async function addMapping(modelName) {
  const chipset = chipsetInputs.value[modelName]?.trim()
  if (!chipset) return
  try {
    await api.addChipsetMapping(modelName, chipset)
    chipsetMsg.value = `${modelName} → ${chipset} 매핑 완료`
    chipsetMsgType.value = 'success'
    chipsetInputs.value[modelName] = ''
    fetchUnmapped()
  } catch (e) { chipsetMsg.value = '오류 발생'; chipsetMsgType.value = 'error' }
}

async function mergeChipsets() {
  try {
    await api.mergeChipsets(mergeSource.value, mergeTarget.value)
    chipsetMsg.value = `병합 완료: ${mergeSource.value} → ${mergeTarget.value}`
    chipsetMsgType.value = 'success'
    mergeSource.value = mergeTarget.value = ''
  } catch (e) { chipsetMsg.value = '오류 발생'; chipsetMsgType.value = 'error' }
}

async function renameChipset() {
  try {
    await api.renameChipset(renameOld.value, renameNew.value)
    chipsetMsg.value = `이름 변경: ${renameOld.value} → ${renameNew.value}`
    chipsetMsgType.value = 'success'
    renameOld.value = renameNew.value = ''
  } catch (e) { chipsetMsg.value = '오류 발생'; chipsetMsgType.value = 'error' }
}

async function updateWatchModels() {
  actionLoading.value = true
  try {
    const res = await axios.post('/api/model/update-watch')
    modelMsg.value = `완료: ${res.data.updated_count}건 업데이트`
    modelMsgType.value = 'success'
  } catch (e) { modelMsg.value = '오류 발생'; modelMsgType.value = 'error' } finally { actionLoading.value = false }
}

async function updateMapping() {
  actionLoading.value = true
  try {
    const res = await axios.post('/api/model/update-mapping')
    modelMsg.value = `완료: ${res.data.updated_count}건 업데이트`
    modelMsgType.value = 'success'
  } catch (e) { modelMsg.value = '오류 발생'; modelMsgType.value = 'error' } finally { actionLoading.value = false }
}

async function resetVoc() {
  if (!confirm('VOC 데이터를 모두 삭제하시겠습니까?')) return
  try {
    const res = await api.resetVocData()
    resetMsg.value = res.data.message
    resetMsgType.value = 'success'
  } catch (e) { resetMsg.value = '오류 발생'; resetMsgType.value = 'error' }
}

async function resetQData() {
  if (!confirm('Q-data를 모두 삭제하시겠습니까?')) return
  try {
    const res = await api.resetQData()
    resetMsg.value = `Q-data ${res.data.deleted_count}건 삭제 완료`
    resetMsgType.value = 'success'
  } catch (e) { resetMsg.value = '오류 발생'; resetMsgType.value = 'error' }
}

async function backupMemos() {
  try {
    const res = await axios.post('/api/memos/backup')
    const blob = new Blob([JSON.stringify(res.data, null, 2)], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url; a.download = `memos_backup_${new Date().toISOString().slice(0,10)}.json`
    a.click(); URL.revokeObjectURL(url)
  } catch (e) { console.error(e) }
}

function handleRestoreFile(e) {
  const file = e.target.files[0]
  if (!file) return
  const reader = new FileReader()
  reader.onload = (ev) => { restoreData.value = JSON.parse(ev.target.result) }
  reader.readAsText(file)
}

async function restoreMemos() {
  try {
    const res = await axios.post('/api/memos/restore', restoreData.value)
    memoMsg.value = `복원 완료: ${res.data.restored}건`
    restoreData.value = null
  } catch (e) { console.error(e) }
}

onMounted(fetchUnmapped)
</script>

<style scoped>
.page-title { font-size: 1.5rem; font-weight: 700; margin-bottom: 20px; color: #1a237e; }
.tabs { display: flex; gap: 8px; margin-bottom: 16px; flex-wrap: wrap; }
.tab-btn { padding: 8px 18px; border: 1px solid #c5cae9; border-radius: 20px; background: #fff; cursor: pointer; font-size: 0.9rem; }
.tab-btn.active { background: #1a237e; color: #fff; border-color: #1a237e; }
.card { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); margin-bottom: 16px; }
.section-title { font-size: 1rem; font-weight: 600; margin-bottom: 16px; color: #333; }
.danger-title { color: #b71c1c; }
.divider { height: 1px; background: #eee; margin: 20px 0; }
.table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
.table th { background: #f8f9ff; padding: 10px; text-align: left; font-weight: 600; color: #555; border-bottom: 2px solid #e8eaf6; }
.table td { padding: 9px 10px; border-bottom: 1px solid #f0f0f0; }
.badge { background: #e8eaf6; color: #1a237e; padding: 2px 10px; border-radius: 12px; font-weight: 600; font-size: 0.8rem; }
.inline-form { display: flex; gap: 8px; align-items: center; }
.sm-input { padding: 6px 10px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.85rem; width: 160px; }
.btn-sm-primary { padding: 6px 14px; background: #1a237e; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 0.82rem; }
.form-row { display: flex; align-items: center; gap: 16px; flex-wrap: wrap; }
.form-row label { display: flex; align-items: center; gap: 8px; font-size: 0.9rem; }
.form-input { padding: 7px 10px; border: 1px solid #ddd; border-radius: 6px; width: 180px; }
.btn-primary { padding: 8px 20px; background: #1a237e; color: #fff; border: none; border-radius: 6px; cursor: pointer; }
.btn-secondary { padding: 8px 20px; background: #fff; color: #1a237e; border: 1px solid #1a237e; border-radius: 6px; cursor: pointer; }
.btn-danger { padding: 8px 20px; background: #b71c1c; color: #fff; border: none; border-radius: 6px; cursor: pointer; }
.action-list { display: flex; flex-direction: column; gap: 16px; }
.action-item { display: flex; justify-content: space-between; align-items: center; padding: 16px; background: #f8f9ff; border-radius: 8px; gap: 16px; }
.action-item strong { display: block; margin-bottom: 4px; }
.action-item p { font-size: 0.85rem; color: #888; margin: 0; }
.warn-text { color: #e65100; font-size: 0.9rem; margin-bottom: 16px; }
.msg { margin-top: 14px; padding: 10px 16px; border-radius: 6px; font-size: 0.88rem; }
.msg.success { background: #e8f5e9; color: #2e7d32; }
.msg.error { background: #fce4ec; color: #b71c1c; }
.loading-sm { color: #888; font-size: 0.9rem; padding: 10px 0; }
.empty-sm { color: #aaa; font-size: 0.88rem; }
.mt-8 { margin-top: 8px; }
.hidden-input { display: none; }
</style>
