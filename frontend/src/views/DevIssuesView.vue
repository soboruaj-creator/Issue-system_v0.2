<template>
  <div class="dev-issues-view">
    <h1 class="page-title">개발이슈 Pending 관리</h1>

    <!-- Search -->
    <div class="card search-card">
      <div class="search-row">
        <input
          v-model="searchCode"
          class="search-input"
          placeholder="케이스 코드 입력 (예: P240101-001)"
          @keyup.enter="search"
        />
        <button class="btn-search" @click="search" :disabled="searching">
          {{ searching ? '검색 중...' : '검색' }}
        </button>
      </div>
      <div v-if="searchError" class="error-msg">{{ searchError }}</div>
    </div>

    <!-- Issue detail -->
    <div v-if="issue" class="card issue-card">

      <!-- Issue header -->
      <div class="issue-header">
        <div class="issue-meta-row">
          <span class="badge status" :class="issue.status">{{ issue.status.toUpperCase() }}</span>
          <span class="badge type" :class="issue.issue_type === 'UT' ? 'ut' : 'self'">{{ issue.issue_type }}</span>
          <span class="issue-code">{{ issue.case_code }}</span>
          <span v-if="issue.is_pending" class="badge pending-on">Pending</span>
        </div>
        <div class="issue-model">{{ issue.model_name || '미분류' }}</div>
        <div class="issue-title">{{ issue.title }}</div>
      </div>

      <hr class="divider" />

      <!-- Pending toggle + memo -->
      <div class="pending-section">
        <div class="pending-row">
          <span class="section-label">Pending 상태</span>
          <button
            class="btn-toggle"
            :class="issue.is_pending ? 'pending-active' : 'pending-inactive'"
            @click="togglePending"
            :disabled="saving"
          >
            {{ issue.is_pending ? 'Pending 해제' : 'Pending 설정' }}
          </button>
        </div>

        <div class="memo-wrap">
          <label class="section-label">메모</label>
          <textarea
            v-model="memo"
            class="memo-input"
            placeholder="Pending 관련 메모를 입력하세요..."
            rows="3"
          ></textarea>
          <button class="btn-save-memo" @click="saveMemo" :disabled="saving">
            {{ saving ? '저장 중...' : '메모 저장' }}
          </button>
        </div>
      </div>

      <hr class="divider" />

      <!-- Attachments -->
      <div class="attach-section">
        <div class="section-label">첨부파일</div>

        <div v-if="issue.pending_attachments?.length" class="attach-list">
          <div v-for="att in issue.pending_attachments" :key="att.stored_name" class="attach-item">
            <span class="attach-icon">📎</span>
            <span class="attach-filename">{{ att.filename }}</span>
            <span class="attach-date">{{ att.uploaded_at }}</span>
            <a :href="getAttachUrl(att.stored_name)" class="btn-download" target="_blank">다운로드</a>
            <button class="btn-delete-attach" @click="deleteAttach(att.stored_name)" :disabled="saving">삭제</button>
          </div>
        </div>
        <div v-else class="no-attach">첨부파일 없음</div>

        <!-- Upload zone -->
        <div
          class="drop-zone"
          :class="{ 'drag-over': dragging }"
          @dragover.prevent
          @drop.prevent="handleDrop"
          @dragenter="dragging = true"
          @dragleave="dragging = false"
        >
          <input type="file" ref="fileInput" @change="handleFileChange" class="hidden-input" />
          <div class="drop-content" @click="$refs.fileInput.click()">
            <span class="drop-icon">📎</span>
            <p>클릭하거나 파일을 드래그하세요</p>
            <p class="sub-text">{{ uploadFile?.name || '파일 미선택' }}</p>
          </div>
        </div>
        <button class="btn-upload-attach" @click="uploadAttach" :disabled="!uploadFile || uploading">
          {{ uploading ? '업로드 중...' : '첨부파일 업로드' }}
        </button>
        <div v-if="attachMsg" class="attach-msg" :class="attachMsgOk ? 'ok' : 'fail'">{{ attachMsg }}</div>
      </div>

    </div>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import {
  searchDevIssue,
  updateDevIssuePending,
  uploadDevIssueAttachment,
  deleteDevIssueAttachment,
  getDevIssueAttachmentUrl,
} from '../api'

const searchCode = ref('')
const searching = ref(false)
const searchError = ref('')
const issue = ref(null)
const memo = ref('')
const saving = ref(false)

const dragging = ref(false)
const fileInput = ref(null)
const uploadFile = ref(null)
const uploading = ref(false)
const attachMsg = ref('')
const attachMsgOk = ref(true)

async function search() {
  if (!searchCode.value.trim()) return
  searching.value = true
  searchError.value = ''
  issue.value = null
  try {
    const res = await searchDevIssue(searchCode.value.trim())
    issue.value = res.data
    memo.value = res.data.pending_memo || ''
  } catch (e) {
    searchError.value = e.response?.data?.detail || '이슈를 찾을 수 없습니다.'
  } finally {
    searching.value = false
  }
}

async function togglePending() {
  if (!issue.value) return
  saving.value = true
  try {
    const res = await updateDevIssuePending(issue.value.case_code, !issue.value.is_pending, memo.value)
    issue.value = res.data
    memo.value = res.data.pending_memo || ''
  } catch (e) {
    console.error(e)
  } finally {
    saving.value = false
  }
}

async function saveMemo() {
  if (!issue.value) return
  saving.value = true
  try {
    const res = await updateDevIssuePending(issue.value.case_code, issue.value.is_pending, memo.value)
    issue.value = res.data
  } catch (e) {
    console.error(e)
  } finally {
    saving.value = false
  }
}

function getAttachUrl(storedName) {
  return getDevIssueAttachmentUrl(issue.value.case_code, storedName)
}

function handleFileChange(e) {
  uploadFile.value = e.target.files[0] || null
}

function handleDrop(e) {
  dragging.value = false
  uploadFile.value = e.dataTransfer.files[0] || null
}

async function uploadAttach() {
  if (!issue.value || !uploadFile.value) return
  uploading.value = true
  attachMsg.value = ''
  try {
    const res = await uploadDevIssueAttachment(issue.value.case_code, uploadFile.value)
    if (!issue.value.pending_attachments) issue.value.pending_attachments = []
    issue.value.pending_attachments.push(res.data)
    uploadFile.value = null
    if (fileInput.value) fileInput.value.value = ''
    attachMsg.value = '업로드 완료'
    attachMsgOk.value = true
  } catch (e) {
    attachMsg.value = e.response?.data?.detail || '업로드 실패'
    attachMsgOk.value = false
  } finally {
    uploading.value = false
  }
}

async function deleteAttach(storedName) {
  if (!issue.value || !confirm('삭제하시겠습니까?')) return
  saving.value = true
  try {
    await deleteDevIssueAttachment(issue.value.case_code, storedName)
    issue.value.pending_attachments = issue.value.pending_attachments.filter(
      a => a.stored_name !== storedName
    )
  } catch (e) {
    console.error(e)
  } finally {
    saving.value = false
  }
}
</script>

<style scoped>
.dev-issues-view { max-width: 800px; margin: 0 auto; }
.page-title { font-size: 1.5rem; font-weight: 700; margin-bottom: 20px; color: #1a237e; }

.card { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); margin-bottom: 16px; }

.search-card { display: flex; flex-direction: column; gap: 8px; }
.search-row { display: flex; gap: 8px; }
.search-input { flex: 1; padding: 9px 14px; border: 1px solid #c5cae9; border-radius: 8px; font-size: 0.9rem; }
.search-input:focus { outline: none; border-color: #1a237e; }
.btn-search { padding: 9px 20px; background: #1a237e; color: #fff; border: none; border-radius: 8px; cursor: pointer; font-size: 0.9rem; font-weight: 600; }
.btn-search:disabled { opacity: 0.5; cursor: not-allowed; }
.error-msg { color: #b71c1c; font-size: 0.85rem; background: #fce4ec; padding: 8px 12px; border-radius: 6px; }

/* Issue header */
.issue-header { margin-bottom: 16px; }
.issue-meta-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-bottom: 8px; }
.badge { padding: 3px 10px; border-radius: 10px; font-weight: 700; font-size: 0.8rem; }
.badge.status.open { background: #e3f2fd; color: #1565c0; }
.badge.status.resolve { background: #e8f5e9; color: #2e7d32; }
.badge.status.close { background: #f5f5f5; color: #757575; }
.badge.type.ut { background: #fce4ec; color: #c62828; }
.badge.type.self { background: #f3e5f5; color: #6a1b9a; }
.badge.pending-on { background: #fff3e0; color: #e65100; }
.issue-code { font-size: 0.85rem; color: #888; font-family: monospace; }
.issue-model { font-size: 1rem; font-weight: 700; color: #1a237e; margin-bottom: 4px; }
.issue-title { font-size: 0.9rem; color: #444; line-height: 1.5; }

.divider { border: none; border-top: 1px solid #eee; margin: 16px 0; }

/* Pending */
.pending-section { display: flex; flex-direction: column; gap: 14px; }
.pending-row { display: flex; align-items: center; gap: 12px; }
.section-label { font-size: 0.88rem; font-weight: 600; color: #555; min-width: 80px; }
.btn-toggle { padding: 7px 18px; border: none; border-radius: 8px; cursor: pointer; font-size: 0.88rem; font-weight: 600; }
.btn-toggle.pending-active { background: #fff3e0; color: #e65100; }
.btn-toggle.pending-inactive { background: #e8eaf6; color: #1a237e; }
.btn-toggle:disabled { opacity: 0.5; cursor: not-allowed; }

.memo-wrap { display: flex; flex-direction: column; gap: 8px; }
.memo-input { padding: 10px 12px; border: 1px solid #c5cae9; border-radius: 8px; font-size: 0.88rem; resize: vertical; font-family: inherit; }
.memo-input:focus { outline: none; border-color: #1a237e; }
.btn-save-memo { align-self: flex-start; padding: 7px 20px; background: #1a237e; color: #fff; border: none; border-radius: 8px; cursor: pointer; font-size: 0.88rem; font-weight: 600; }
.btn-save-memo:disabled { opacity: 0.5; cursor: not-allowed; }

/* Attachments */
.attach-section { display: flex; flex-direction: column; gap: 10px; }
.attach-list { display: flex; flex-direction: column; gap: 6px; }
.attach-item { display: flex; align-items: center; gap: 8px; padding: 8px 12px; background: #f8f9ff; border-radius: 8px; font-size: 0.85rem; flex-wrap: wrap; }
.attach-icon { font-size: 1rem; }
.attach-filename { flex: 1; min-width: 120px; font-weight: 500; color: #333; word-break: break-all; }
.attach-date { font-size: 0.78rem; color: #aaa; white-space: nowrap; }
.btn-download { padding: 4px 12px; background: #e8f5e9; color: #2e7d32; border-radius: 6px; text-decoration: none; font-size: 0.8rem; font-weight: 600; }
.btn-delete-attach { padding: 4px 10px; background: #fce4ec; color: #b71c1c; border: none; border-radius: 6px; cursor: pointer; font-size: 0.8rem; font-weight: 600; }
.btn-delete-attach:disabled { opacity: 0.5; cursor: not-allowed; }
.no-attach { font-size: 0.85rem; color: #aaa; }

.drop-zone { border: 2px dashed #c5cae9; border-radius: 8px; padding: 20px; text-align: center; cursor: pointer; transition: border-color 0.2s, background 0.2s; }
.drop-zone.drag-over { border-color: #1a237e; background: #f0f2ff; }
.drop-content { pointer-events: none; }
.drop-icon { font-size: 1.5rem; }
.drop-content p { margin-top: 4px; font-size: 0.85rem; color: #555; }
.sub-text { color: #1a237e; font-weight: 500; margin-top: 2px !important; }
.hidden-input { display: none; }
.btn-upload-attach { padding: 9px 20px; background: #283593; color: #fff; border: none; border-radius: 8px; cursor: pointer; font-size: 0.88rem; font-weight: 600; align-self: flex-start; }
.btn-upload-attach:disabled { opacity: 0.5; cursor: not-allowed; }
.attach-msg { font-size: 0.85rem; padding: 6px 12px; border-radius: 6px; }
.attach-msg.ok { background: #e8f5e9; color: #2e7d32; }
.attach-msg.fail { background: #fce4ec; color: #b71c1c; }
</style>
