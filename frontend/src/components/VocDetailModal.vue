<template>
  <div class="modal-overlay" @click.self="$emit('close')">
    <div class="modal">
      <div class="modal-header">
        <h2>VOC 상세 정보</h2>
        <button class="close-btn" @click="$emit('close')">✕</button>
      </div>
      <div class="modal-body">
        <div class="detail-grid">
          <div class="detail-row"><span class="label">사례코드</span><span class="value code">{{ voc.case_code }}</span></div>
          <div class="detail-row"><span class="label">모델명</span><span class="value">{{ voc.model_name || '-' }}</span></div>
          <div class="detail-row"><span class="label">칩셋</span><span class="value">{{ voc.chipset || '-' }}</span></div>
          <div class="detail-row"><span class="label">빌드 버전</span><span class="value">{{ voc.build_version || '-' }}</span></div>
          <div class="detail-row"><span class="label">OS 버전</span><span class="value">{{ voc.os_version || '-' }}</span></div>
          <div class="detail-row"><span class="label">이슈 유형</span>
            <span :class="['value', 'badge', voc.issue_type === '내부이슈' ? 'internal' : 'external']">
              {{ voc.issue_type || '-' }}
            </span>
          </div>
          <div class="detail-row"><span class="label">3rd Party 앱</span><span class="value">{{ voc.third_party_app || '-' }}</span></div>
          <div class="detail-row"><span class="label">등록일</span><span class="value">{{ voc.created_date || '-' }}</span></div>
        </div>

        <div class="section" v-if="voc.problem">
          <h3>문제 현상</h3>
          <p class="content-text">{{ voc.problem }}</p>
        </div>
        <div class="section" v-if="voc.original_content">
          <h3>원문 내용</h3>
          <p class="content-text">{{ voc.original_content }}</p>
        </div>
        <div class="section" v-if="voc.reproduction_path">
          <h3>재현 경로</h3>
          <pre class="content-pre">{{ voc.reproduction_path }}</pre>
        </div>
        <div class="section" v-if="voc.cause || voc.solution">
          <h3>원인 / 대책</h3>
          <p v-if="voc.cause"><strong>원인:</strong> {{ voc.cause }}</p>
          <p v-if="voc.solution"><strong>대책:</strong> {{ voc.solution }}</p>
        </div>

        <!-- 댓글 -->
        <div class="section comments-section">
          <h3>댓글</h3>
          <div v-if="voc.comments?.length" class="comments-list">
            <div v-for="c in voc.comments" :key="c._id" class="comment-item">
              <p class="comment-text">{{ c.comment }}</p>
              <span class="comment-date">{{ c.created_date }}</span>
            </div>
          </div>
          <p v-else class="no-comments">댓글이 없습니다.</p>
          <div class="comment-input-row">
            <input v-model="newComment" placeholder="댓글 입력..." @keyup.enter="addComment" />
            <button @click="addComment" :disabled="!newComment.trim()">등록</button>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { getVocDetail, addVocComment } from '../api'

const props = defineProps({ voc: Object })
const emit = defineEmits(['close'])

const detail = ref(null)
const newComment = ref('')

onMounted(async () => {
  try {
    const res = await getVocDetail(props.voc.case_code)
    detail.value = res.data
    // 댓글 반영
    props.voc.comments = res.data.comments
  } catch (e) { console.error(e) }
})

async function addComment() {
  if (!newComment.value.trim()) return
  try {
    await addVocComment(props.voc.case_code, newComment.value.trim())
    newComment.value = ''
    const res = await getVocDetail(props.voc.case_code)
    props.voc.comments = res.data.comments
  } catch (e) { console.error(e) }
}
</script>

<style scoped>
.modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; z-index: 1000; padding: 20px; }
.modal { background: #fff; border-radius: 12px; width: 100%; max-width: 760px; max-height: 90vh; display: flex; flex-direction: column; }
.modal-header { display: flex; justify-content: space-between; align-items: center; padding: 18px 24px; border-bottom: 1px solid #eee; }
.modal-header h2 { font-size: 1.1rem; font-weight: 700; color: #1a237e; }
.close-btn { background: none; border: none; font-size: 1.2rem; cursor: pointer; color: #888; }
.modal-body { overflow-y: auto; padding: 20px 24px; }
.detail-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 20px; }
.detail-row { display: flex; align-items: center; gap: 10px; }
.label { font-size: 0.8rem; color: #888; min-width: 80px; }
.value { font-size: 0.9rem; color: #333; }
.value.code { font-family: monospace; background: #e8eaf6; padding: 2px 6px; border-radius: 4px; }
.badge { padding: 2px 10px; border-radius: 10px; font-size: 0.8rem; }
.badge.internal { background: #e8eaf6; color: #3949ab; }
.badge.external { background: #fff3e0; color: #e65100; }
.section { margin-bottom: 18px; }
.section h3 { font-size: 0.9rem; font-weight: 600; color: #555; margin-bottom: 8px; border-bottom: 1px solid #eee; padding-bottom: 4px; }
.content-text { font-size: 0.88rem; color: #444; line-height: 1.6; white-space: pre-wrap; }
.content-pre { font-size: 0.82rem; color: #444; line-height: 1.6; white-space: pre-wrap; background: #f8f9ff; padding: 12px; border-radius: 6px; }
.comments-list { margin-bottom: 12px; }
.comment-item { background: #f8f9ff; border-radius: 6px; padding: 10px 14px; margin-bottom: 8px; }
.comment-text { font-size: 0.88rem; color: #333; }
.comment-date { font-size: 0.78rem; color: #aaa; margin-top: 4px; display: block; }
.no-comments { font-size: 0.85rem; color: #aaa; margin-bottom: 12px; }
.comment-input-row { display: flex; gap: 8px; }
.comment-input-row input { flex: 1; padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88rem; }
.comment-input-row button { padding: 8px 18px; background: #1a237e; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 0.85rem; }
.comment-input-row button:disabled { opacity: 0.5; cursor: not-allowed; }
</style>
