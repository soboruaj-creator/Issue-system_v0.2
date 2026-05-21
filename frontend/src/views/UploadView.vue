<template>
  <div class="upload-view">
    <h1 class="page-title">파일 업로드</h1>

    <div class="upload-grid">
      <!-- VOC 업로드 -->
      <div class="card upload-card">
        <h2>📄 Members issue 업로드</h2>
        <p class="desc">VOC 엑셀 파일(.xlsx/.xls)을 업로드합니다.</p>
        <div class="drop-zone" @dragover.prevent @drop.prevent="handleDrop($event, 'voc')"
             :class="{ 'drag-over': dragging === 'voc' }"
             @dragenter="dragging='voc'" @dragleave="dragging=null">
          <input type="file" accept=".xlsx,.xls" @change="handleFile($event, 'voc')" ref="vocInput" class="hidden-input" />
          <div class="drop-content" @click="$refs.vocInput.click()">
            <span class="drop-icon">📂</span>
            <p>클릭하거나 파일을 드래그하세요</p>
            <p class="sub-text">{{ vocFile?.name || '파일 미선택' }}</p>
          </div>
        </div>
        <button class="btn-upload" @click="upload('voc')" :disabled="!vocFile || vocLoading">
          {{ vocLoading ? '업로드 중...' : '업로드' }}
        </button>
        <div v-if="vocResult" :class="['result', vocResult.success ? 'success' : 'error']">
          {{ vocResult.message }}
          <div v-if="vocResult.unmapped_models?.length" class="unmapped">
            미매핑 모델: {{ vocResult.unmapped_models.join(', ') }}
          </div>
        </div>
      </div>

      <!-- 칩셋 매핑 업로드 -->
      <div class="card upload-card">
        <h2>🔧 칩셋 매핑 업로드</h2>
        <p class="desc">모델명/칩셋 매핑 엑셀 파일을 업로드합니다.</p>
        <div class="drop-zone" @dragover.prevent @drop.prevent="handleDrop($event, 'chipset')"
             :class="{ 'drag-over': dragging === 'chipset' }"
             @dragenter="dragging='chipset'" @dragleave="dragging=null">
          <input type="file" accept=".xlsx,.xls" @change="handleFile($event, 'chipset')" ref="chipsetInput" class="hidden-input" />
          <div class="drop-content" @click="$refs.chipsetInput.click()">
            <span class="drop-icon">📂</span>
            <p>클릭하거나 파일을 드래그하세요</p>
            <p class="sub-text">{{ chipsetFile?.name || '파일 미선택' }}</p>
          </div>
        </div>
        <button class="btn-upload" @click="upload('chipset')" :disabled="!chipsetFile || chipsetLoading">
          {{ chipsetLoading ? '업로드 중...' : '업로드' }}
        </button>
        <div v-if="chipsetResult" :class="['result', chipsetResult.success ? 'success' : 'error']">
          {{ chipsetResult.message }}
        </div>
      </div>

      <!-- 앱 키워드 업로드 -->
      <div class="card upload-card">
        <h2>📱 앱 키워드 업로드</h2>
        <p class="desc">3rd party 앱 키워드 엑셀 파일을 업로드합니다.</p>
        <div class="drop-zone" @dragover.prevent @drop.prevent="handleDrop($event, 'app')"
             :class="{ 'drag-over': dragging === 'app' }"
             @dragenter="dragging='app'" @dragleave="dragging=null">
          <input type="file" accept=".xlsx,.xls" @change="handleFile($event, 'app')" ref="appInput" class="hidden-input" />
          <div class="drop-content" @click="$refs.appInput.click()">
            <span class="drop-icon">📂</span>
            <p>클릭하거나 파일을 드래그하세요</p>
            <p class="sub-text">{{ appFile?.name || '파일 미선택' }}</p>
          </div>
        </div>
        <button class="btn-upload" @click="upload('app')" :disabled="!appFile || appLoading">
          {{ appLoading ? '업로드 중...' : '업로드' }}
        </button>
        <div v-if="appResult" :class="['result', appResult.success ? 'success' : 'error']">
          {{ appResult.message }}
        </div>
      </div>

      <!-- 출시일 업로드 -->
      <div class="card upload-card">
        <h2>📅 출시일 업로드</h2>
        <p class="desc">모델 출시일 엑셀 파일을 업로드합니다. (A열: 모델명, B열: 출시일)</p>
        <div class="drop-zone" @dragover.prevent @drop.prevent="handleDrop($event, 'launch')"
             :class="{ 'drag-over': dragging === 'launch' }"
             @dragenter="dragging='launch'" @dragleave="dragging=null">
          <input type="file" accept=".xlsx,.xls" @change="handleFile($event, 'launch')" ref="launchInput" class="hidden-input" />
          <div class="drop-content" @click="$refs.launchInput.click()">
            <span class="drop-icon">📂</span>
            <p>클릭하거나 파일을 드래그하세요</p>
            <p class="sub-text">{{ launchFile?.name || '파일 미선택' }}</p>
          </div>
        </div>
        <button class="btn-upload" @click="upload('launch')" :disabled="!launchFile || launchLoading">
          {{ launchLoading ? '업로드 중...' : '업로드' }}
        </button>
        <div v-if="launchResult" :class="['result', launchResult.success ? 'success' : 'error']">
          {{ launchResult.message }}
        </div>
      </div>

      <!-- Q-data 업로드 -->
      <div class="card upload-card">
        <h2>📊 Q-data 업로드</h2>
        <p class="desc">서비스 Q-data 엑셀 파일을 업로드합니다.</p>
        <div class="ppm-input-wrap">
          <label class="ppm-label">PPM (선택)</label>
          <input type="number" v-model="qdataPpm" class="ppm-input" placeholder="예: 1234.5" step="0.1" min="0" />
          <span class="ppm-hint">웹사이트에서 조회한 PPM 값을 입력하세요</span>
        </div>
        <div class="drop-zone" @dragover.prevent @drop.prevent="handleDrop($event, 'qdata')"
             :class="{ 'drag-over': dragging === 'qdata' }"
             @dragenter="dragging='qdata'" @dragleave="dragging=null">
          <input type="file" accept=".xlsx,.xls" @change="handleFile($event, 'qdata')" ref="qdataInput" class="hidden-input" />
          <div class="drop-content" @click="$refs.qdataInput.click()">
            <span class="drop-icon">📂</span>
            <p>클릭하거나 파일을 드래그하세요</p>
            <p class="sub-text">{{ qdataFile?.name || '파일 미선택' }}</p>
          </div>
        </div>
        <button class="btn-upload" @click="uploadQdataWithPpm" :disabled="!qdataFile || qdataLoading">
          {{ qdataLoading ? '업로드 중...' : '업로드' }}
        </button>
        <div v-if="qdataResult" :class="['result', qdataResult.success ? 'success' : 'error']">
          {{ qdataResult.message }}
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import { uploadVoc, uploadChipsetMapping, uploadAppKeywords, uploadQData, uploadLaunchDates } from '../api'

const dragging = ref(null)

const vocFile = ref(null), vocLoading = ref(false), vocResult = ref(null)
const chipsetFile = ref(null), chipsetLoading = ref(false), chipsetResult = ref(null)
const appFile = ref(null), appLoading = ref(false), appResult = ref(null)
const qdataFile = ref(null), qdataLoading = ref(false), qdataResult = ref(null), qdataPpm = ref('')
const launchFile = ref(null), launchLoading = ref(false), launchResult = ref(null)

function handleFile(e, type) {
  const file = e.target.files[0]
  setFile(type, file)
}

function handleDrop(e, type) {
  dragging.value = null
  const file = e.dataTransfer.files[0]
  setFile(type, file)
}

function setFile(type, file) {
  if (type === 'voc') { vocFile.value = file; vocResult.value = null }
  if (type === 'chipset') { chipsetFile.value = file; chipsetResult.value = null }
  if (type === 'app') { appFile.value = file; appResult.value = null }
  if (type === 'qdata') { qdataFile.value = file; qdataResult.value = null }
  if (type === 'launch') { launchFile.value = file; launchResult.value = null }
}

async function upload(type) {
  const apis = { voc: [vocFile, vocLoading, vocResult, uploadVoc],
                 chipset: [chipsetFile, chipsetLoading, chipsetResult, uploadChipsetMapping],
                 app: [appFile, appLoading, appResult, uploadAppKeywords],
                 launch: [launchFile, launchLoading, launchResult, uploadLaunchDates] }
  const [fileRef, loadingRef, resultRef, apiFn] = apis[type]

  loadingRef.value = true
  resultRef.value = null
  try {
    const res = await apiFn(fileRef.value)
    resultRef.value = res.data
  } catch (e) {
    resultRef.value = { success: false, message: e.response?.data?.detail || '업로드 실패' }
  } finally {
    loadingRef.value = false
  }
}

async function uploadQdataWithPpm() {
  qdataLoading.value = true
  qdataResult.value = null
  try {
    const ppm = qdataPpm.value !== '' ? parseFloat(qdataPpm.value) : null
    const res = await uploadQData(qdataFile.value, ppm)
    qdataResult.value = res.data
  } catch (e) {
    qdataResult.value = { success: false, message: e.response?.data?.detail || '업로드 실패' }
  } finally {
    qdataLoading.value = false
  }
}
</script>

<style scoped>
.page-title { font-size: 1.5rem; font-weight: 700; margin-bottom: 20px; color: #1a237e; }
.upload-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 20px; }
.card { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }
.upload-card h2 { font-size: 1rem; font-weight: 600; margin-bottom: 8px; }
.desc { font-size: 0.85rem; color: #888; margin-bottom: 16px; }
.drop-zone { border: 2px dashed #c5cae9; border-radius: 8px; padding: 24px; text-align: center; cursor: pointer; transition: border-color 0.2s, background 0.2s; }
.drop-zone.drag-over { border-color: #1a237e; background: #f0f2ff; }
.drop-content { pointer-events: none; }
.drop-icon { font-size: 2rem; }
.drop-content p { margin-top: 6px; font-size: 0.85rem; color: #555; }
.sub-text { color: #1a237e; font-weight: 500; margin-top: 4px !important; }
.hidden-input { display: none; }
.btn-upload { margin-top: 14px; width: 100%; padding: 10px; background: #1a237e; color: #fff; border: none; border-radius: 6px; font-size: 0.9rem; cursor: pointer; transition: opacity 0.2s; }
.btn-upload:disabled { opacity: 0.5; cursor: not-allowed; }
.result { margin-top: 12px; padding: 10px 14px; border-radius: 6px; font-size: 0.85rem; white-space: pre-line; }
.result.success { background: #e8f5e9; color: #2e7d32; }
.result.error { background: #fce4ec; color: #b71c1c; }
.unmapped { margin-top: 6px; font-size: 0.8rem; color: #e65100; }
.ppm-input-wrap { display: flex; flex-direction: column; gap: 4px; margin-bottom: 12px; }
.ppm-label { font-size: 0.85rem; font-weight: 600; color: #555; }
.ppm-input { padding: 7px 10px; border: 1px solid #c5cae9; border-radius: 6px; font-size: 0.9rem; width: 160px; }
.ppm-hint { font-size: 0.78rem; color: #999; }
</style>
