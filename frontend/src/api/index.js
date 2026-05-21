import axios from 'axios'

const api = axios.create({
  baseURL: '/api',
  timeout: 60000,
})

// ─── 업로드 ────────────────────────────────────────────────────────────────────

export const uploadVoc = (file) => {
  const form = new FormData()
  form.append('file', file)
  return api.post('/upload/internal_voc', form)
}

export const uploadChipsetMapping = (file) => {
  const form = new FormData()
  form.append('file', file)
  return api.post('/upload/chipset_mapping', form)
}

export const uploadAppKeywords = (file) => {
  const form = new FormData()
  form.append('file', file)
  return api.post('/upload/app_keywords', form)
}

export const uploadQData = (file) => {
  const form = new FormData()
  form.append('file', file)
  return api.post('/upload/qdata', form)
}

// ─── VOC ──────────────────────────────────────────────────────────────────────

export const listVoc = (params = {}) => api.get('/voc', { params })
export const getVocDetail = (caseCode) => api.get(`/voc/${caseCode}`)
export const addVocComment = (caseCode, comment) =>
  api.post(`/voc/${caseCode}/comment`, { comment })

// ─── 통계 ─────────────────────────────────────────────────────────────────────

export const getDashboard = () => api.get('/statistics/dashboard')
export const getModelStats = (params = {}) => api.get('/statistics/model', { params })
export const getWeeklyStats = (params = {}) => api.get('/statistics/weekly', { params })
export const getMonthlyStats = (params = {}) => api.get('/statistics/monthly', { params })
export const getChipsetStats = (params = {}) => api.get('/statistics/chipset', { params })
export const getAppStats = (params = {}) => api.get('/statistics/app', { params })
export const getModelMonthlyStats = (modelName) =>
  api.get(`/statistics/model/${encodeURIComponent(modelName)}/monthly`)
export const getModelsMonthlyStats = (models) =>
  api.post('/statistics/models/monthly', { models })

// ─── Q-data ───────────────────────────────────────────────────────────────────

export const listQData = (params = {}) => api.get('/qdata', { params })
export const getQDataModelStats = (params = {}) => api.get('/statistics/qdata/model', { params })
export const getQDataWeeklyStats = (params = {}) => api.get('/statistics/qdata/weekly', { params })
export const getQDataMonthlyStats = (params = {}) => api.get('/statistics/qdata/monthly', { params })
export const checkQDataDuplicates = () => api.get('/qdata/check-duplicates')
export const removeQDataDuplicates = () => api.post('/qdata/remove-duplicates')
export const resetQData = () => api.post('/qdata/reset')

// ─── 메모 ─────────────────────────────────────────────────────────────────────

export const getMonthlyMemos = () => api.get('/monthly-memos')
export const saveMonthlyMemo = (month, memo) => api.post('/monthly-memos', { month, memo })
export const deleteMonthlyMemo = (month) => api.delete(`/monthly-memos/${month}`)

export const getWeeklyMemos = () => api.get('/weekly-memos')
export const saveWeeklyMemo = (week, memo) => api.post('/weekly-memos', { week, memo })
export const deleteWeeklyMemo = (week) => api.delete(`/weekly-memos/${week}`)

export const getModelMonthlyMemos = (modelName) =>
  api.get(`/model-monthly-memos/${encodeURIComponent(modelName)}`)
export const saveModelMonthlyMemo = (modelName, month, memo) =>
  api.post('/model-monthly-memos', { model_name: modelName, month, memo })

// ─── 칩셋 ─────────────────────────────────────────────────────────────────────

export const getUnmappedModels = () => api.get('/unmapped-models')
export const addChipsetMapping = (modelName, chipset) =>
  api.post('/chipset-mapping/add', { model_name: modelName, chipset })
export const listChipsetMappings = () => api.get('/chipset-mappings')
export const mergeChipsets = (source, target) =>
  api.post('/chipset/merge', { source_chipset: source, target_chipset: target })
export const renameChipset = (oldName, newName) =>
  api.post('/chipset/rename', { old_name: oldName, new_name: newName })

// ─── 내보내기 ─────────────────────────────────────────────────────────────────

export const exportVocExcel = (params = {}) => {
  const query = new URLSearchParams(params).toString()
  window.location.href = `/api/export/excel${query ? '?' + query : ''}`
}

export const exportQDataExcel = (params = {}) => {
  const query = new URLSearchParams(params).toString()
  window.location.href = `/api/export/qdata/excel${query ? '?' + query : ''}`
}

// ─── 초기화 ───────────────────────────────────────────────────────────────────

export const resetVocData = () => api.post('/qdata/reset-all')

// ─── Q-data 통계 (모델별 월별) ────────────────────────────────────────────────

export const getQDataModelsMonthlyStats = (models) =>
  api.post('/statistics/qdata/models/monthly', { models })

// ─── 마케팅명 통합 통계 ───────────────────────────────────────────────────────

export const getEffectiveModels = () => api.get('/statistics/effective-models')

export const getEffectiveNameMonthly = (name, params = {}) =>
  api.get('/statistics/effective-name/monthly', { params: { name, ...params } })

export const getMonthDetail = (month) =>
  api.get(`/statistics/month-detail/${month}`)

// ─── 출시일/개통일 비교 ───────────────────────────────────────────────────────

export const getLaunchModels = () => api.get('/launch/models')

export const compareLaunchModels = (models) =>
  api.get('/launch/compare', { params: { models: models.join(',') } })

// ─── 모델 노트 ────────────────────────────────────────────────────────────────

export const getModelNotes = (modelName) =>
  api.get(`/model-notes/${encodeURIComponent(modelName)}`)

export const createModelNote = (modelName, data) =>
  api.post(`/model-notes/${encodeURIComponent(modelName)}`, data)

export const updateModelNote = (noteId, data) =>
  api.put(`/model-notes/${noteId}`, data)

export const deleteModelNote = (noteId) =>
  api.delete(`/model-notes/${noteId}`)

export const uploadLaunchDates = (file) => {
  const form = new FormData()
  form.append('file', file)
  return api.post('/upload/launch_dates', form)
}
