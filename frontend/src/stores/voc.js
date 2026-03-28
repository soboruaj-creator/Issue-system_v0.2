import { defineStore } from 'pinia'
import { ref, computed } from 'vue'
import * as api from '../api'

export const useVocStore = defineStore('voc', () => {
  const vocList = ref([])
  const totalVoc = ref(0)
  const currentPage = ref(1)
  const pageSize = ref(20)
  const loading = ref(false)
  const error = ref(null)

  const dashboardData = ref(null)
  const modelStats = ref([])
  const monthlyStats = ref([])
  const weeklyStats = ref([])
  const chipsetStats = ref([])
  const appStats = ref([])

  async function fetchVocList(params = {}) {
    loading.value = true
    error.value = null
    try {
      const res = await api.listVoc({ page: currentPage.value, size: pageSize.value, ...params })
      vocList.value = res.data.items
      totalVoc.value = res.data.total
    } catch (e) {
      error.value = e.response?.data?.detail || e.message
    } finally {
      loading.value = false
    }
  }

  async function fetchDashboard() {
    try {
      const res = await api.getDashboard()
      dashboardData.value = res.data
    } catch (e) {
      console.error(e)
    }
  }

  async function fetchModelStats(params = {}) {
    try {
      const res = await api.getModelStats(params)
      modelStats.value = res.data
    } catch (e) {
      console.error(e)
    }
  }

  async function fetchMonthlyStats(params = {}) {
    try {
      const res = await api.getMonthlyStats(params)
      monthlyStats.value = res.data
    } catch (e) {
      console.error(e)
    }
  }

  async function fetchWeeklyStats(params = {}) {
    try {
      const res = await api.getWeeklyStats(params)
      weeklyStats.value = res.data
    } catch (e) {
      console.error(e)
    }
  }

  async function fetchChipsetStats(params = {}) {
    try {
      const res = await api.getChipsetStats(params)
      chipsetStats.value = res.data
    } catch (e) {
      console.error(e)
    }
  }

  async function fetchAppStats(params = {}) {
    try {
      const res = await api.getAppStats(params)
      appStats.value = res.data
    } catch (e) {
      console.error(e)
    }
  }

  return {
    vocList, totalVoc, currentPage, pageSize, loading, error,
    dashboardData, modelStats, monthlyStats, weeklyStats, chipsetStats, appStats,
    fetchVocList, fetchDashboard, fetchModelStats, fetchMonthlyStats,
    fetchWeeklyStats, fetchChipsetStats, fetchAppStats,
  }
})
