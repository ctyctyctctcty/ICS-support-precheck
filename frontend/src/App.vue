<script setup>
import { computed, onMounted, ref } from 'vue'
import workflowImage from './assets/workflow.svg'

const apiBase = import.meta.env.VITE_API_BASE || ''
const files = ref([])
const results = ref([])
const folders = ref({})
const errorMessage = ref('')
const busy = ref(false)
const dragging = ref(false)

const canSubmit = computed(() => files.value.length > 0 && !busy.value)

function keepXlsx(fileList) {
  return Array.from(fileList || []).filter((file) => {
    return file.name.toLowerCase().endsWith('.xlsx') && !file.name.startsWith('~$')
  })
}

function addFiles(fileList) {
  errorMessage.value = ''
  const incoming = keepXlsx(fileList)
  const known = new Set(files.value.map((file) => `${file.name}:${file.size}:${file.lastModified}`))
  for (const file of incoming) {
    const key = `${file.name}:${file.size}:${file.lastModified}`
    if (!known.has(key)) {
      files.value.push(file)
      known.add(key)
    }
  }
  if (!incoming.length) {
    errorMessage.value = '通常の .xlsx ファイルを選択してください。'
  }
}

function onDrop(event) {
  dragging.value = false
  addFiles(event.dataTransfer.files)
}

function removeFile(index) {
  files.value.splice(index, 1)
}

function clearAll() {
  files.value = []
  results.value = []
  errorMessage.value = ''
}

async function loadHealth() {
  try {
    const response = await fetch(`${apiBase}/api/health`)
    if (!response.ok) return
    const data = await response.json()
    folders.value = data.folders || {}
  } catch {
    // The upload action will show the useful error if the backend is unavailable.
  }
}

async function submit() {
  if (!canSubmit.value) return
  busy.value = true
  errorMessage.value = ''
  results.value = []

  const body = new FormData()
  files.value.forEach((file) => body.append('files', file))

  try {
    const response = await fetch(`${apiBase}/api/process`, {
      method: 'POST',
      body,
    })
    const data = await response.json()
    if (!response.ok) {
      throw new Error(data.detail || '処理に失敗しました。')
    }
    results.value = data.results || []
    files.value = []
    await loadHealth()
  } catch (error) {
    errorMessage.value = error.message || '処理に失敗しました。'
  } finally {
    busy.value = false
  }
}

function statusLabel(status) {
  return {
    network_ready: 'Network ready',
    needs_confirmation: 'Needs confirmation',
    error: 'Error',
  }[status] || status
}

function statusText(status) {
  return {
    network_ready: 'Network team へ連携できます。',
    needs_confirmation: '申請者または管理者への確認が必要です。',
    error: '不備があります。レポートを確認してください。',
  }[status] || ''
}

onMounted(loadHealth)
</script>

<template>
  <main class="shell">
    <section class="topline">
      <div>
        <p class="eyebrow">ICS Support Precheck</p>
        <h1>申請書をアップロード</h1>
        <p class="lead">Excel 申請書を標準形式へ変換し、結果を三つのフォルダへ振り分けます。</p>
      </div>
      <img class="workflow" :src="workflowImage" alt="source, check, output workflow" />
    </section>

    <section
      class="upload-area"
      :class="{ active: dragging }"
      @dragover.prevent="dragging = true"
      @dragleave.prevent="dragging = false"
      @drop.prevent="onDrop"
    >
      <div>
        <h2>.xlsx をここへドロップ</h2>
        <p>複数ファイルをまとめて処理できます。</p>
      </div>
      <label class="file-button">
        ファイルを選択
        <input type="file" accept=".xlsx" multiple @change="addFiles($event.target.files)" />
      </label>
    </section>

    <section v-if="files.length" class="queue">
      <div class="section-title">
        <h2>処理待ち</h2>
        <button type="button" class="link-button" @click="clearAll">クリア</button>
      </div>
      <ul class="file-list">
        <li v-for="(file, index) in files" :key="`${file.name}-${file.size}-${index}`">
          <span>{{ file.name }}</span>
          <button type="button" @click="removeFile(index)">削除</button>
        </li>
      </ul>
    </section>

    <div class="actions">
      <button type="button" class="primary" :disabled="!canSubmit" @click="submit">
        {{ busy ? '処理中...' : 'チェック開始' }}
      </button>
      <p v-if="errorMessage" class="error-text">{{ errorMessage }}</p>
    </div>

    <section v-if="results.length" class="results">
      <h2>処理結果</h2>
      <article v-for="result in results" :key="result.original_name" class="result-row" :data-status="result.status">
        <div>
          <p class="file-name">{{ result.original_name }}</p>
          <p class="status-line">{{ statusLabel(result.status) }} - {{ statusText(result.status) }}</p>
          <p class="folder-path">{{ result.folder }}</p>
        </div>
        <div class="downloads">
          <a
            v-for="file in result.files"
            :key="file.name"
            :href="`${apiBase}${file.download_url}`"
          >
            {{ file.name }}
          </a>
        </div>
      </article>
    </section>

    <section class="folders">
      <h2>出力フォルダ</h2>
      <dl>
        <div>
          <dt>Network ready</dt>
          <dd>{{ folders.network_ready || 'data/network_ready' }}</dd>
        </div>
        <div>
          <dt>Needs confirmation</dt>
          <dd>{{ folders.needs_confirmation || 'data/needs_confirmation' }}</dd>
        </div>
        <div>
          <dt>Error</dt>
          <dd>{{ folders.error || 'data/error' }}</dd>
        </div>
      </dl>
    </section>
  </main>
</template>
