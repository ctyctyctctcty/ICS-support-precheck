<script setup>
import { computed, onMounted, onUnmounted, ref } from 'vue'
import workflowImage from './assets/workflow.svg'

const apiBase = import.meta.env.VITE_API_BASE || ''
const files = ref([])
const results = ref([])
const folders = ref({})
const errorMessage = ref('')
const busy = ref(false)
const dragging = ref(false)
const progress = ref(0)
const progressMessage = ref('待機中')
let progressTimer = null

const processSteps = [
  { title: '申請書を読み込み', text: 'アップロードされた Excel から申請内容を抽出します。' },
  { title: '入力内容を確認', text: '必須項目、IPアドレス、ホスト名の形式を確認します。' },
  { title: '社内情報と照合', text: 'AD、DNS、DHCP参照ファイルを使って確認事項を整理します。' },
  { title: '標準形式へ変換', text: 'Network team へ渡せる標準 xlsx を作成します。' },
  { title: '結果を振り分け', text: '連携可能、確認待ち、要修正のフォルダへ保存します。' },
]

const canSubmit = computed(() => files.value.length > 0 && !busy.value)
const totalSize = computed(() => {
  const bytes = files.value.reduce((sum, file) => sum + file.size, 0)
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`
})

const currentStep = computed(() => {
  if (!busy.value && progress.value === 100) return processSteps.length
  if (!busy.value) return 0
  return Math.min(processSteps.length, Math.max(1, Math.ceil(progress.value / 20)))
})

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
  progress.value = 0
  progressMessage.value = '待機中'
}

function startProgress() {
  stopProgress()
  progress.value = 6
  progressMessage.value = '申請書を確認しています。'
  progressTimer = window.setInterval(() => {
    if (progress.value < 88) {
      progress.value += Math.max(1, Math.round((92 - progress.value) / 12))
    }
    if (progress.value >= 70) {
      progressMessage.value = '結果ファイルを作成しています。'
    } else if (progress.value >= 38) {
      progressMessage.value = 'AD、DNS、DHCP情報を照合しています。'
    }
  }, 500)
}

function finishProgress() {
  stopProgress()
  progress.value = 100
  progressMessage.value = '処理が完了しました。'
}

function stopProgress() {
  if (progressTimer) {
    window.clearInterval(progressTimer)
    progressTimer = null
  }
}

async function loadHealth() {
  try {
    const response = await fetch(`${apiBase}/api/health`)
    if (!response.ok) return
    const data = await response.json()
    folders.value = data.folders || {}
  } catch {
    // Upload errors are shown when the user starts processing.
  }
}

async function submit() {
  if (!canSubmit.value) return
  busy.value = true
  errorMessage.value = ''
  results.value = []
  startProgress()

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
    finishProgress()
    await loadHealth()
  } catch (error) {
    stopProgress()
    progressMessage.value = '処理を完了できませんでした。'
    errorMessage.value = error.message || '処理に失敗しました。'
  } finally {
    busy.value = false
  }
}

function statusLabel(status) {
  return {
    network_ready: '連携可能',
    needs_confirmation: '確認待ち',
    error: '要修正',
  }[status] || status
}

function statusText(status) {
  return {
    network_ready: '標準形式への変換が完了しました。Network team へ連携できます。',
    needs_confirmation: '変換は完了しました。申請者または管理者への確認が必要です。',
    error: '不備があります。出力されたレポートを確認してください。',
  }[status] || ''
}

function folderLabel(key) {
  return {
    network_ready: '連携可能',
    needs_confirmation: '確認待ち',
    error: '要修正',
  }[key] || key
}

onMounted(loadHealth)
onUnmounted(stopProgress)
</script>

<template>
  <main class="shell">
    <section class="hero">
      <div class="hero-copy">
        <p class="eyebrow">申請書チェックツール</p>
        <h1>VPN申請書を確認して標準形式へ変換</h1>
        <p class="lead">
          Excel申請書をアップロードすると、確認結果に応じて
          「連携可能」「確認待ち」「要修正」の出力先へ自動で振り分けます。
        </p>
      </div>
      <div class="hero-visual" aria-hidden="true">
        <img :src="workflowImage" alt="" />
      </div>
    </section>

    <section class="flow-panel">
      <div class="section-title">
        <div>
          <h2>処理の流れ</h2>
          <p>アップロード後、この順番でチェックと変換を行います。</p>
        </div>
      </div>
      <ol class="flow-steps">
        <li v-for="(step, index) in processSteps" :key="step.title" :class="{ active: busy && currentStep === index + 1, done: currentStep > index + 1 }">
          <span>{{ index + 1 }}</span>
          <div>
            <strong>{{ step.title }}</strong>
            <p>{{ step.text }}</p>
          </div>
        </li>
      </ol>
    </section>

    <section
      class="upload-area"
      :class="{ active: dragging }"
      @dragover.prevent="dragging = true"
      @dragleave.prevent="dragging = false"
      @drop.prevent="onDrop"
    >
      <div class="upload-copy">
        <span class="upload-icon">XLSX</span>
        <div>
          <h2>ファイルをここへドロップ</h2>
          <p>通常の .xlsx ファイルを複数まとめて処理できます。</p>
        </div>
      </div>
      <label class="file-button">
        ファイルを選択
        <input type="file" accept=".xlsx" multiple @change="addFiles($event.target.files)" />
      </label>
    </section>

    <section v-if="files.length" class="queue">
      <div class="section-title">
        <div>
          <h2>処理待ちファイル</h2>
          <p>{{ files.length }}件 / {{ totalSize }}</p>
        </div>
        <button type="button" class="ghost-button" @click="clearAll">一覧をクリア</button>
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

    <section v-if="busy || progress > 0" class="progress-panel">
      <div class="progress-head">
        <strong>{{ progressMessage }}</strong>
        <span>{{ progress }}%</span>
      </div>
      <div class="progress-track" aria-hidden="true">
        <div class="progress-fill" :style="{ width: `${progress}%` }"></div>
      </div>
    </section>

    <section v-if="results.length" class="results">
      <div class="section-title">
        <div>
          <h2>処理結果</h2>
          <p>生成されたファイルは下のリンクから開けます。</p>
        </div>
      </div>
      <article v-for="result in results" :key="result.original_name" class="result-row" :data-status="result.status">
        <div class="result-main">
          <span class="status-badge">{{ statusLabel(result.status) }}</span>
          <p class="file-name">{{ result.original_name }}</p>
          <p class="status-line">{{ statusText(result.status) }}</p>
          <p class="folder-path">{{ result.folder }}</p>
        </div>
        <div class="downloads">
          <p>出力ファイル</p>
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
      <div class="section-title">
        <div>
          <h2>出力先フォルダ</h2>
          <p>処理結果はローカルの各フォルダにも保存されます。</p>
        </div>
      </div>
      <dl>
        <div v-for="key in ['network_ready', 'needs_confirmation', 'error']" :key="key">
          <dt>{{ folderLabel(key) }}</dt>
          <dd>{{ folders[key] || `data/${key}` }}</dd>
        </div>
      </dl>
    </section>
  </main>
</template>
