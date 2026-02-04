<template>
  <div class="settings">
    <div class="topbar">
      <h2>{{ $t('settingsTitle') }}</h2>
    </div>

    <div class="content">
      <!-- 第一行：API 配置（全宽） -->
      <div class="row row-first">
        <div class="card">
          <h4>{{ $t('apiConfig') }}</h4>
          <div class="api-config">
            <label>{{ $t('apiEndpointLabel') }}</label>
            <input v-model="apiEndpoint" :placeholder="$t('apiEndpointPlaceholder')" />
          </div>
          <div class="api-config api-key-config">
            <label>{{ $t('apiKeyInputLabel') }}</label>
            <input v-model="apiKey" :placeholder="$t('apiKeyPlaceholder')" type="password" />
          </div>
          <div class="api-actions">
            <button @click="saveApiSettings" class="save-btn">{{ $t('saveBtn') }}</button>
          </div>
          <div class="muted" style="margin-top: 8px;">{{ $t('currentApi', { api: apiEndpoint }) }}</div>
          <div class="muted" v-if="apiKeyMasked">{{'api-key: ' + apiKeyMasked }}</div>
        </div>
      </div>

      <!-- 第二行：服务状态 与 系统配置 并排 -->
      <div class="row row-second">
        <div class="card service-status-card">
          <h4>{{ $t('serviceStatus') }}</h4>
          <div class="status-info">
                  <div class="status-item">
                    <span class="status-label">{{ $t('backendStatusLabel') }}:</span>
                    <span :class="['status-value', backendStatusClass]">{{ backendStatus }}</span>
                  </div>
                  <div class="status-item">
                    <span class="status-label">{{ $t('apiKeyLabel') }}:</span>
                    <span :class="['status-value', apiKeyStatusClass]">{{ apiKeyStatus }}</span>
                  </div>
                  <div class="status-item">
                    <span class="status-label">{{ $t('wsConfigLabel') }}:</span>
                    <span :class="['status-value', wsConfigStatusClass]">{{ wsConfigStatus }}</span>
                  </div>
          </div>
          <button @click="checkStatus" class="check-status-btn">{{ $t('checkStatusBtn') }}</button>
        </div>

        <div class="card system-config-card">
          <h4>{{ $t('systemConfigTitle') }}</h4>
          <div class="system-config">
            <div class="config-item">
              <label>
                <input type="checkbox" v-model="autoSave" /> {{ $t('autoSaveLabel') }}
              </label>
            </div>
            <div class="config-item">
              <label>
                <input type="checkbox" v-model="notifications" /> {{ $t('notificationsLabel') }}
              </label>
            </div>
            <div class="config-item">
              <label>{{ $t('themeLabel') }}: 
                <select v-model="theme">
                  <option value="light">{{ $t('themeLight') }}</option>
                  <option value="dark">{{ $t('themeDark') }}</option>
                </select>
              </label>
            </div>
          </div>
          <button @click="saveSystemConfig" class="save-btn">{{ $t('saveConfigBtn') }}</button>
        </div>
      </div>

      <!-- 第三行：关于（全宽） -->
      <div class="row row-third">
        <div class="card about-card">
          <h4>{{ $t('aboutTitle') }}</h4>
          <div class="about-info">
            <p>{{ $t('aboutLine1') }}</p>
            <p>{{ $t('aboutLine2') }}</p>
            <p>{{ $t('aboutCopyright') }}</p>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, computed } from 'vue'
import { useI18n } from 'vue-i18n'
import { checkServiceStatus, getSystemSettings, saveSystemSettings } from '@/services/api'
import { fetchJson, setApiBaseUrl } from '@/services/apiClient'

const { t } = useI18n()

// API配置：默认使用页面 origin
const apiEndpoint = ref((typeof window !== 'undefined' && window.location && window.location.origin) ? window.location.origin : 'http://localhost:8081')
const apiKey = ref('')
const apiKeyMasked = ref('')

// 服务状态
const backendStatus = ref(t('unknown'))
const apiKeyStatus = ref(t('unknown'))
const wsConfigStatus = ref(t('unknown'))

// 系统配置
const autoSave = ref(true)
const notifications = ref(true)
const theme = ref('dark')

// 计算属性
const backendStatusClass = computed(() => {
  if (backendStatus.value === t('checking')) return 'checking'
  if (backendStatus.value === t('unableConnect')) return 'error'
  return 'success'
})

const apiKeyStatusClass = computed(() => {
  if (apiKeyStatus.value === t('checking')) return 'checking'
  if (apiKeyStatus.value === t('unableConnect')) return 'error'
  return 'success'
})

const wsConfigStatusClass = computed(() => {
  if (wsConfigStatus.value === t('checking')) return 'checking'
  if (wsConfigStatus.value === t('unableConnect')) return 'error'
  return 'success'
})

// 方法
async function saveApiSettings() {
  const base = (apiEndpoint.value || '').replace(/\/$/, '')
  const messages: string[] = []

  // 保存 endpoint
  try {
    const settings = { apiEndpoint: apiEndpoint.value }
    const result = await saveSystemSettings(settings, apiEndpoint.value)
    if (result && result.success) {
      try { setApiBaseUrl(base) } catch (e) { /* ignore */ }
      messages.push(t('savedApi', { api: apiEndpoint.value }))
    } else {
      messages.push(t('saveApiFailed'))
    }
  } catch (e) {
    console.error('保存 API 配置失败:', e)
    messages.push(t('saveApiFailed'))
  }

  // 保存 apiKey（可选）
  if (apiKey.value) {
    try {
      const result = await fetchJson(`${base}/api/sendapi`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ apiKey: apiKey.value, baseUrl: base })
      })
      if (result && result.success) {
        apiKeyMasked.value = result.apiKeyMasked || ''
        try { localStorage.setItem('aiexcel_api_key', apiKey.value) } catch (e) { /* ignore */ }
        try { setApiBaseUrl(base) } catch (e) { /* ignore */ }
        messages.push(t('apiKeySaved'))
        await checkStatus()
      } else {
        messages.push(t('apiKeySaveFailed'))
      }
    } catch (error) {
      console.error('保存 API Key 失败:', error)
      messages.push(t('apiKeySaveFailed'))
    }
  }

  if (messages.length) alert(messages.join('\n'))
}

async function checkStatus() {
  backendStatus.value = t('checking')
  apiKeyStatus.value = t('checking')
  wsConfigStatus.value = t('checking')

  try {
    // Prefer using the API endpoint configured in this page (unsaved value allowed)
    const base = (apiEndpoint.value || '').replace(/\/$/, '')
    const url = `${base}/api/status`
    const result = await fetchJson(url)

    if (result && typeof result === 'object') {
      backendStatus.value = result.status === 'available' ? t('status_active') : t('status_inactive')
      apiKeyStatus.value = result.hasApiKey ? t('api_active') : t('api_inactive')
      wsConfigStatus.value = result.apiConfigured ? t('enabled') : t('disabled')
    } else {
      backendStatus.value = t('unknown')
      apiKeyStatus.value = t('unknown')
      wsConfigStatus.value = t('unknown')
    }
  } catch (error) {
    console.error('检查服务状态失败:', error)
    backendStatus.value = t('unableConnect')
    apiKeyStatus.value = t('unableConnect')
    wsConfigStatus.value = t('unableConnect')
  }
}

async function saveSystemConfig() {
  const settings = {
    apiEndpoint: apiEndpoint.value,
    autoSave: autoSave.value,
    notifications: notifications.value,
    theme: theme.value
  }

  try {
    const result = await saveSystemSettings(settings)
    if (result.success) {
      alert(t('saveConfigSuccess'))
    }
  } catch (error) {
    console.error('保存配置失败:', error)
    alert(t('saveConfigFailed'))
  }
}

// 初始化
onMounted(async () => {
  try {
    const result = await getSystemSettings(apiEndpoint.value)
    if (result.success) {
      const data = result.data
      apiEndpoint.value = data.apiEndpoint || apiEndpoint.value
      autoSave.value = data.autoSave ?? true
      notifications.value = data.notifications ?? true
      theme.value = data.theme || 'dark'
    }
  } catch (error) {
    console.error('获取系统设置失败:', error)
  }

  try {
    const base = (apiEndpoint.value || '').replace(/\/$/, '')
    const status = await fetchJson(`${base}/api/status?reveal=true`)
    if (status?.diagnosis?.keyPlain) {
      apiKey.value = status.diagnosis.keyPlain
    } else if (status?.apiKeyMasked) {
      apiKeyMasked.value = status.apiKeyMasked
    }
  } catch (e) {
    // ignore
  }

  if (!apiKey.value) {
    try {
      const cached = localStorage.getItem('aiexcel_api_key')
      if (cached) apiKey.value = cached
    } catch (e) {
      // ignore
    }
  }
})
</script>

<style scoped>
.settings {
  --bg-deep: #05070a;
  --accent-gold: #c5a059;
  --accent-teal: #19b394;
  --text-primary: #e0e0e0;
  --text-dim: rgba(224, 224, 224, 0.72);
  --panel: rgba(255, 255, 255, 0.03);
  --panel-strong: rgba(255, 255, 255, 0.08);
  --grid-line: rgba(197, 160, 89, 0.12);
  display: flex;
  flex-direction: column;
  height: 100%;
  background: radial-gradient(circle at 18% 22%, rgba(25, 179, 148, 0.1), transparent 34%),
              radial-gradient(circle at 78% 72%, rgba(197, 160, 89, 0.1), transparent 38%),
              linear-gradient(135deg, #06080d 0%, #0a0c11 50%, #05070a 100%);
  color: var(--text-primary);
  font-family: 'Space Grotesk', 'IBM Plex Mono', system-ui, sans-serif;
  position: relative;
  overflow: auto;
}

/* Match landing page scrollbar style */
.settings::-webkit-scrollbar { width: 10px; }
.settings::-webkit-scrollbar-track { background: rgba(255,255,255,0.02); border-left: 1px solid rgba(197,160,89,0.15); }
.settings::-webkit-scrollbar-thumb {
  background: #f5f5f5;
  border-radius: 10px;
  border: 1px solid rgba(10,11,14,0.6);
}
.settings::-webkit-scrollbar-thumb:hover { background: #ffffff; }

.settings::before {
  content: '';
  position: absolute;
  inset: 0;
  background-image: radial-gradient(var(--grid-line) 1px, transparent 0);
  background-size: 120px 120px;
  opacity: 0.4;
  pointer-events: none;
}

.settings > .content {
  display: flex;
  flex-direction: column;
  gap: 16px;
  position: relative;
  z-index: 1;
}

.topbar h2 {
  font-family: 'Orbitron', 'Space Grotesk', system-ui, sans-serif;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  margin: 0;
  color: var(--text-primary);
  animation: slideUpReveal 0.9s ease forwards;
}

.topbar .muted { color: var(--text-dim); }

.row-second {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 16px;
}

.service-status-card { min-height: 220px; }
.system-config-card { min-height: 220px; }
.about-card { min-height: 120px; }

@media (max-width: 900px) {
  .row-second { grid-template-columns: 1fr; }
}

.api-config,
.status-info,
.system-config {
  margin-top: 16px;
}

.card {
  background: var(--panel);
  border: 1px solid var(--grid-line);
  border-radius: 14px;
  padding: 16px;
  box-shadow: 0 18px 50px rgba(0, 0, 0, 0.45);
  position: relative;
  overflow: hidden;
  animation: slideUpReveal 0.9s cubic-bezier(0.16, 1, 0.3, 1) forwards;
}

.card::after {
  content: '';
  position: absolute;
  inset: 0;
  background: radial-gradient(circle at 18% 18%, rgba(25, 179, 148, 0.05), transparent 55%);
  pointer-events: none;
}

.card h4 {
  font-family: 'Orbitron', 'Space Grotesk', system-ui, sans-serif;
  letter-spacing: 0.06em;
  text-transform: uppercase;
  margin: 0 0 6px 0;
  color: var(--accent-gold);
}

.muted { color: var(--text-dim); }

.api-config input,
select {
  width: 100%;
  padding: 10px 12px;
  border-radius: 10px;
  border: 1px solid var(--panel-strong);
  background: rgba(8, 11, 18, 0.85);
  color: var(--text-primary);
  margin-top: 10px;
  margin-bottom: 10px;
  font-family: 'Space Grotesk', 'IBM Plex Mono', system-ui, sans-serif;
}

.status-item {
  display: flex;
  justify-content: space-between;
  margin-bottom: 8px;
  padding-bottom: 8px;
  border-bottom: 1px solid rgba(255, 255, 255, 0.06);
}

.status-label { color: var(--text-dim); }

.status-value {
  font-weight: 700;
  letter-spacing: 0.02em;
}

.status-value.success { color: var(--accent-teal); }
.status-value.error { color: #ef4444; }
.status-value.checking { color: #f59e0b; }

.check-status-btn,
.save-btn {
  background: var(--accent-gold);
  color: #0a0b0e;
  border: 0;
  padding: 10px 16px;
  border-radius: 10px;
  cursor: pointer;
  margin-top: 12px;
  font-weight: 700;
  letter-spacing: 0.02em;
  box-shadow: 0 10px 28px rgba(197, 160, 89, 0.26);
  transition: transform 0.2s ease, box-shadow 0.2s ease;
  font-family: 'Space Grotesk', 'IBM Plex Mono', system-ui, sans-serif;
}

.check-status-btn:hover,
.save-btn:hover {
  transform: translateY(-2px);
  box-shadow: 0 14px 34px rgba(0, 0, 0, 0.35);
}

.config-item { margin-bottom: 12px; }

.about-info p {
  margin: 6px 0;
  color: var(--text-dim);
}

@keyframes slideUpReveal {
  0% { transform: translateY(28px); opacity: 0; clip-path: inset(100% 0 0 0); }
  100% { transform: translateY(0); opacity: 1; clip-path: inset(0 0 0 0); }
}
</style>