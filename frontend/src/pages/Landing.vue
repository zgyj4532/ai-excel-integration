<template>
  <div style="padding:32px">
    <h1>{{ $t('landingTitle') }}</h1>
    <p class="muted">{{ $t('landingSubtitle') }}</p>

    <div style="margin-top:24px; display:flex; gap:12px">
      <button @click="goWorkspace">{{ $t('landingGoWorkspace') }}</button>
      <button @click="openDemo">{{ $t('landingOpenDemo') }}</button>
    </div>

    <section style="margin-top:20px">
      <h3>{{ $t('landingCoreTitle') }}</h3>
      <div class="core-grid-1">
        <div v-for="(group, idx) in coreGroupsRow1" :key="idx" class="core-card">
          <div class="core-title">{{ group.title }}</div>
          <div class="core-illustration" v-html="svgMap[group.key]"></div>
          <ul>
            <li v-for="(item, i) in group.items" :key="i">{{ item }}</li>
          </ul>
        </div>
      </div>
      <div class="core-grid-2">
        <div v-for="(group, idx) in coreGroupsRow2" :key="idx" class="core-card">
          <div class="core-title">{{ group.title }}</div>
          <div class="core-illustration" v-html="svgMap[group.key]"></div>
          <ul>
            <li v-for="(item, i) in group.items" :key="i">{{ item }}</li>
          </ul>
        </div>
      </div>
    </section>
  </div>
</template>

<script setup lang="ts">
import { useRouter } from 'vue-router'
const router = useRouter()
function goWorkspace(){ router.push({ name: 'Workspace' }) }
import { useI18n } from 'vue-i18n'
const { t } = useI18n()
function openDemo(){ alert(t('landingDemoAlert')) }
import { computed } from 'vue'

const svgMap: Record<string, string> = {
  ai: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="6" y="18" width="208" height="84" rx="10" fill="rgba(138,43,226,0.08)" stroke="rgba(138,43,226,0.35)" stroke-width="2"/>
    <rect x="18" y="30" width="80" height="12" rx="6" fill="#8a2be2" opacity="0.9"/>
    <rect x="18" y="50" width="56" height="8" rx="4" fill="#c084fc" opacity="0.9"/>
    <rect x="18" y="64" width="96" height="8" rx="4" fill="#8a2be2" opacity="0.5"/>
    <rect x="120" y="38" width="72" height="32" rx="8" fill="#0ea5e9" opacity="0.12" stroke="#0ea5e9" stroke-width="2"/>
    <text x="156" y="58" text-anchor="middle" fill="#0ea5e9" font-size="13" font-family="Arial">AI</text>
    <circle cx="192" cy="90" r="10" fill="#22c55e" opacity="0.18" stroke="#22c55e" stroke-width="2"/>
    <path d="M188 90l6 6 8-12" stroke="#22c55e" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
  </svg>
  `,
  data: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(14,165,233,0.08)" stroke="rgba(14,165,233,0.35)" stroke-width="2"/>
    <rect x="28" y="74" width="24" height="22" rx="4" fill="#8b5cf6" opacity="0.8"/>
    <rect x="64" y="62" width="24" height="34" rx="4" fill="#0ea5e9" opacity="0.8"/>
    <rect x="100" y="52" width="24" height="44" rx="4" fill="#22c55e" opacity="0.8"/>
    <rect x="136" y="40" width="24" height="56" rx="4" fill="#f59e0b" opacity="0.85"/>
    <rect x="172" y="30" width="24" height="66" rx="4" fill="#ef4444" opacity="0.8"/>
    <path d="M28 82c20-24 60-42 100-28 16 6 30 16 48 8" stroke="#0ea5e9" stroke-width="3" stroke-linecap="round" fill="none"/>
    <circle cx="62" cy="54" r="5" fill="#0ea5e9"/>
    <circle cx="122" cy="62" r="5" fill="#0ea5e9"/>
    <circle cx="176" cy="48" r="5" fill="#0ea5e9"/>
  </svg>
  `,
  biz: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(34,197,94,0.06)" stroke="rgba(34,197,94,0.35)" stroke-width="2"/>
    <rect x="26" y="34" width="60" height="50" rx="8" fill="#22c55e" opacity="0.18" stroke="#22c55e" stroke-width="2"/>
    <rect x="40" y="56" width="32" height="20" rx="4" fill="#22c55e"/>
    <rect x="92" y="30" width="60" height="60" rx="10" fill="#0ea5e9" opacity="0.14" stroke="#0ea5e9" stroke-width="2"/>
    <path d="M108 70h28" stroke="#0ea5e9" stroke-width="3" stroke-linecap="round"/>
    <path d="M108 60h36" stroke="#0ea5e9" stroke-width="3" stroke-linecap="round"/>
    <path d="M108 50h22" stroke="#0ea5e9" stroke-width="3" stroke-linecap="round"/>
    <rect x="164" y="26" width="36" height="68" rx="10" fill="#8b5cf6" opacity="0.18" stroke="#8b5cf6" stroke-width="2"/>
    <path d="M174 76c6-8 10-16 10-24" stroke="#8b5cf6" stroke-width="3" stroke-linecap="round"/>
    <path d="M174 84c10-2 18-8 22-16" stroke="#8b5cf6" stroke-width="3" stroke-linecap="round"/>
  </svg>
  `,
  advanced: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(239,68,68,0.06)" stroke="rgba(239,68,68,0.35)" stroke-width="2"/>
    <path d="M36 74c18-22 50-42 82-26 14 6 26 16 42 12 10-2 18-10 24-22" stroke="#ef4444" stroke-width="3" stroke-linecap="round" fill="none"/>
    <rect x="34" y="32" width="24" height="12" rx="6" fill="#ef4444" opacity="0.2"/>
    <rect x="78" y="40" width="24" height="12" rx="6" fill="#ef4444" opacity="0.3"/>
    <rect x="122" y="52" width="24" height="12" rx="6" fill="#ef4444" opacity="0.5"/>
    <rect x="166" y="36" width="24" height="12" rx="6" fill="#ef4444" opacity="0.7"/>
    <circle cx="60" cy="92" r="8" fill="#0ea5e9" opacity="0.2" stroke="#0ea5e9" stroke-width="2"/>
    <circle cx="100" cy="92" r="8" fill="#22c55e" opacity="0.2" stroke="#22c55e" stroke-width="2"/>
    <circle cx="140" cy="92" r="8" fill="#8b5cf6" opacity="0.2" stroke="#8b5cf6" stroke-width="2"/>
    <circle cx="180" cy="92" r="8" fill="#f59e0b" opacity="0.2" stroke="#f59e0b" stroke-width="2"/>
  </svg>
  `,
  realtime: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(59,130,246,0.06)" stroke="rgba(59,130,246,0.32)" stroke-width="2"/>
    <rect x="22" y="34" width="72" height="52" rx="8" fill="#3b82f6" opacity="0.12" stroke="#3b82f6" stroke-width="2"/>
    <rect x="36" y="50" width="44" height="8" rx="4" fill="#3b82f6" opacity="0.8"/>
    <rect x="36" y="64" width="32" height="8" rx="4" fill="#60a5fa" opacity="0.9"/>
    <rect x="108" y="26" width="86" height="64" rx="10" fill="#0ea5e9" opacity="0.1" stroke="#0ea5e9" stroke-width="2"/>
    <rect x="118" y="46" width="34" height="8" rx="4" fill="#0ea5e9" opacity="0.9"/>
    <rect x="118" y="62" width="46" height="8" rx="4" fill="#3b82f6" opacity="0.9"/>
    <circle cx="190" cy="46" r="10" fill="#22c55e" opacity="0.16" stroke="#22c55e" stroke-width="2"/>
    <path d="M186 46l6 6 8-12" stroke="#22c55e" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
    <path d="M22 30h54" stroke="#3b82f6" stroke-width="3" stroke-linecap="round"/>
    <path d="M108 92h58" stroke="#0ea5e9" stroke-width="3" stroke-linecap="round"/>
  </svg>
  `
}

const coreGroupsRow1 = computed(() => [
  { key: 'ai', title: t('landingCoreAI'), items: [t('landingCoreAI1'), t('landingCoreAI2'), t('landingCoreAI3'), t('landingCoreAI4')] },
  { key: 'biz', title: t('landingCoreBiz'), items: [t('landingCoreBiz1'), t('landingCoreBiz2'), t('landingCoreBiz3'), t('landingCoreBiz4')] },
  { key: 'data', title: t('landingCoreData'), items: [t('landingCoreData1'), t('landingCoreData2'), t('landingCoreData3'), t('landingCoreData4')] }
])

const coreGroupsRow2 = computed(() => [
  { key: 'advanced', title: t('landingCoreAdvanced'), items: [t('landingCoreAdvanced1'), t('landingCoreAdvanced2'), t('landingCoreAdvanced3'), t('landingCoreAdvanced4')] },
  { key: 'realtime', title: t('landingCoreRealtime'), items: [t('landingCoreRealtime1'), t('landingCoreRealtime2'), t('landingCoreRealtime3'), t('landingCoreRealtime4')] }
])
</script>

<style scoped>
.mermaid-diagram {
  background: rgba(255,255,255,0.02);
  border-radius: 6px;
  padding: 12px;
  margin-top: 12px;
}
.core-grid-1, .core-grid-2 { display:flex; justify-content:center; align-items:stretch; gap:16px; margin-top:12px; flex-wrap:wrap }
.core-card { background: rgba(255,255,255,0.04); border:1px solid rgba(255,255,255,0.08); border-radius:10px; padding:12px }
.core-title { font-weight:700; margin-bottom:8px; font-size:15px }
.core-card ul { padding-left:16px; margin:0; display:flex; flex-direction:column; gap:4px; line-height:1.4 }
.core-illustration { margin:4px 0 10px; border:1px dashed rgba(255,255,255,0.08); border-radius:8px; padding:6px; background: rgba(255,255,255,0.02) }
.core-illustration svg { display:block; width:100%; height:auto }
</style>

