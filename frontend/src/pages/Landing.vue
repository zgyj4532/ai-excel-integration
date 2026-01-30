<template>
  <div class="landing-shell">
    <div class="grain"></div>
    <div class="mesh mesh-left"></div>
    <div class="mesh mesh-right"></div>

    <header class="hero">
      <div class="hero-text">
        <p class="eyebrow">{{ $t('prototype') }}</p>
        <h1 class="headline">{{ $t('landingTitle') }}</h1>
        <p class="subline">{{ $t('landingSubtitle') }}</p>
        <div class="divider"></div>
        <div class="cta-row">
          <button class="btn primary" @click="goWorkspace">
            <span>{{ $t('landingGoWorkspace') }}</span>
            <span class="arrow">→</span>
          </button>
          <button class="btn ghost" @click="openDemo">{{ $t('landingOpenDemo') }}</button>
        </div>
      </div>
      <div class="hero-meta">
        <div class="meta-card">
          <p class="meta-label">AI · Excel</p>
          <p class="meta-value">{{ $t('landingCoreAI3') }}</p>
        </div>
        <div class="meta-card">
          <p class="meta-label">Data Flow</p>
          <p class="meta-value">{{ $t('landingCoreAdvanced2') }}</p>
        </div>
      </div>
    </header>

    <section class="capabilities">
      <div class="section-head">
        <h3>{{ $t('landingCoreTitle') }}</h3>
        <p class="section-note">{{ $t('landingCoreAI2') }}</p>
      </div>

      <div class="core-grid-1">
        <div
          v-for="(group, idx) in coreGroupsRow1"
          :key="idx"
          class="core-card"
          :style="{ '--delay': (idx * 0.08) + 's' }"
        >
          <div class="core-title">{{ group.title }}</div>
          <div class="core-illustration" v-html="svgMap[group.key]"></div>
          <ul>
            <li v-for="(item, i) in group.items" :key="i">{{ item }}</li>
          </ul>
        </div>
      </div>

      <div class="core-grid-2">
        <div
          v-for="(group, idx) in coreGroupsRow2"
          :key="idx"
          class="core-card"
          :style="{ '--delay': (idx * 0.08) + 's' }"
        >
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
    <rect x="6" y="18" width="208" height="84" rx="10" fill="rgba(197,160,89,0.06)" stroke="rgba(197,160,89,0.45)" stroke-width="2"/>
    <rect x="18" y="30" width="80" height="12" rx="6" fill="#c5a059" opacity="0.9"/>
    <rect x="18" y="50" width="56" height="8" rx="4" fill="#8fd0b4" opacity="0.9"/>
    <rect x="18" y="64" width="96" height="8" rx="4" fill="#c5a059" opacity="0.4"/>
    <rect x="120" y="38" width="72" height="32" rx="8" fill="rgba(46,204,113,0.08)" stroke="#2ecc71" stroke-width="2"/>
    <text x="156" y="58" text-anchor="middle" fill="#c5a059" font-size="13" font-family="IBM Plex Mono">AI</text>
    <circle cx="192" cy="90" r="10" fill="rgba(46,204,113,0.2)" stroke="#2ecc71" stroke-width="2"/>
    <path d="M188 90l6 6 8-12" stroke="#c5a059" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
  </svg>
  `,
  data: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(10,11,14,0.7)" stroke="rgba(197,160,89,0.35)" stroke-width="2"/>
    <rect x="28" y="74" width="24" height="22" rx="4" fill="#c5a059" opacity="0.9"/>
    <rect x="64" y="62" width="24" height="34" rx="4" fill="#8fd0b4" opacity="0.9"/>
    <rect x="100" y="52" width="24" height="44" rx="4" fill="#2ecc71" opacity="0.9"/>
    <rect x="136" y="40" width="24" height="56" rx="4" fill="#d97706" opacity="0.9"/>
    <rect x="172" y="30" width="24" height="66" rx="4" fill="#9a3412" opacity="0.9"/>
    <path d="M28 82c20-24 60-42 100-28 16 6 30 16 48 8" stroke="#c5a059" stroke-width="3" stroke-linecap="round" fill="none"/>
    <circle cx="62" cy="54" r="5" fill="#c5a059"/>
    <circle cx="122" cy="62" r="5" fill="#2ecc71"/>
    <circle cx="176" cy="48" r="5" fill="#8fd0b4"/>
  </svg>
  `,
  biz: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(10,11,14,0.72)" stroke="rgba(197,160,89,0.4)" stroke-width="2"/>
    <rect x="26" y="34" width="60" height="50" rx="8" fill="rgba(197,160,89,0.16)" stroke="#c5a059" stroke-width="2"/>
    <rect x="40" y="56" width="32" height="20" rx="4" fill="#c5a059"/>
    <rect x="92" y="30" width="60" height="60" rx="10" fill="rgba(46,204,113,0.12)" stroke="#2ecc71" stroke-width="2"/>
    <path d="M108 70h28" stroke="#2ecc71" stroke-width="3" stroke-linecap="round"/>
    <path d="M108 60h36" stroke="#2ecc71" stroke-width="3" stroke-linecap="round"/>
    <path d="M108 50h22" stroke="#2ecc71" stroke-width="3" stroke-linecap="round"/>
    <rect x="164" y="26" width="36" height="68" rx="10" fill="rgba(143,208,180,0.18)" stroke="#8fd0b4" stroke-width="2"/>
    <path d="M174 76c6-8 10-16 10-24" stroke="#8fd0b4" stroke-width="3" stroke-linecap="round"/>
    <path d="M174 84c10-2 18-8 22-16" stroke="#8fd0b4" stroke-width="3" stroke-linecap="round"/>
  </svg>
  `,
  advanced: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(10,11,14,0.72)" stroke="rgba(197,160,89,0.38)" stroke-width="2"/>
    <path d="M36 74c18-22 50-42 82-26 14 6 26 16 42 12 10-2 18-10 24-22" stroke="#c5a059" stroke-width="3" stroke-linecap="round" fill="none"/>
    <rect x="34" y="32" width="24" height="12" rx="6" fill="#c5a059" opacity="0.16"/>
    <rect x="78" y="40" width="24" height="12" rx="6" fill="#8fd0b4" opacity="0.22"/>
    <rect x="122" y="52" width="24" height="12" rx="6" fill="#2ecc71" opacity="0.35"/>
    <rect x="166" y="36" width="24" height="12" rx="6" fill="#d97706" opacity="0.4"/>
    <circle cx="60" cy="92" r="8" fill="#c5a059" opacity="0.25" stroke="#c5a059" stroke-width="2"/>
    <circle cx="100" cy="92" r="8" fill="#2ecc71" opacity="0.25" stroke="#2ecc71" stroke-width="2"/>
    <circle cx="140" cy="92" r="8" fill="#8fd0b4" opacity="0.25" stroke="#8fd0b4" stroke-width="2"/>
    <circle cx="180" cy="92" r="8" fill="#d97706" opacity="0.25" stroke="#d97706" stroke-width="2"/>
  </svg>
  `,
  realtime: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(10,11,14,0.72)" stroke="rgba(197,160,89,0.32)" stroke-width="2"/>
    <rect x="22" y="34" width="72" height="52" rx="8" fill="rgba(197,160,89,0.08)" stroke="rgba(197,160,89,0.4)" stroke-width="2"/>
    <rect x="36" y="50" width="44" height="8" rx="4" fill="#c5a059" opacity="0.9"/>
    <rect x="36" y="64" width="32" height="8" rx="4" fill="#8fd0b4" opacity="0.9"/>
    <rect x="108" y="26" width="86" height="64" rx="10" fill="rgba(46,204,113,0.12)" stroke="#2ecc71" stroke-width="2"/>
    <rect x="118" y="46" width="34" height="8" rx="4" fill="#c5a059" opacity="0.9"/>
    <rect x="118" y="62" width="46" height="8" rx="4" fill="#2ecc71" opacity="0.9"/>
    <circle cx="190" cy="46" r="10" fill="rgba(46,204,113,0.2)" stroke="#2ecc71" stroke-width="2"/>
    <path d="M186 46l6 6 8-12" stroke="#c5a059" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
    <path d="M22 30h54" stroke="#c5a059" stroke-width="3" stroke-linecap="round"/>
    <path d="M108 92h58" stroke="#2ecc71" stroke-width="3" stroke-linecap="round"/>
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
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@500;700&family=IBM+Plex+Mono:wght@400;600&display=swap');

:global(body) {
  background: #05070a;
}

.landing-shell {
  position: relative;
  padding: 48px 32px 64px;
  color: #e0e0e0;
  background: linear-gradient(145deg, rgba(10,11,14,0.95), rgba(10,12,16,0.92));
  overflow: hidden;
  min-height: 100vh;
  max-height: 100vh;
  overflow-y: auto;
}

.grain {
  position: absolute;
  inset: 0;
  background-image: radial-gradient(rgba(197,160,89,0.06) 1px, transparent 1px);
  background-size: 4px 4px;
  opacity: 0.25;
  pointer-events: none;
  mix-blend-mode: soft-light;
  z-index: 0;
}

.mesh {
  position: absolute;
  width: 420px;
  height: 420px;
  filter: blur(60px);
  opacity: 0.25;
  pointer-events: none;
  z-index: 0;
}

.mesh-left { top: -120px; left: -80px; background: radial-gradient(circle at 30% 30%, rgba(197,160,89,0.35), transparent 50%); }
.mesh-right { bottom: -120px; right: -120px; background: radial-gradient(circle at 70% 70%, rgba(46,204,113,0.25), transparent 55%); }

.hero {
  position: relative;
  z-index: 1;
  display: grid;
  grid-template-columns: 2fr 1fr;
  gap: 32px;
  align-items: center;
  margin-bottom: 40px;
}

.hero-text {
  max-width: 720px;
}

.eyebrow {
  font-family: 'IBM Plex Mono', monospace;
  letter-spacing: 0.2em;
  text-transform: uppercase;
  color: rgba(197,160,89,0.85);
  margin: 0 0 8px;
}

.headline {
  font-family: 'Orbitron', sans-serif;
  font-size: clamp(32px, 4vw, 56px);
  margin: 0;
  letter-spacing: 0.03em;
}

.subline {
  margin: 12px 0 16px;
  color: rgba(224,224,224,0.75);
  font-size: 16px;
}

.divider {
  width: 120px;
  height: 2px;
  background: linear-gradient(90deg, rgba(197,160,89,0.8), rgba(46,204,113,0.5));
  margin: 16px 0 20px;
}

.cta-row {
  display: flex;
  gap: 12px;
  flex-wrap: wrap;
}

.btn {
  border: 1px solid rgba(197,160,89,0.6);
  background: transparent;
  color: #e0e0e0;
  padding: 12px 18px;
  border-radius: 12px;
  font-family: 'IBM Plex Mono', monospace;
  letter-spacing: 0.05em;
  cursor: pointer;
  transition: all 0.25s ease;
  display: inline-flex;
  align-items: center;
  gap: 8px;
}

.btn.primary {
  background: linear-gradient(135deg, rgba(197,160,89,0.9), rgba(197,160,89,0.65));
  color: #05070a;
  box-shadow: 0 10px 40px rgba(197,160,89,0.35);
}

.btn.ghost { border-color: rgba(224,224,224,0.25); color: rgba(224,224,224,0.85); }

.btn:hover { transform: translateY(-2px); border-color: rgba(197,160,89,0.9); box-shadow: 0 10px 40px rgba(197,160,89,0.15); }
.btn.primary:hover { box-shadow: 0 14px 50px rgba(197,160,89,0.4); }

.arrow { font-size: 16px; transition: transform 0.25s ease; }
.btn:hover .arrow { transform: translateX(4px); }

.hero-meta { display: grid; gap: 12px; align-content: start; }
.meta-card {
  background: rgba(255,255,255,0.02);
  border: 1px solid rgba(197,160,89,0.2);
  border-radius: 12px;
  padding: 12px 14px;
  box-shadow: 0 12px 40px rgba(0,0,0,0.35);
}
.meta-label { margin: 0; color: rgba(224,224,224,0.6); font-family: 'IBM Plex Mono', monospace; font-size: 12px; letter-spacing: 0.1em; }
.meta-value { margin: 4px 0 0; font-weight: 700; color: #e0e0e0; }

.capabilities { position: relative; z-index: 1; }
.section-head { display: flex; flex-direction: column; gap: 6px; margin-bottom: 12px; opacity: 0; transform: translateY(8px); animation: riseFade 0.8s cubic-bezier(0.16, 1, 0.3, 1) 0.05s forwards; }
.section-head h3 { margin: 0; font-family: 'Orbitron', sans-serif; letter-spacing: 0.05em; }
.section-note { margin: 0; color: rgba(224,224,224,0.65); font-family: 'IBM Plex Mono', monospace; }

.core-grid-1, .core-grid-2 {
  display: flex;
  justify-content: center;
  align-items: stretch;
  gap: 18px;
  margin-top: 14px;
  flex-wrap: wrap;
}

.core-card {
  background: rgba(255,255,255,0.02);
  border: 1px solid rgba(197,160,89,0.18);
  border-radius: 14px;
  padding: 14px;
  width: min(360px, 32vw);
  box-shadow: 0 18px 45px rgba(0,0,0,0.35);
  backdrop-filter: blur(2px);
  transition: transform 0.25s ease, border-color 0.25s ease, box-shadow 0.25s ease;
  opacity: 0;
  transform: translateY(18px) scale(0.98);
  animation: riseFade 0.8s cubic-bezier(0.16, 1, 0.3, 1) forwards;
  animation-delay: var(--delay, 0s);
}

.core-card:hover {
  transform: translateY(-6px);
  border-color: rgba(197,160,89,0.35);
  box-shadow: 0 24px 55px rgba(0,0,0,0.45);
}

.core-title {
  font-weight: 700;
  margin-bottom: 10px;
  font-size: 16px;
  font-family: 'Orbitron', sans-serif;
  letter-spacing: 0.04em;
}

.core-card ul {
  padding-left: 16px;
  margin: 10px 0 0;
  display: flex;
  flex-direction: column;
  gap: 4px;
  line-height: 1.4;
  color: rgba(224,224,224,0.8);
}

.core-illustration {
  margin: 4px 0 10px;
  border: 1px dashed rgba(197,160,89,0.2);
  border-radius: 10px;
  padding: 8px;
  background: rgba(255,255,255,0.02);
  box-shadow: inset 0 0 0 1px rgba(255,255,255,0.02);
  animation: glowPulse 3.6s ease-in-out infinite alternate;
  animation-delay: calc(var(--delay, 0s) + 0.3s);
}

.core-illustration svg { display: block; width: 100%; height: auto; }

@keyframes riseFade {
  from {
    opacity: 0;
    transform: translateY(18px) scale(0.98);
    filter: blur(2px);
  }
  to {
    opacity: 1;
    transform: translateY(0) scale(1);
    filter: blur(0);
  }
}

@keyframes glowPulse {
  from { box-shadow: inset 0 0 0 1px rgba(255,255,255,0.02), 0 0 0 rgba(197,160,89,0); }
  to { box-shadow: inset 0 0 0 1px rgba(255,255,255,0.08), 0 10px 40px rgba(197,160,89,0.2); }
}

/* Custom scrollbar to match industrial palette */
.landing-shell::-webkit-scrollbar { width: 10px; }
.landing-shell::-webkit-scrollbar-track { background: rgba(255,255,255,0.02); border-left: 1px solid rgba(197,160,89,0.15); }
.landing-shell::-webkit-scrollbar-thumb {
  background: #f5f5f5;
  border-radius: 10px;
  border: 1px solid rgba(10,11,14,0.6);
}
.landing-shell::-webkit-scrollbar-thumb:hover {
  background: #ffffff;
}

@media (max-width: 960px) {
  .hero { grid-template-columns: 1fr; }
  .hero-meta { grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); }
  .core-card { width: 100%; }
}

@media (max-width: 640px) {
  .landing-shell { padding: 32px 20px 48px; }
  .headline { font-size: clamp(28px, 8vw, 40px); }
}
</style>

