<template>
  <div class="landing-shell">
    <div class="grain"></div>
    <div class="mesh mesh-left"></div>
    <div class="mesh mesh-right"></div>

    <!-- Anchor Navigation (VB-01) -->
    <nav class="anchor-nav" v-if="showNav">
      <a href="#overview" class="nav-link" @click.prevent="scrollTo('overview')">{{ $t('landingNavOverview') }}</a>
      <a href="#features" class="nav-link" @click.prevent="scrollTo('features')">{{ $t('landingNavFeatures') }}</a>
      <a href="#proof" class="nav-link" @click.prevent="scrollTo('proof')">{{ $t('landingNavProof') }}</a>
    </nav>

    <!-- Hero Section (VB-01: reduced noise, more whitespace) -->
    <header id="overview" class="hero">
      <div class="hero-text">
        <p class="eyebrow">{{ $t('prototype') }}</p>
        <h1 class="headline">{{ $t('landingTitle') }}</h1>
        <h1 class="headline headline-sub">{{ $t('landingTitle1') }}</h1>
        <p class="subline">{{ $t('landingSubtitle') }}</p>
        <div class="divider"></div>

        <!-- Stats Row (CT-01: quantified value) -->
        <div class="stats-row">
          <div class="stat-item">
            <span class="stat-value">{{ $t('landingStat1') }}</span>
            <span class="stat-desc">{{ $t('landingStat1Desc') }}</span>
          </div>
          <div class="stat-item">
            <span class="stat-value">{{ $t('landingStat2') }}</span>
            <span class="stat-desc">{{ $t('landingStat2Desc') }}</span>
          </div>
          <div class="stat-item">
            <span class="stat-value">{{ $t('landingStat3') }}</span>
            <span class="stat-desc">{{ $t('landingStat3Desc') }}</span>
          </div>
        </div>

        <!-- CTA Buttons (IX-01: enhanced with hover effects) -->
        <div class="cta-row">
          <button class="btn primary" @click="goWorkspace">
            <span>{{ $t('landingGoWorkspace') }}</span>
            <span class="arrow">→</span>
          </button>
          <button class="btn ghost" @click="openDemo">
            <svg class="play-icon" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <polygon points="5,3 19,12 5,21" fill="currentColor"/>
            </svg>
            <span>{{ $t('landingOpenDemo') }}</span>
          </button>
        </div>
      </div>

      <!-- Meta cards: hover-reveal labels (VB-01) -->
      <div class="hero-meta">
        <div class="meta-card">
          <div class="meta-label-wrapper">
            <p class="meta-label">AI · Excel</p>
          </div>
          <p class="meta-value">{{ $t('landingCoreAI3') }}</p>
        </div>
        <div class="meta-card">
          <div class="meta-label-wrapper">
            <p class="meta-label">Data Flow</p>
          </div>
          <p class="meta-value">{{ $t('landingCoreAdvanced2') }}</p>
        </div>
      </div>
    </header>

    <!-- Bento Box Grid (TC-02) -->
    <section id="features" class="capabilities">
      <div class="section-head">
        <h3>{{ $t('landingCoreTitle') }}</h3>
        <p class="section-note">{{ $t('landingCoreAI2') }}</p>
      </div>

      <!-- Scenario cards (CT-02) -->
      <div class="scenario-strip">
        <div class="scenario-card" v-for="(s, i) in scenarios" :key="i">
          <span class="scenario-num">0{{ i + 1 }}</span>
          <p class="scenario-text">{{ s }}</p>
        </div>
      </div>

      <!-- Bento grid -->
      <div class="bento-grid">
        <div
          v-for="(group, idx) in allCoreGroups"
          :key="idx"
          class="bento-card"
          :class="group.span"
          :style="{ '--delay': (idx * 0.08) + 's' }"
        >
          <div class="bento-header">
            <div class="bento-title">{{ group.title }}</div>
            <div class="bento-illustration" v-html="svgMap[group.key]"></div>
          </div>
          <ul>
            <li v-for="(item, i) in group.items" :key="i">{{ item }}</li>
          </ul>
        </div>
      </div>
    </section>

    <!-- Social Proof (CT-03) -->
    <section id="proof" class="social-proof">
      <h3 class="proof-title">{{ $t('landingSocialProof') }}</h3>
      <p class="proof-desc">{{ $t('landingSocialProofDesc') }}</p>
      <div class="proof-logos">
        <div class="proof-logo" v-for="n in 5" :key="n">
          <div class="logo-placeholder"></div>
        </div>
      </div>
    </section>
  </div>
</template>

<script setup lang="ts">
import { useRouter } from 'vue-router'
import { useI18n } from 'vue-i18n'
import { computed, ref, onMounted, onUnmounted } from 'vue'

const router = useRouter()
const { t } = useI18n()
const showNav = ref(false)

function goWorkspace() { router.push({ name: 'Workspace' }) }
function openDemo() { alert(t('landingDemoAlert')) }

function scrollTo(id: string) {
  document.getElementById(id)?.scrollIntoView({ behavior: 'smooth' })
}

function onScroll() {
  showNav.value = window.scrollY > 120
}

onMounted(() => window.addEventListener('scroll', onScroll, { passive: true }))
onUnmounted(() => window.removeEventListener('scroll', onScroll))

const scenarios = computed(() => [t('landingFeatureScenario1'), t('landingFeatureScenario2'), t('landingFeatureScenario3')])

const svgMap: Record<string, string> = {
  ai: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="6" y="18" width="208" height="84" rx="10" fill="rgba(16,185,129,0.06)" stroke="rgba(16,185,129,0.45)" stroke-width="2"/>
    <rect x="18" y="30" width="80" height="12" rx="6" fill="#10b981" opacity="0.9"/>
    <rect x="18" y="50" width="56" height="8" rx="4" fill="#6ee7b7" opacity="0.9"/>
    <rect x="18" y="64" width="96" height="8" rx="4" fill="#10b981" opacity="0.4"/>
    <rect x="120" y="38" width="72" height="32" rx="8" fill="rgba(16,185,129,0.08)" stroke="#10b981" stroke-width="2"/>
    <text x="156" y="58" text-anchor="middle" fill="#10b981" font-size="13" font-family="IBM Plex Mono">AI</text>
    <circle cx="192" cy="90" r="10" fill="rgba(16,185,129,0.2)" stroke="#10b981" stroke-width="2"/>
    <path d="M188 90l6 6 8-12" stroke="#10b981" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
  </svg>
  `,
  data: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(10,11,14,0.7)" stroke="rgba(16,185,129,0.35)" stroke-width="2"/>
    <rect x="28" y="74" width="24" height="22" rx="4" fill="#10b981" opacity="0.9"/>
    <rect x="64" y="62" width="24" height="34" rx="4" fill="#6ee7b7" opacity="0.9"/>
    <rect x="100" y="52" width="24" height="44" rx="4" fill="#34d399" opacity="0.9"/>
    <rect x="136" y="40" width="24" height="56" rx="4" fill="#059669" opacity="0.9"/>
    <rect x="172" y="30" width="24" height="66" rx="4" fill="#047857" opacity="0.9"/>
    <path d="M28 82c20-24 60-42 100-28 16 6 30 16 48 8" stroke="#10b981" stroke-width="3" stroke-linecap="round" fill="none"/>
    <circle cx="62" cy="54" r="5" fill="#10b981"/>
    <circle cx="122" cy="62" r="5" fill="#34d399"/>
    <circle cx="176" cy="48" r="5" fill="#6ee7b7"/>
  </svg>
  `,
  biz: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(10,11,14,0.72)" stroke="rgba(16,185,129,0.4)" stroke-width="2"/>
    <rect x="26" y="34" width="60" height="50" rx="8" fill="rgba(16,185,129,0.16)" stroke="#10b981" stroke-width="2"/>
    <rect x="40" y="56" width="32" height="20" rx="4" fill="#10b981"/>
    <rect x="92" y="30" width="60" height="60" rx="10" fill="rgba(52,211,153,0.12)" stroke="#34d399" stroke-width="2"/>
    <path d="M108 70h28" stroke="#34d399" stroke-width="3" stroke-linecap="round"/>
    <path d="M108 60h36" stroke="#34d399" stroke-width="3" stroke-linecap="round"/>
    <path d="M108 50h22" stroke="#34d399" stroke-width="3" stroke-linecap="round"/>
    <rect x="164" y="26" width="36" height="68" rx="10" fill="rgba(110,231,183,0.18)" stroke="#6ee7b7" stroke-width="2"/>
    <path d="M174 76c6-8 10-16 10-24" stroke="#6ee7b7" stroke-width="3" stroke-linecap="round"/>
    <path d="M174 84c10-2 18-8 22-16" stroke="#6ee7b7" stroke-width="3" stroke-linecap="round"/>
  </svg>
  `,
  advanced: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(10,11,14,0.72)" stroke="rgba(16,185,129,0.38)" stroke-width="2"/>
    <path d="M36 74c18-22 50-42 82-26 14 6 26 16 42 12 10-2 18-10 24-22" stroke="#10b981" stroke-width="3" stroke-linecap="round" fill="none"/>
    <rect x="34" y="32" width="24" height="12" rx="6" fill="#10b981" opacity="0.16"/>
    <rect x="78" y="40" width="24" height="12" rx="6" fill="#6ee7b7" opacity="0.22"/>
    <rect x="122" y="52" width="24" height="12" rx="6" fill="#34d399" opacity="0.35"/>
    <rect x="166" y="36" width="24" height="12" rx="6" fill="#059669" opacity="0.4"/>
    <circle cx="60" cy="92" r="8" fill="#10b981" opacity="0.25" stroke="#10b981" stroke-width="2"/>
    <circle cx="100" cy="92" r="8" fill="#34d399" opacity="0.25" stroke="#34d399" stroke-width="2"/>
    <circle cx="140" cy="92" r="8" fill="#6ee7b7" opacity="0.25" stroke="#6ee7b7" stroke-width="2"/>
    <circle cx="180" cy="92" r="8" fill="#059669" opacity="0.25" stroke="#059669" stroke-width="2"/>
  </svg>
  `,
  realtime: `
  <svg width="100%" viewBox="0 0 220 120" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="10" y="18" width="200" height="84" rx="10" fill="rgba(10,11,14,0.72)" stroke="rgba(16,185,129,0.32)" stroke-width="2"/>
    <rect x="22" y="34" width="72" height="52" rx="8" fill="rgba(16,185,129,0.08)" stroke="rgba(16,185,129,0.4)" stroke-width="2"/>
    <rect x="36" y="50" width="44" height="8" rx="4" fill="#10b981" opacity="0.9"/>
    <rect x="36" y="64" width="32" height="8" rx="4" fill="#6ee7b7" opacity="0.9"/>
    <rect x="108" y="26" width="86" height="64" rx="10" fill="rgba(52,211,153,0.12)" stroke="#34d399" stroke-width="2"/>
    <rect x="118" y="46" width="34" height="8" rx="4" fill="#10b981" opacity="0.9"/>
    <rect x="118" y="62" width="46" height="8" rx="4" fill="#34d399" opacity="0.9"/>
    <circle cx="190" cy="46" r="10" fill="rgba(52,211,153,0.2)" stroke="#34d399" stroke-width="2"/>
    <path d="M186 46l6 6 8-12" stroke="#10b981" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
    <path d="M22 30h54" stroke="#10b981" stroke-width="3" stroke-linecap="round"/>
    <path d="M108 92h58" stroke="#34d399" stroke-width="3" stroke-linecap="round"/>
  </svg>
  `
}

const allCoreGroups = computed(() => [
  { key: 'ai', title: t('landingCoreAI'), items: [t('landingCoreAI1'), t('landingCoreAI2'), t('landingCoreAI3'), t('landingCoreAI4')], span: 'span-wide' },
  { key: 'data', title: t('landingCoreData'), items: [t('landingCoreData1'), t('landingCoreData2'), t('landingCoreData3'), t('landingCoreData4')], span: '' },
  { key: 'biz', title: t('landingCoreBiz'), items: [t('landingCoreBiz1'), t('landingCoreBiz2'), t('landingCoreBiz3'), t('landingCoreBiz4')], span: '' },
  { key: 'advanced', title: t('landingCoreAdvanced'), items: [t('landingCoreAdvanced1'), t('landingCoreAdvanced2'), t('landingCoreAdvanced3'), t('landingCoreAdvanced4')], span: 'span-wide' },
  { key: 'realtime', title: t('landingCoreRealtime'), items: [t('landingCoreRealtime1'), t('landingCoreRealtime2'), t('landingCoreRealtime3'), t('landingCoreRealtime4')], span: 'span-full' }
])
</script>

<style scoped>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;500;700;900&family=IBM+Plex+Mono:wght@300;400;600&display=swap');

:global(body) {
  background: var(--bg-deep);
}

/* ── Anchor Nav (VB-01) ── */
.anchor-nav {
  position: fixed;
  top: 16px;
  left: 50%;
  transform: translateX(-50%);
  z-index: 100;
  display: flex;
  gap: 8px;
  background: var(--landing-nav-bg);
  backdrop-filter: blur(12px);
  border: 1px solid rgba(16, 185, 129, 0.2);
  border-radius: 24px;
  padding: 6px 16px;
  animation: navFadeIn 0.4s ease forwards;
}

.nav-link {
  color: var(--text-dim);
  text-decoration: none;
  font-family: 'IBM Plex Mono', monospace;
  font-size: 13px;
  padding: 4px 12px;
  border-radius: 16px;
  transition: color 0.2s, background 0.2s;
}

.nav-link:hover {
  color: #10b981;
  background: rgba(16, 185, 129, 0.08);
}

@keyframes navFadeIn {
  from { opacity: 0; transform: translateX(-50%) translateY(-8px); }
  to { opacity: 1; transform: translateX(-50%) translateY(0); }
}

/* ── Shell ── */
.landing-shell {
  position: relative;
  padding: 64px 48px 80px;
  color: var(--text-primary);
  background: var(--landing-bg);
  overflow: hidden;
  min-height: 100vh;
  max-height: 100vh;
  overflow-y: auto;
}

.grain {
  position: absolute;
  inset: 0;
  background-image: radial-gradient(rgba(16, 185, 129, 0.04) 1px, transparent 1px);
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

.mesh-left { top: -120px; left: -80px; background: radial-gradient(circle at 30% 30%, rgba(16, 185, 129, 0.35), transparent 50%); }
.mesh-right { bottom: -120px; right: -120px; background: radial-gradient(circle at 70% 70%, rgba(52, 211, 153, 0.25), transparent 55%); }

/* ── Hero ── */
.hero {
  position: relative;
  z-index: 1;
  display: grid;
  grid-template-columns: 2fr 1fr;
  gap: 48px;
  align-items: center;
  margin-bottom: 64px;
}

.hero-text {
  max-width: 720px;
}

.eyebrow {
  font-family: 'IBM Plex Mono', monospace;
  font-weight: 300;
  letter-spacing: 0.25em;
  text-transform: uppercase;
  color: rgba(16, 185, 129, 0.85);
  margin: 0 0 12px;
  font-size: 13px;
}

.headline {
  font-family: 'Orbitron', sans-serif;
  font-weight: 900;
  font-size: clamp(36px, 4.5vw, 60px);
  margin: 0;
  letter-spacing: 0.02em;
  line-height: 1.15;
}

.headline-sub {
  font-weight: 500;
  font-size: clamp(24px, 3vw, 40px);
  color: var(--text-primary);
  margin-top: 4px;
}

.subline {
  margin: 16px 0 20px;
  color: var(--text-dim);
  font-size: 16px;
  font-weight: 300;
  letter-spacing: 0.02em;
}

.divider {
  width: 120px;
  height: 2px;
  background: linear-gradient(90deg, rgba(16, 185, 129, 0.8), rgba(52, 211, 153, 0.5));
  margin: 16px 0 24px;
}

/* ── Stats (CT-01) ── */
.stats-row {
  display: flex;
  gap: 32px;
  margin-bottom: 28px;
}

.stat-item {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.stat-value {
  font-family: 'Orbitron', sans-serif;
  font-weight: 900;
  font-size: 32px;
  color: #10b981;
  letter-spacing: 0.02em;
}

.stat-desc {
  font-family: 'IBM Plex Mono', monospace;
  font-size: 12px;
  font-weight: 300;
  color: var(--text-dim);
  letter-spacing: 0.05em;
}

/* ── CTA (IX-01) ── */
.cta-row {
  display: flex;
  gap: 14px;
  flex-wrap: wrap;
}

.btn {
  border: 1px solid rgba(16, 185, 129, 0.5);
  background: transparent;
  color: var(--text-primary);
  padding: 14px 24px;
  border-radius: 12px;
  font-family: 'IBM Plex Mono', monospace;
  font-weight: 600;
  letter-spacing: 0.05em;
  cursor: pointer;
  transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
  display: inline-flex;
  align-items: center;
  gap: 10px;
}

.btn.primary {
  background: linear-gradient(135deg, rgba(16, 185, 129, 0.9), rgba(5, 150, 105, 0.85));
  color: var(--text-on-accent);
  box-shadow: 0 10px 40px rgba(16, 185, 129, 0.3);
  border-color: transparent;
}

.btn.ghost {
  border-color: var(--border-subtle);
  color: var(--text-primary);
}

.btn:hover {
  transform: translateY(-3px);
  border-color: rgba(16, 185, 129, 0.9);
  box-shadow: 0 14px 50px rgba(16, 185, 129, 0.2);
}

.btn.primary:hover {
  box-shadow: 0 16px 60px rgba(16, 185, 129, 0.45);
  background: linear-gradient(135deg, rgba(16, 185, 129, 1), rgba(5, 150, 105, 0.95));
}

.arrow {
  font-size: 16px;
  transition: transform 0.3s ease;
}

.btn:hover .arrow { transform: translateX(5px); }

.play-icon {
  width: 16px;
  height: 16px;
}

.btn.ghost:hover .play-icon {
  filter: drop-shadow(0 0 4px rgba(16, 185, 129, 0.5));
}

/* ── Meta cards (VB-01: hover reveal) ── */
.hero-meta {
  display: grid;
  gap: 16px;
  align-content: start;
}

.meta-card {
  background: rgba(255, 255, 255, 0.02);
  border: 1px solid rgba(16, 185, 129, 0.15);
  border-radius: 14px;
  padding: 16px;
  box-shadow: 0 12px 40px rgba(0, 0, 0, 0.3);
  transition: border-color 0.3s, box-shadow 0.3s, transform 0.3s;
}

.meta-card:hover {
  border-color: rgba(16, 185, 129, 0.35);
  box-shadow: 0 16px 50px rgba(16, 185, 129, 0.1);
  transform: translateY(-2px);
}

.meta-label-wrapper {
  overflow: hidden;
  max-height: 0;
  opacity: 0;
  transition: max-height 0.4s ease, opacity 0.3s ease, margin 0.3s ease;
  margin-bottom: 0;
}

.meta-card:hover .meta-label-wrapper {
  max-height: 40px;
  opacity: 1;
  margin-bottom: 6px;
}

.meta-label {
  margin: 0;
  color: rgba(16, 185, 129, 0.7);
  font-family: 'IBM Plex Mono', monospace;
  font-size: 12px;
  letter-spacing: 0.15em;
}

.meta-value {
  margin: 0;
  font-weight: 700;
  color: var(--text-primary);
}

/* ── Section ── */
.capabilities {
  position: relative;
  z-index: 1;
}

.section-head {
  display: flex;
  flex-direction: column;
  gap: 6px;
  margin-bottom: 20px;
  opacity: 0;
  transform: translateY(8px);
  animation: riseFade 0.8s cubic-bezier(0.16, 1, 0.3, 1) 0.05s forwards;
}

.section-head h3 {
  margin: 0;
  font-family: 'Orbitron', sans-serif;
  font-weight: 700;
  letter-spacing: 0.05em;
}

.section-note {
  margin: 0;
  color: var(--text-dim);
  font-family: 'IBM Plex Mono', monospace;
  font-weight: 300;
}

/* ── Scenario strip (CT-02) ── */
.scenario-strip {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 16px;
  margin-bottom: 28px;
}

.scenario-card {
  background: rgba(16, 185, 129, 0.04);
  border: 1px solid rgba(16, 185, 129, 0.12);
  border-radius: 12px;
  padding: 16px;
  transition: border-color 0.3s, transform 0.3s, box-shadow 0.3s;
}

.scenario-card:hover {
  border-color: rgba(16, 185, 129, 0.3);
  transform: translateY(-3px);
  box-shadow: 0 12px 32px rgba(16, 185, 129, 0.08);
}

.scenario-num {
  font-family: 'Orbitron', sans-serif;
  font-weight: 700;
  font-size: 20px;
  color: rgba(16, 185, 129, 0.3);
  display: block;
  margin-bottom: 8px;
}

.scenario-text {
  margin: 0;
  font-size: 14px;
  color: var(--text-dim);
  line-height: 1.6;
  font-weight: 300;
}

/* ── Bento Grid (TC-02) ── */
.bento-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  grid-template-rows: auto auto;
  gap: 18px;
}

.bento-card.span-wide {
  grid-column: span 2;
}

.bento-card.span-full {
  grid-column: span 3;
}

.bento-card {
  background: rgba(255, 255, 255, 0.02);
  border: 1px solid rgba(16, 185, 129, 0.15);
  border-radius: 16px;
  padding: 20px;
  box-shadow: 0 18px 45px rgba(0, 0, 0, 0.3);
  backdrop-filter: blur(2px);
  transition: transform 0.3s cubic-bezier(0.16, 1, 0.3, 1), border-color 0.3s, box-shadow 0.3s;
  opacity: 0;
  transform: translateY(18px) scale(0.98);
  animation: riseFade 0.8s cubic-bezier(0.16, 1, 0.3, 1) forwards;
  animation-delay: var(--delay, 0s);
}

.bento-card:hover {
  transform: translateY(-4px);
  border-color: rgba(16, 185, 129, 0.4);
  box-shadow:
    0 24px 55px rgba(0, 0, 0, 0.4),
    0 0 0 1px rgba(16, 185, 129, 0.08),
    0 0 40px rgba(16, 185, 129, 0.06);
}

.bento-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  gap: 12px;
  margin-bottom: 12px;
}

.bento-title {
  font-weight: 700;
  font-size: 16px;
  font-family: 'Orbitron', sans-serif;
  letter-spacing: 0.04em;
  flex-shrink: 0;
}

.bento-illustration {
  flex-shrink: 0;
  width: 140px;
  border: 1px dashed rgba(16, 185, 129, 0.15);
  border-radius: 10px;
  padding: 6px;
  background: rgba(255, 255, 255, 0.02);
  transition: border-color 0.3s, box-shadow 0.3s;
}

.bento-card:hover .bento-illustration {
  border-color: rgba(16, 185, 129, 0.3);
  box-shadow: 0 0 20px rgba(16, 185, 129, 0.08);
}

.bento-illustration svg { display: block; width: 100%; height: auto; }

.bento-card ul {
  padding-left: 16px;
  margin: 0;
  display: flex;
  flex-direction: column;
  gap: 6px;
  line-height: 1.5;
  color: var(--text-dim);
  font-weight: 300;
}

/* ── Social Proof (CT-03) ── */
.social-proof {
  position: relative;
  z-index: 1;
  margin-top: 64px;
  text-align: center;
}

.proof-title {
  font-family: 'Orbitron', sans-serif;
  font-weight: 700;
  font-size: 20px;
  letter-spacing: 0.05em;
  margin: 0 0 8px;
}

.proof-desc {
  color: var(--text-dim);
  font-family: 'IBM Plex Mono', monospace;
  font-weight: 300;
  font-size: 14px;
  margin: 0 0 24px;
}

.proof-logos {
  display: flex;
  justify-content: center;
  gap: 32px;
  flex-wrap: wrap;
}

.proof-logo {
  opacity: 0.4;
  transition: opacity 0.3s;
}

.proof-logo:hover { opacity: 0.7; }

.logo-placeholder {
  width: 80px;
  height: 36px;
  border: 1px solid rgba(16, 185, 129, 0.15);
  border-radius: 8px;
  background: rgba(16, 185, 129, 0.04);
}

/* ── Animations ── */
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

/* ── Scrollbar ── */
.landing-shell::-webkit-scrollbar { width: 10px; }
.landing-shell::-webkit-scrollbar-track { background: var(--panel); border-left: 1px solid rgba(16, 185, 129, 0.1); }
.landing-shell::-webkit-scrollbar-thumb {
  background: var(--scrollbar-thumb);
  border-radius: 10px;
  border: 1px solid var(--bg-deep);
}
.landing-shell::-webkit-scrollbar-thumb:hover {
  background: rgba(16, 185, 129, 0.4);
}

/* ── Responsive (TC-01) ── */
@media (max-width: 960px) {
  .hero { grid-template-columns: 1fr; gap: 32px; }
  .hero-meta { grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); }
  .bento-grid { grid-template-columns: 1fr; }
  .bento-card.span-wide { grid-column: span 1; }
  .bento-card.span-full { grid-column: span 1; }
  .scenario-strip { grid-template-columns: 1fr; }
  .stats-row { gap: 20px; }
}

@media (max-width: 640px) {
  .landing-shell { padding: 32px 20px 48px; }
  .headline { font-size: clamp(28px, 8vw, 40px); }
  .stat-value { font-size: 24px; }
  .cta-row { flex-direction: column; }
  .btn { width: 100%; justify-content: center; }
}
</style>
