#!/bin/bash
# Frontend verification script for optimization tracker items
# Validates that all required changes are present in the codebase

PASS=0
FAIL=0

check() {
  local desc="$1"
  local file="$2"
  local pattern="$3"
  if grep -q "$pattern" "$file" 2>/dev/null; then
    echo "  PASS: $desc"
    PASS=$((PASS + 1))
  else
    echo "  FAIL: $desc"
    FAIL=$((FAIL + 1))
  fi
}

echo "=== VB-01: 首屏信息降噪，增加留白 ==="
check "锚点导航组件" "frontend/src/pages/Landing.vue" "anchor-nav"
check "hover交互meta标签" "frontend/src/pages/Landing.vue" "meta-label-wrapper"
check "滚动监听显示导航" "frontend/src/pages/Landing.vue" "showNav"

echo ""
echo "=== VB-02: 品牌色调整为数据绿 ==="
check "数据绿色值 #10b981" "frontend/src/pages/Landing.vue" "#10b981"
check "绿色渐变SVG" "frontend/src/pages/Landing.vue" "rgba(16, 185, 129"

echo ""
echo "=== VB-03: 字体层级清晰化 ==="
check "字重900用于标题" "frontend/src/pages/Landing.vue" "font-weight: 900"
check "字重300用于副文本" "frontend/src/pages/Landing.vue" "font-weight: 300"
check "Orbitron字体加载900" "frontend/src/pages/Landing.vue" "Orbitron:wght@400;500;700;900"

echo ""
echo "=== IX-01: CTA按钮强化与动效 ==="
check "播放图标SVG" "frontend/src/pages/Landing.vue" "play-icon"
check "hover transform Y" "frontend/src/pages/Landing.vue" "translateY(-3px)"
check "箭头动画" "frontend/src/pages/Landing.vue" "translateX(5px)"

echo ""
echo "=== IX-02: 核心能力卡片hover交互 ==="
check "卡片hover上浮" "frontend/src/pages/Landing.vue" "translateY(-4px)"
check "边框发光效果" "frontend/src/pages/Landing.vue" "0 0 0 1px rgba(16, 185, 129"

echo ""
echo "=== CT-01: 价值主张量化 ==="
check "统计数值组件" "frontend/src/pages/Landing.vue" "stat-value"
check "i18n统计键 - 中文" "frontend/src/i18n.ts" "landingStat1:"
check "i18n统计键 - 英文" "frontend/src/i18n.ts" "landingStat1Desc:"

echo ""
echo "=== CT-02: 功能描述场景化 ==="
check "场景卡片组件" "frontend/src/pages/Landing.vue" "scenario-card"
check "i18n场景键 - 中文" "frontend/src/i18n.ts" "landingFeatureScenario1:"

echo ""
echo "=== CT-03: 增加社会证明区块 ==="
check "社会证明区块" "frontend/src/pages/Landing.vue" "social-proof"
check "logo占位组件" "frontend/src/pages/Landing.vue" "logo-placeholder"

echo ""
echo "=== TC-01: 响应式与性能优化 ==="
check "移动端960px断点" "frontend/src/pages/Landing.vue" "@media (max-width: 960px)"
check "移动端640px断点" "frontend/src/pages/Landing.vue" "@media (max-width: 640px)"
check "CTA列堆叠" "frontend/src/pages/Landing.vue" "flex-direction: column"

echo ""
echo "=== TC-02: Bento Box布局 ==="
check "Bento网格布局" "frontend/src/pages/Landing.vue" "bento-grid"
check "宽卡片span" "frontend/src/pages/Landing.vue" "span-wide"
check "grid-template-columns" "frontend/src/pages/Landing.vue" "grid-template-columns: repeat(3"

echo ""
echo "=== 前端撤回功能 (BE-08) ==="
check "撤回API调用" "frontend/src/services/aiService.ts" "undoLastOperation"
check "撤回按钮组件" "frontend/src/components/workspace/WorkspaceTabs.vue" "undo-btn"
check "撤回状态管理" "frontend/src/pages/Workspace.vue" "handleUndoClick"
check "撤回i18n中文" "frontend/src/i18n.ts" "undoFile: '撤回'"
check "撤回i18n英文" "frontend/src/i18n.ts" "undoFile: 'Undo'"

echo ""
echo "=== Backend undo/redo API ==="
check "undo端点" "src/main/java/com/example/aiexcel/controller/OperationHistoryController.java" "undo/{fileId}"
check "redo端点" "src/main/java/com/example/aiexcel/controller/OperationHistoryController.java" "redo/{fileId}"

echo ""
echo "=== 结果 ==="
echo "通过: $PASS"
echo "失败: $FAIL"
if [ $FAIL -eq 0 ]; then
  echo "全部通过!"
  exit 0
else
  echo "有失败项，请检查"
  exit 1
fi
