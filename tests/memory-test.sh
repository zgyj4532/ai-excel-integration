#!/bin/bash
# 内存测试脚本 - 用于验证2c2g部署可行性
# 使用方法: ./tests/memory-test.sh [server_port]

PORT=${1:-8080}
BASE_URL="http://localhost:${PORT}"

echo "=== ai-excel-integration 内存测试 ==="
echo "目标: 2GB RAM 服务器"
echo "服务地址: ${BASE_URL}"
echo ""

# 1. 检查服务是否运行
echo "[1/5] 检查服务健康状态..."
HEALTH=$(curl -s -o /dev/null -w "%{http_code}" "${BASE_URL}/api/health" 2>/dev/null)
if [ "$HEALTH" = "200" ]; then
    echo "  ✓ 服务运行正常 (HTTP 200)"
else
    echo "  ✗ 服务未运行或不可达 (HTTP ${HEALTH})"
    echo "  请先启动服务: mvn spring-boot:run"
    exit 1
fi

# 2. 检查JVM进程内存
echo ""
echo "[2/5] 检查JVM进程内存..."
PID=$(pgrep -f "ai-excel-integration" 2>/dev/null)
if [ -n "$PID" ]; then
    RSS=$(ps -o rss= -p "$PID" 2>/dev/null | tr -d ' ')
    VSZ=$(ps -o vsz= -p "$PID" 2>/dev/null | tr -d ' ')
    RSS_MB=$((RSS / 1024))
    VSZ_MB=$((VSZ / 1024))
    echo "  PID: ${PID}"
    echo "  RSS (物理内存): ${RSS_MB} MB"
    echo "  VSZ (虚拟内存): ${VSZ_MB} MB"

    if [ "$RSS_MB" -lt 512 ]; then
        echo "  ✓ 内存使用在512MB以内，适合2c2g部署"
    elif [ "$RSS_MB" -lt 768 ]; then
        echo "  ⚠ 内存使用在512-768MB，需要关注"
    else
        echo "  ✗ 内存使用超过768MB，不适合2c2g部署"
    fi
else
    echo "  ⚠ 未找到JVM进程 (可能通过IDE启动)"
fi

# 3. 测试Actuator端点
echo ""
echo "[3/5] 检查Actuator指标..."
METRICS=$(curl -s "${BASE_URL}/actuator/health" 2>/dev/null)
if [ -n "$METRICS" ]; then
    echo "  ✓ Actuator可用"
    echo "  响应: ${METRICS}" | head -c 200
    echo ""
else
    echo "  ⚠ Actuator不可用 (需要添加spring-boot-starter-actuator依赖)"
fi

# 4. 测试文件上传内存占用
echo ""
echo "[4/5] 测试文件上传..."
UPLOAD_RESULT=$(curl -s -o /dev/null -w "%{http_code}" \
    -X POST "${BASE_URL}/api/upload" \
    -F "file=@/dev/null;filename=empty.xlsx" \
    -H "Content-Type: multipart/form-data" 2>/dev/null)
echo "  空文件上传: HTTP ${UPLOAD_RESULT}"

# 5. 检查响应压缩
echo ""
echo "[5/5] 检查响应压缩..."
COMPRESSED=$(curl -s -o /dev/null -w "%{http_code}" \
    -H "Accept-Encoding: gzip" \
    "${BASE_URL}/api/health" 2>/dev/null)
echo "  压缩请求: HTTP ${COMPRESSED}"

echo ""
echo "=== 测试完成 ==="
echo ""
echo "优化建议:"
echo "  1. 确保JVM参数: -Xms256m -Xmx768m -XX:MaxMetaspaceSize=128m"
echo "  2. 使用G1GC: -XX:+UseG1GC -XX:MaxGCPauseMillis=200"
echo "  3. 启用字符串去重: -XX:+UseStringDeduplication"
echo "  4. 设置OOM时dump: -XX:+HeapDumpOnOutOfMemoryError"
