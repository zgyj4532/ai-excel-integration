#!/bin/bash

# 加载 .env 文件，只处理非注释行
if [ -f .env ]; then
    # 读取 .env 文件内容并设置环境变量
    while IFS= read -r line; do
        # 跳过空行和注释
        if [[ -z "$line" || "$line" =~ ^[[:space:]]*# ]]; then
            continue
        fi

        # 解析 KEY=VALUE 行
        if [[ "$line" =~ ^[[:space:]]*([^=]+)[[:space:]]*=[[:space:]]*(.*)[[:space:]]*$ ]]; then
            key="${BASH_REMATCH[1]}"
            value="${BASH_REMATCH[2]}"

            # 移除引号
            value="${value%\"}"
            value="${value#\"}"
            value="${value%\'}"
            value="${value#\'}"

            # 设置环境变量
            export "$key=$value"
        fi
    done < .env
fi

echo "Starting AI Excel Integration Application..."
echo "Using QWEN_MODEL_NAME: ${QWEN_MODEL_NAME:-default}"
echo "Using SERVER_PORT: ${SERVER_PORT:-8080}"

# 运行应用
mvn spring-boot:run