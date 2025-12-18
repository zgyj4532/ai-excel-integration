# AI Excel Integration

一个整合了AI功能的Excel处理工具，使用Spring Boot后端与Qwen API集成，支持通过自然语言处理Excel数据。

## 功能特性

- 上传Excel文件进行分析
- 通过自然语言与Excel数据交互
- 生成Excel公式
- 数据分析和洞察
- WebSocket实时通信
- 与Qwen API集成

## 技术栈

- Spring Boot 3.2.0
- Java 17
- Apache POI (Excel操作)
- Qwen API (AI集成)
- WebSocket (实时通信)
- Bootstrap 5 (前端UI)

## 环境要求

- Java 17+
- Maven 3.6+
- Qwen API密钥

## 快速开始

1. 克隆项目
   ```bash
   git clone <repository-url>
   cd ai-excel-integration
   ```

2. 配置环境变量
   - 复制环境配置文件:
     ```bash
     cp .env.example .env
     ```
   - 编辑 `.env` 文件，填入您的配置:
     ```bash
     # 编辑 .env 文件
     QWEN_API_KEY=your_actual_qwen_api_key_here
     QWEN_MODEL_NAME=qwen-max  # 可选: qwen-turbo, qwen-plus, qwen-max
     ```

   或者通过环境变量设置:
   ```bash
   export QWEN_API_KEY=your_qwen_api_key_here
   export QWEN_MODEL_NAME=qwen-max
   ```

   或者在 `application.properties` 中设置:
   ```properties
   qwen.api.api-key=your_qwen_api_key_here
   qwen.api.default-model=qwen-max
   ```

3. 编译项目
   ```bash
   mvn clean compile
   ```

4. 运行应用
   ```bash
   mvn spring-boot:run
   ```

5. 访问应用
   打开浏览器访问 `http://localhost:8080`

## API端点

- `POST /api/upload` - 上传Excel文件
- `POST /api/ai/excel-with-ai` - 用AI处理Excel命令
- `POST /api/ai/generate-formula` - 生成Excel公式
- `POST /api/ai/excel-analyze` - 分析Excel数据
- `POST /api/ai/suggest-charts` - 获取图表建议
- `GET /api/health` - 检查健康状态
- `WebSocket /websocket/{clientId}` - 实时通信

## 配置

在 `application.properties` 中可以配置:

```properties
# 服务器端口
server.port=8080

# Redis配置 (可选)
spring.redis.host=localhost
spring.redis.port=6379

# Qwen API配置
qwen.api.base-url=https://dashscope.aliyuncs.com/compatible-mode/v1
qwen.api.default-model=qwen-max
```

## 前端界面

访问 `http://localhost:8080` 可以使用图形化界面:

1. 上传Excel文件
2. 输入自然语言命令 (例如："分析销售数据", "计算平均值") 
3. 查看AI生成的响应和建议的公式
4. 使用实时模式与WebSocket通信

## 项目结构

```
src/main/java/com/example/aiexcel/
├── AiExcelIntegrationApplication.java  # 主应用类
├── config/                            # 配置类
│   ├── ServiceConfig.java
│   └── WebSocketConfig.java
├── controller/                        # 控制器
│   └── AiExcelController.java
├── dto/                               # 数据传输对象
│   ├── AiRequest.java
│   └── AiResponse.java
├── service/                           # 业务逻辑
│   ├── AiExcelIntegrationService.java
│   ├── ai/
│   │   ├── AiService.java
│   │   └── impl/QwenAiService.java
│   └── excel/
│       ├── ExcelService.java
│       └── impl/ExcelServiceImpl.java
└── websocket/                         # WebSocket处理器
    └── ExcelWebSocketHandler.java
```

## 集成说明

本项目整合了以下三个开源项目的核心功能:

1. **excel-ai-assistant** - 提供本地AI集成方式，本项目改用在线API
2. **ExcelFlow** - 提供Web界面和实时数据交互功能 
3. **smart-excel-ai** - 提供公式生成功能

所有AI功能统一使用Qwen API，通过配置的API密钥进行认证。