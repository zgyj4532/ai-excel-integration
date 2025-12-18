// 全局变量
let currentFile = null;
let currentWorkspace = null;
let currentFileId = null;
let ws = null;
let wsConnected = false;
let hotInstance = null; // Handsontable实例
let currentExcelData = [];

// DOM元素映射
const elements = {
    // 文件管理
    fileInput: document.getElementById('fileInput'),
    browseFileBtn: document.getElementById('browseFileBtn'),
    fileDropArea: document.getElementById('fileDropArea'),
    fileList: document.getElementById('fileList'),
    refreshFilesBtn: document.getElementById('refreshFilesBtn'),

    // Excel预览
    excelPreview: document.getElementById('excelPreview'),
    refreshPreviewBtn: document.getElementById('refreshPreviewBtn'),

    // AI功能
    aiCommand: document.getElementById('aiCommand'),
    sendAiCommandBtn: document.getElementById('sendAiCommandBtn'),
    aiResponseContainer: document.getElementById('aiResponseContainer'),
    quickCommand: document.getElementById('quickCommand'),
    sendQuickCommandBtn: document.getElementById('sendQuickCommandBtn'),

    // WebSocket
    toggleWebSocketBtn: document.getElementById('toggleWebSocketBtn'),
    wsStatus: document.getElementById('wsStatus'),

    // API状态
    apiStatus: document.getElementById('apiStatus'),

    // 格式设置
    rangeInput: document.getElementById('rangeInput'),
    applyBoldBtn: document.getElementById('applyBoldBtn'),
    applyItalicBtn: document.getElementById('applyItalicBtn'),
    applyUnderlineBtn: document.getElementById('applyUnderlineBtn'),
    applyBackgroundColorBtn: document.getElementById('applyBackgroundColorBtn'),
    applyTextColorBtn: document.getElementById('applyTextColorBtn'),
    applyBorderBtn: document.getElementById('applyBorderBtn'),
    applyFormattingBtn: document.getElementById('applyFormattingBtn'),

    // 模态框
    formatModal: document.getElementById('formatModal'),
    bgColorPicker: document.getElementById('bgColorPicker'),
    textColorPicker: document.getElementById('textColorPicker'),
    fontSizeInput: document.getElementById('fontSizeInput'),
    confirmFormatBtn: document.getElementById('confirmFormatBtn')
};

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    // 初始化Handsontable
    initializeHandsontable();

    // 绑定事件监听器
    bindEventListeners();

    // 检查API状态
    checkApiStatus();

    // 创建默认工作区并加载文件列表
    createDefaultWorkspace().then(() => {
        loadFileList();
    });

    // 初始化WebSocket状态
    updateWsStatus();
});

// 绑定所有事件监听器
function bindEventListeners() {
    // 文件上传相关
    elements.browseFileBtn.addEventListener('click', () => elements.fileInput.click());
    elements.fileInput.addEventListener('change', handleFileUpload);
    elements.fileDropArea.addEventListener('dragover', handleDragOver);
    elements.fileDropArea.addEventListener('drop', handleFileDrop);
    elements.refreshFilesBtn.addEventListener('click', loadFileList);

    // AI功能相关
    elements.sendAiCommandBtn.addEventListener('click', handleAiCommand);
    elements.sendQuickCommandBtn.addEventListener('click', handleQuickCommand);
    elements.aiCommand.addEventListener('keypress', function(e) {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleAiCommand();
        }
    });
    elements.quickCommand.addEventListener('keypress', function(e) {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleQuickCommand();
        }
    });

    // 预览和WebSocket
    elements.refreshPreviewBtn.addEventListener('click', refreshPreview);
    elements.toggleWebSocketBtn.addEventListener('click', toggleWebSocket);

    // 格式设置
    elements.applyFormattingBtn.addEventListener('click', handleApplyFormatting);
    elements.confirmFormatBtn.addEventListener('click', confirmFormatChanges);

    // 特定格式按钮
    elements.applyBoldBtn.addEventListener('click', () => applyFormat('bold'));
    elements.applyItalicBtn.addEventListener('click', () => applyFormat('italic'));
    elements.applyUnderlineBtn.addEventListener('click', () => applyFormat('underline'));
    elements.applyBackgroundColorBtn.addEventListener('click', () => openFormatModal('background'));
    elements.applyTextColorBtn.addEventListener('click', () => openFormatModal('text'));
    elements.applyBorderBtn.addEventListener('click', () => applyFormat('border'));

    // AI功能卡片
    document.querySelectorAll('.feature-card').forEach(card => {
        card.addEventListener('click', function() {
            const feature = this.getAttribute('data-feature');
            handleAiFeature(feature);
        });
    });
}

// 创建默认工作区
async function createDefaultWorkspace() {
    try {
        // 获取用户ID，如果没有则生成一个
        let userId = localStorage.getItem('userId');
        if (!userId) {
            userId = 'user_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
            localStorage.setItem('userId', userId);
        }

        // 检查是否存在默认工作区
        const response = await fetch(`/api/files/workspaces/user/${userId}`);
        const result = await response.json();

        if (result.success && result.data && result.data.length > 0) {
            // 使用第一个工作区
            currentWorkspace = result.data[0];
        } else {
            // 创建默认工作区
            const workspaceData = {
                name: '我的工作区',
                userId: userId,
                description: '默认工作区'
            };

            const createResponse = await fetch('/api/files/workspace/create', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(workspaceData)
            });

            const createResult = await createResponse.json();

            if (createResult.success) {
                currentWorkspace = createResult.data;
            } else {
                console.error('创建默认工作区失败:', createResult.error);
                return false;
            }
        }

        return true;
    } catch (error) {
        console.error('创建默认工作区时出错:', error);
        return false;
    }
}

// 初始化Handsontable
function initializeHandsontable() {
    if (!elements.excelPreview) return;

    if (hotInstance) {
        hotInstance.destroy();
    }

    hotInstance = new Handsontable(elements.excelPreview, {
        data: [],
        colHeaders: true,
        rowHeaders: true,
        height: '100%',
        stretchH: 'all',
        manualRowResize: true,
        manualColumnResize: true,
        manualRowMove: true,
        manualColumnMove: true,
        licenseKey: 'non-commercial-and-evaluation',
        contextMenu: true,
        afterChange: function(changes, source) {
            if (source !== 'loadData' && wsConnected && ws) {
                // 如果启用了WebSocket，发送数据变更
                const changeData = {
                    type: 'data_change',
                    changes: changes,
                    timestamp: new Date().toISOString()
                };
                ws.send(JSON.stringify(changeData));
            }
        }
    });
}

// 刷新预览
function refreshPreview() {
    if (currentFileId) {
        loadExcelPreview(currentFileId);
    } else if (currentFile) {
        uploadAndPreviewFile(currentFile);
    } else {
        addResponseMessage('请先上传文件', 'system');
    }
}

// 处理文件上传
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    await uploadAndPreviewFile(file);
}

// 拖拽上传处理
function handleDragOver(event) {
    event.preventDefault();
    event.stopPropagation();
    elements.fileDropArea.style.borderColor = '#4361ee';
    elements.fileDropArea.style.backgroundColor = '#eef2ff';
}

function handleFileDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    elements.fileDropArea.style.borderColor = '#dee2e6';
    elements.fileDropArea.style.backgroundColor = '#f8f9fa';

    const file = event.dataTransfer.files[0];
    if (file && file.type.match('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet|application/vnd.ms-excel|text/csv')) {
        uploadAndPreviewFile(file);
    }
}

// 上传并预览文件
async function uploadAndPreviewFile(file) {
    if (!currentWorkspace) {
        addResponseMessage('请先创建或选择工作区', 'system');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    addResponseMessage(`📄 正在上传文件: ${file.name}`, 'system');

    try {
        // 上传文件到工作区
        const response = await fetch(`/api/files/workspace/${currentWorkspace.id}/upload?userId=user_123`, {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            currentFile = file;
            currentFileId = result.data.id; // 使用数据库中的文件ID
            addResponseMessage(`✅ 文件上传成功: ${file.name}`, 'system');

            // 使用文件预览API获取数据并显示在表格中
            const previewResponse = await fetch('/api/excel/preview', {
                method: 'POST',
                body: formData
            });

            const previewResult = await previewResponse.json();

            if (previewResult.success) {
                currentExcelData = previewResult.data || [];

                if (hotInstance) {
                    hotInstance.loadData(currentExcelData);
                    hotInstance.render();
                    addResponseMessage(`✅ 加载了 ${currentExcelData.length} 行数据`, 'system');
                }
            } else {
                addResponseMessage(`❌ 预览加载失败: ${previewResult.error}`, 'system');
            }

            // 刷新文件列表
            await loadFileList();
        } else {
            addResponseMessage(`❌ 上传失败: ${result.error}`, 'system');
        }
    } catch (error) {
        addResponseMessage(`❌ 上传错误: ${error.message}`, 'system');
    }
}

// 加载Excel预览
async function loadExcelPreview(fileId) {
    if (!fileId) return;

    // 由于后端没有直接通过ID获取文件数据的API，我们从文件对象中加载
    // 这里需要重新上传文件以获取预览数据
    if (currentFile) {
        const formData = new FormData();
        formData.append('file', currentFile);

        try {
            const response = await fetch('/api/excel/preview', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                currentExcelData = result.data || [];

                if (hotInstance) {
                    hotInstance.loadData(currentExcelData);
                    hotInstance.render();
                    addResponseMessage(`✅ 加载了 ${currentExcelData.length} 行数据`, 'system');
                }
            } else {
                addResponseMessage(`❌ 加载预览失败: ${result.error}`, 'system');
            }
        } catch (error) {
            addResponseMessage(`❌ 加载预览错误: ${error.message}`, 'system');
        }
    }
}

// 加载文件列表
async function loadFileList() {
    if (!currentWorkspace) {
        // 尝试创建默认工作区
        await createDefaultWorkspace();
    }

    if (!currentWorkspace) {
        elements.fileList.innerHTML = '<div class="text-center text-muted py-3">请先创建工作区</div>';
        return;
    }

    try {
        const response = await fetch(`/api/files/workspace/${currentWorkspace.id}/files`);
        const result = await response.json();

        if (result.success) {
            displayFileList(result.data || []);
        } else {
            elements.fileList.innerHTML = '<div class="text-center text-muted py-3">加载文件列表失败</div>';
        }
    } catch (error) {
        elements.fileList.innerHTML = '<div class="text-center text-muted py-3">加载文件列表失败</div>';
        console.error('加载文件列表错误:', error);
    }
}

// 显示文件列表
function displayFileList(files) {
    if (files.length === 0) {
        elements.fileList.innerHTML = '<div class="text-center text-muted py-3">暂无文件</div>';
        return;
    }

    let html = '';
    files.forEach(file => {
        html += `
            <div class="template-item" data-file-id="${file.id}">
                <div class="d-flex justify-content-between align-items-center">
                    <div>
                        <div><i class="bi bi-file-earmark-excel text-success me-1"></i>${file.fileName}</div>
                        <small class="text-muted">${new Date(file.uploadTime).toLocaleString()}</small>
                    </div>
                    <div>
                        <button class="btn btn-sm btn-outline-primary me-1" onclick="loadFile(${file.id}, '${file.fileName}')">
                            <i class="bi bi-eye"></i>
                        </button>
                        <button class="btn btn-sm btn-outline-danger" onclick="deleteFile(${file.id}, event)">
                            <i class="bi bi-trash"></i>
                        </button>
                    </div>
                </div>
            </div>
        `;
    });

    elements.fileList.innerHTML = html;
}

// 加载特定文件
async function loadFile(fileId, fileName) {
    // 由于后端没有直接通过ID获取文件数据的API，我们只设置文件ID
    // 实际的预览需要重新上传文件
    currentFileId = fileId;
    addResponseMessage(`已选择文件: ${fileName}`, 'system');
}

// 删除文件
async function deleteFile(fileId, event) {
    event.stopPropagation();

    if (!confirm('确定要删除这个文件吗？')) return;

    // 由于后端没有提供删除文件的API，我们只能通过工作区管理
    // 这里我们只是显示提醒
    addResponseMessage('当前版本不支持通过API删除文件', 'system');
}

// 处理AI命令
async function handleAiCommand() {
    const command = elements.aiCommand.value.trim();
    if (!command) {
        addResponseMessage('请输入命令', 'system');
        return;
    }

    addResponseMessage(command, 'user', '您');
    elements.aiCommand.value = '';

    await sendAiCommandHttp(command);
}

// 快速命令
async function handleQuickCommand() {
    const command = elements.quickCommand.value.trim();
    if (!command) return;

    addResponseMessage(command, 'user', '快速命令');
    elements.quickCommand.value = '';

    await sendAiCommandHttp(command);
}

// 通过HTTP发送AI命令
async function sendAiCommandHttp(command) {
    if (!currentFile) {
        addResponseMessage('请先上传文件', 'system');
        return;
    }

    addResponseMessage('🤖 AI正在处理...', 'ai');

    const formData = new FormData();
    formData.append('file', currentFile);
    formData.append('command', command);

    try {
        const response = await fetch('/api/ai/excel-with-ai', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            addResponseMessage(result.aiResponse, 'ai');

            // 如果AI返回了需要更新表格的指令，执行它们
            if (result.excelInstruction) {
                executeExcelInstruction(result.excelInstruction);
            }
        } else {
            addResponseMessage(`❌ AI处理失败: ${result.error}`, 'system');
        }
    } catch (error) {
        addResponseMessage(`❌ AI请求错误: ${error.message}`, 'system');
    }
}

// 执行Excel指令
function executeExcelInstruction(instruction) {
    if (!hotInstance) return;

    try {
        // 这里可以根据AI返回的指令格式来更新表格
        // 示例：AI可能返回要更新的单元格坐标和值
        if (instruction.type === 'update_cells' && instruction.cells) {
            instruction.cells.forEach(cell => {
                const { row, col, value } = cell;
                hotInstance.setDataAtCell(row, col, value);
            });
            hotInstance.render();
            addResponseMessage('✅ 表格已根据AI建议更新', 'system');
        }
    } catch (error) {
        addResponseMessage(`❌ 执行Excel指令错误: ${error.message}`, 'system');
    }
}

// 处理AI功能卡片
function handleAiFeature(feature) {
    let command = '';

    switch(feature) {
        case 'formula':
            command = '根据当前数据帮我生成合适的Excel公式';
            break;
        case 'analysis':
            command = '分析当前数据的主要趋势和模式';
            break;
        case 'format':
            command = '推荐适合当前数据的格式化方案';
            break;
        case 'visualization':
            command = '推荐适合当前数据的可视化图表';
            break;
    }

    if (command) {
        elements.aiCommand.value = command;
        handleAiCommand();
    }
}

// 切换WebSocket
function toggleWebSocket() {
    if (wsConnected) {
        disconnectWebSocket();
        elements.toggleWebSocketBtn.innerHTML = '<i class="bi bi-broadcast me-1"></i>实时模式';
        addResponseMessage('WebSocket已断开', 'system');
    } else {
        connectWebSocket();
        elements.toggleWebSocketBtn.innerHTML = '<i class="bi bi-broadcast me-1"></i>实时模式';
    }
}

// 连接WebSocket
function connectWebSocket() {
    if (ws && wsConnected) return;

    const clientId = 'client_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    const wsUrl = `ws://${window.location.host}/websocket/${clientId}`;

    try {
        ws = new WebSocket(wsUrl);

        ws.onopen = function(event) {
            wsConnected = true;
            updateWsStatus();
            addResponseMessage('🔌 WebSocket连接已建立', 'system');
        };

        ws.onmessage = function(event) {
            try {
                const data = JSON.parse(event.data);

                if (data.type === 'ai_response') {
                    addResponseMessage(data.message, 'ai');
                } else if (data.type === 'excel_update') {
                    // 更新Excel表格
                    if (hotInstance && data.data) {
                        hotInstance.loadData(data.data);
                        hotInstance.render();
                    }
                } else if (data.type === 'system_message') {
                    addResponseMessage(data.message, 'system');
                } else {
                    addResponseMessage(`WebSocket消息: ${event.data}`, 'system');
                }
            } catch (e) {
                addResponseMessage(`WebSocket数据解析错误: ${event.data}`, 'system');
            }
        };

        ws.onclose = function(event) {
            wsConnected = false;
            updateWsStatus();
            addResponseMessage('🔌 WebSocket连接已关闭', 'system');
        };

        ws.onerror = function(error) {
            wsConnected = false;
            updateWsStatus();
            addResponseMessage(`WebSocket错误: ${error.message}`, 'system');
        };
    } catch (error) {
        addResponseMessage(`WebSocket连接失败: ${error.message}`, 'system');
    }
}

// 断开WebSocket连接
function disconnectWebSocket() {
    if (ws) {
        ws.close();
        ws = null;
    }
    wsConnected = false;
    updateWsStatus();
}

// 更新WebSocket状态显示
function updateWsStatus() {
    if (wsConnected) {
        elements.wsStatus.textContent = 'WebSocket在线';
        elements.wsStatus.className = 'status-indicator status-connected';
    } else {
        elements.wsStatus.textContent = 'WebSocket离线';
        elements.wsStatus.className = 'status-indicator status-disconnected';
    }
}

// 检查API状态
async function checkApiStatus() {
    try {
        const response = await fetch('/api/status');
        const result = await response.json();

        if (result.hasApiKey) {
            elements.apiStatus.textContent = 'API在线';
            elements.apiStatus.className = 'status-indicator status-connected';
        } else {
            elements.apiStatus.textContent = 'API离线';
            elements.apiStatus.className = 'status-indicator status-disconnected';
        }
    } catch (error) {
        elements.apiStatus.textContent = 'API状态未知';
        elements.apiStatus.className = 'status-indicator status-disconnected';
    }
}

// 应用格式设置
function applyFormat(formatType) {
    const range = elements.rangeInput.value.trim();
    if (!range) {
        addResponseMessage('请先输入单元格范围 (例如: A1:B5)', 'system');
        return;
    }

    if (!currentFile) {
        addResponseMessage('请先上传文件', 'system');
        return;
    }

    // 根据格式类型执行相应的格式设置
    switch(formatType) {
        case 'bold':
            applyCellFormat(range, { fontBold: true });
            break;
        case 'italic':
            applyCellFormat(range, { fontItalic: true });
            break;
        case 'underline':
            applyCellFormat(range, { fontUnderline: true });
            break;
        case 'border':
            applyCellFormat(range, {
                borderLeft: 'THIN',
                borderRight: 'THIN',
                borderTop: 'THIN',
                borderBottom: 'THIN'
            });
            break;
    }
}

// 应用单元格格式
async function applyCellFormat(range, formatOptions) {
    if (!currentFile) {
        addResponseMessage('请先上传文件', 'system');
        return;
    }

    try {
        // 解析范围
        const rangeObj = parseRange(range);

        const formData = new FormData();
        formData.append('file', currentFile);
        formData.append('sheetName', 'Sheet1'); // 默认工作表名称
        formData.append('startRow', rangeObj.startRow);
        formData.append('startCol', rangeObj.startCol);
        formData.append('endRow', rangeObj.endRow);
        formData.append('endCol', rangeObj.endCol);

        // 以JSON格式发送格式选项
        const response = await fetch('/api/excel/format-range', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            addResponseMessage(`✅ 已应用格式到范围 ${range}`, 'system');
            // 刷新预览以显示格式更改
            await loadExcelPreview(currentFileId);
        } else {
            addResponseMessage(`❌ 格式应用失败: ${result.error}`, 'system');
        }
    } catch (error) {
        addResponseMessage(`❌ 格式应用错误: ${error.message}`, 'system');
    }
}

// 打开格式模态框
function openFormatModal(type) {
    const range = elements.rangeInput.value.trim();
    if (!range) {
        addResponseMessage('请先输入单元格范围 (例如: A1:B5)', 'system');
        return;
    }

    if (!currentFile) {
        addResponseMessage('请先上传文件', 'system');
        return;
    }

    // 显示模态框
    const modal = new bootstrap.Modal(elements.formatModal);
    modal.show();
}

// 确认格式更改
async function confirmFormatChanges() {
    const bgColor = elements.bgColorPicker.value;
    const textColor = elements.textColorPicker.value;
    const fontSize = parseInt(elements.fontSizeInput.value);

    const range = elements.rangeInput.value.trim();
    if (!range) {
        addResponseMessage('请先输入单元格范围', 'system');
        return;
    }

    if (!currentFile) {
        addResponseMessage('请先上传文件', 'system');
        return;
    }

    try {
        // 解析范围
        const rangeObj = parseRange(range);

        // 创建格式选项对象
        const formatOptions = {
            backgroundColor: bgColor,
            fontColor: textColor,
            fontSize: fontSize
        };

        const formData = new FormData();
        formData.append('file', currentFile);
        formData.append('sheetName', 'Sheet1'); // 默认工作表名称
        formData.append('startRow', rangeObj.startRow);
        formData.append('startCol', rangeObj.startCol);
        formData.append('endRow', rangeObj.endRow);
        formData.append('endCol', rangeObj.endCol);

        // 添加格式选项作为JSON字符串
        for (const [key, value] of Object.entries(formatOptions)) {
            formData.append(key, value);
        }

        const response = await fetch('/api/excel/format-range', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            addResponseMessage(`✅ 已应用格式到范围 ${range}`, 'system');
            // 刷新预览以显示格式更改
            await loadExcelPreview(currentFileId);

            // 关闭模态框
            const modal = bootstrap.Modal.getInstance(elements.formatModal);
            if (modal) modal.hide();
        } else {
            addResponseMessage(`❌ 格式应用失败: ${result.error}`, 'system');
        }
    } catch (error) {
        addResponseMessage(`❌ 格式应用错误: ${error.message}`, 'system');
    }
}

// 处理应用格式按钮
async function handleApplyFormatting() {
    const range = elements.rangeInput.value.trim();
    if (!range) {
        addResponseMessage('请先输入单元格范围 (例如: A1:B5)', 'system');
        return;
    }

    if (!currentFile) {
        addResponseMessage('请先上传文件', 'system');
        return;
    }

    // 获取当前选中范围的格式信息（批量获取）
    const formData = new FormData();
    formData.append('file', currentFile);
    formData.append('range', JSON.stringify(parseRange(range)));

    try {
        const response = await fetch('/api/excel/bulk-cell-format', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            addResponseMessage(`✅ 获取了范围 ${range} 的格式信息`, 'system');
            console.log('格式信息:', result.formatData);
        } else {
            addResponseMessage(`❌ 获取格式信息失败: ${result.error}`, 'system');
        }
    } catch (error) {
        addResponseMessage(`❌ 获取格式信息错误: ${error.message}`, 'system');
    }
}

// 解析范围字符串
function parseRange(rangeStr) {
    // 解析类似 "A1:B5" 的范围字符串
    const parts = rangeStr.split(':');
    if (parts.length !== 2) {
        throw new Error('无效的范围格式，应为 "A1:B5" 格式');
    }

    const start = parseCell(parts[0]);
    const end = parseCell(parts[1]);

    return {
        startRow: start.row,
        startCol: start.col,
        endRow: end.row,
        endCol: end.col
    };
}

// 解析单元格坐标
function parseCell(cellStr) {
    const col = cellStr.match(/[A-Z]+/)[0];
    const row = parseInt(cellStr.match(/\d+/)[0]) - 1; // 转换为0基索引

    // 将列字母转换为数字索引 (A=0, B=1, ...)
    let colIndex = 0;
    for (let i = 0; i < col.length; i++) {
        colIndex = colIndex * 26 + (col.charCodeAt(i) - 'A'.charCodeAt(0));
    }

    return { row: row, col: colIndex };
}

// 添加响应消息
function addResponseMessage(content, type, sender = null) {
    if (!elements.aiResponseContainer) return;

    const messageDiv = document.createElement('div');
    messageDiv.className = `message-${type} p-3 rounded mb-2`;

    if (type === 'user') {
        messageDiv.style.backgroundColor = '#dbf0ff';
        messageDiv.style.alignSelf = 'flex-end';
        messageDiv.style.marginLeft = '20%';
    } else if (type === 'ai') {
        messageDiv.style.backgroundColor = '#e9ecef';
        messageDiv.style.alignSelf = 'flex-start';
        messageDiv.style.marginRight = '20%';
    } else if (type === 'system') {
        messageDiv.style.backgroundColor = '#d1e7dd';
        messageDiv.style.alignSelf = 'center';
        messageDiv.style.maxWidth = '100%';
        messageDiv.style.textAlign = 'center';
        messageDiv.style.fontStyle = 'italic';
    }

    if (sender) {
        const senderSpan = document.createElement('strong');
        senderSpan.textContent = `${sender}: `;
        messageDiv.appendChild(senderSpan);
    }

    // 将换行符转换为<br>
    const contentParts = content.split('\n');
    contentParts.forEach((part, index) => {
        messageDiv.appendChild(document.createTextNode(part));
        if (index < contentParts.length - 1) {
            messageDiv.appendChild(document.createElement('br'));
        }
    });

    elements.aiResponseContainer.appendChild(messageDiv);

    // 滚动到底部
    elements.aiResponseContainer.scrollTop = elements.aiResponseContainer.scrollHeight;
}

// 页面卸载时断开连接
window.addEventListener('beforeunload', function() {
    if (ws) {
        ws.close();
    }
});

// 菜单导航功能
document.querySelectorAll('#menu .nav-link').forEach(link => {
    link.addEventListener('click', function(e) {
        e.preventDefault();

        // 移除所有活动状态
        document.querySelectorAll('#menu .nav-link').forEach(item => {
            item.classList.remove('active');
        });

        // 添加当前活动状态
        this.classList.add('active');

        // 可以在这里添加页面内容切换逻辑
        const target = this.getAttribute('href');
        console.log(`导航到: ${target}`);
    });
});