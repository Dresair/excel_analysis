// 全局变量
let currentSessionId = null;
let currentTaskId = null;
let isDarkMode = localStorage.getItem('darkMode') === 'true';
let isUploading = false; // 防止重复上传

// 页面初始化
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
    setupEventListeners();
    loadSettings();
    applyTheme();
});

// 初始化应用
function initializeApp() {
    console.log('PPT生成工具已启动');
    refreshOutputFiles();
}

// 设置事件监听器
function setupEventListeners() {
    // 文件输入
    const fileInput = document.getElementById('file-input');
    fileInput.addEventListener('change', handleFileSelect);
    
    // 拖拽上传
    const uploadArea = document.getElementById('upload-area');
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleFileDrop);
    // 移除了 click 事件，改为在 HTML 中的特定元素上触发
    
    // 聊天输入
    const chatInput = document.getElementById('chat-input');
    chatInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            sendMessage();
        }
    });
    
    // 模态框关闭
    window.addEventListener('click', function(e) {
        const modal = document.getElementById('settings-modal');
        if (e.target === modal) {
            closeSettings();
        }
    });
}

// =============== 主题切换 ===============
function toggleTheme() {
    isDarkMode = !isDarkMode;
    localStorage.setItem('darkMode', isDarkMode);
    applyTheme();
}

function applyTheme() {
    if (isDarkMode) {
        document.documentElement.setAttribute('data-theme', 'dark');
        document.getElementById('theme-text').textContent = '浅色模式';
    } else {
        document.documentElement.removeAttribute('data-theme');
        document.getElementById('theme-text').textContent = '深色模式';
    }
}

// =============== 文件上传处理 ===============
function triggerFileInput() {
    document.getElementById('file-input').click();
}

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        // 确保文件确实被选择了才上传
        console.log('文件已选择:', file.name);
        uploadFile(file);
        // 清空文件输入，避免重复触发
        event.target.value = '';
    }
}

function handleDragOver(event) {
    event.preventDefault();
    event.currentTarget.classList.add('dragover');
}

function handleDragLeave(event) {
    event.currentTarget.classList.remove('dragover');
}

function handleFileDrop(event) {
    event.preventDefault();
    event.currentTarget.classList.remove('dragover');
    
    const files = event.dataTransfer.files;
    if (files.length > 0) {
        uploadFile(files[0]);
    }
}

async function uploadFile(file) {
    // 验证文件类型
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        showNotification('请选择Excel文件(.xlsx或.xls)', 'error');
        return;
    }
    
    showLoading('正在上传和分析Excel文件...');
    
    try {
        const formData = new FormData();
        formData.append('file', file);
        
        const response = await fetch('/api/upload-excel', {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success) {
            currentSessionId = result.session_id;
            document.getElementById('file-name').textContent = result.filename;
            
            // 切换到工作区域
            document.getElementById('welcome-section').style.display = 'none';
            document.getElementById('workspace-section').style.display = 'block';
            
            // 添加系统消息
            addMessage('system', result.message);
            
            showNotification('Excel文件上传成功！', 'success');
        } else {
            throw new Error(result.detail || '上传失败');
        }
    } catch (error) {
        console.error('上传失败:', error);
        showNotification('上传失败: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

function uploadNewFile() {
    // 清除当前会话
    if (currentSessionId) {
        clearSession();
    }
    
    // 返回欢迎页面
    document.getElementById('workspace-section').style.display = 'none';
    document.getElementById('welcome-section').style.display = 'block';
    
    // 重置文件输入
    document.getElementById('file-input').value = '';
}

// =============== 聊天功能 ===============
async function sendMessage() {
    const input = document.getElementById('chat-input');
    const message = input.value.trim();
    
    if (!message) return;
    if (!currentSessionId) {
        showNotification('请先上传Excel文件', 'error');
        return;
    }
    
    // 清空输入框
    input.value = '';
    
    // 添加用户消息
    addMessage('user', message);
    
    // 禁用发送按钮
    const sendBtn = document.getElementById('send-btn');
    sendBtn.disabled = true;
    sendBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 处理中...';
    
    try {
        const response = await fetch('/api/chat', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                message: message,
                session_id: currentSessionId
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            addMessage('assistant', result.response);
        } else {
            throw new Error(result.detail || '处理失败');
        }
    } catch (error) {
        console.error('发送消息失败:', error);
        addMessage('system', '❌ 消息处理失败: ' + error.message);
    } finally {
        // 恢复发送按钮
        sendBtn.disabled = false;
        sendBtn.innerHTML = '<i class="fas fa-paper-plane"></i> 发送';
    }
}

function addMessage(role, content) {
    const messagesContainer = document.getElementById('chat-messages');
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${role}`;
    
    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';
    
    // 处理内容：检查是否包含HTML表格
    if (typeof content === 'string') {
        // 检查是否包含表格或特殊HTML内容
        const hasTable = content.includes('<table class="data-table">') || 
                         content.includes('<div class="sql-results-container">') ||
                         content.includes('<div class="error-message">') ||
                         content.includes('<div class="table-container">');
        
        if (hasTable) {
            // 直接设置HTML内容（因为这是我们生成的安全HTML）
            contentDiv.innerHTML = content;
            // 添加表格样式类（兼容不支持:has()的浏览器）
            contentDiv.classList.add('contains-table');
        } else {
            // 普通文本内容：转义HTML并处理换行
            const escapedContent = content
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;')
                .replace(/\n/g, '<br>');
            
            contentDiv.innerHTML = escapedContent;
        }
    } else {
        contentDiv.textContent = content;
    }
    
    const timeDiv = document.createElement('div');
    timeDiv.className = 'message-time';
    timeDiv.textContent = new Date().toLocaleTimeString();
    
    messageDiv.appendChild(contentDiv);
    messageDiv.appendChild(timeDiv);
    messagesContainer.appendChild(messageDiv);
    
    // 滚动到底部
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
}

function clearChat() {
    if (confirm('确定要清空聊天记录吗？')) {
        document.getElementById('chat-messages').innerHTML = '';
        if (currentSessionId) {
            // 这里可以调用API清空服务器端的会话历史
            fetch(`/api/session/${currentSessionId}`, {
                method: 'DELETE'
            }).catch(console.error);
        }
    }
}

// =============== PPT生成功能 ===============
async function generatePPT() {
    const input = document.getElementById('chat-input');
    const message = input.value.trim();
    
    if (!message) {
        showNotification('请输入PPT生成需求', 'error');
        return;
    }
    
    if (!currentSessionId) {
        showNotification('请先上传Excel文件', 'error');
        return;
    }
    
    // 清空输入框
    input.value = '';
    
    // 添加用户消息
    addMessage('user', `[PPT生成] ${message}`);
    
    // 显示进度卡片
    showPPTProgress();
    
    try {
        const response = await fetch('/api/generate-ppt', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                message: message,
                session_id: currentSessionId
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            currentTaskId = result.task_id;
            addMessage('system', '✅ PPT生成任务已启动，正在处理...');
            
            // 开始轮询任务状态
            pollTaskStatus();
        } else {
            throw new Error(result.detail || 'PPT生成启动失败');
        }
    } catch (error) {
        console.error('PPT生成失败:', error);
        addMessage('system', '❌ PPT生成失败: ' + error.message);
        hidePPTProgress();
    }
}

function showPPTProgress() {
    const progressCard = document.getElementById('ppt-progress');
    progressCard.style.display = 'block';
    updateProgress(0, '正在启动PPT生成...');
}

function hidePPTProgress() {
    const progressCard = document.getElementById('ppt-progress');
    progressCard.style.display = 'none';
    currentTaskId = null;
}

function updateProgress(percent, message) {
    document.getElementById('progress-fill').style.width = percent + '%';
    document.getElementById('progress-text').textContent = message;
}

async function pollTaskStatus() {
    if (!currentTaskId) return;
    
    try {
        const response = await fetch(`/api/task-status/${currentTaskId}`);
        const status = await response.json();
        
        updateProgress(status.progress, status.message);
        
        if (status.status === 'completed') {
            addMessage('system', '✅ ' + status.message);
            hidePPTProgress();
            refreshOutputFiles();
        } else if (status.status === 'failed') {
            addMessage('system', '❌ ' + status.message);
            hidePPTProgress();
        } else {
            // 继续轮询
            setTimeout(pollTaskStatus, 2000);
        }
    } catch (error) {
        console.error('获取任务状态失败:', error);
        addMessage('system', '❌ 获取任务状态失败');
        hidePPTProgress();
    }
}

// =============== 快速操作 ===============
function quickAnalysis(action) {
    if (!currentSessionId) {
        showNotification('请先上传Excel文件', 'error');
        return;
    }
    
    document.getElementById('chat-input').value = action;
    sendMessage();
}

// =============== 输出文件管理 ===============
async function refreshOutputFiles() {
    try {
        const response = await fetch('/api/output-files');
        const result = await response.json();
        
        const outputList = document.getElementById('output-list');
        outputList.innerHTML = '';
        
        if (result.files && result.files.length > 0) {
            result.files.forEach(file => {
                const fileItem = createFileItem(file);
                outputList.appendChild(fileItem);
            });
        } else {
            outputList.innerHTML = '<p style="text-align: center; color: var(--text-secondary); padding: 2rem;">暂无生成的文件</p>';
        }
    } catch (error) {
        console.error('获取文件列表失败:', error);
    }
}

function createFileItem(file) {
    const item = document.createElement('div');
    item.className = 'output-item';
    
    const infoDiv = document.createElement('div');
    infoDiv.className = 'file-info-text';
    
    const nameDiv = document.createElement('div');
    nameDiv.className = 'file-name';
    nameDiv.textContent = file.filename;
    
    const metaDiv = document.createElement('div');
    metaDiv.className = 'file-meta';
    metaDiv.textContent = `${formatFileSize(file.size)} • ${formatDate(file.created_time)}`;
    
    infoDiv.appendChild(nameDiv);
    infoDiv.appendChild(metaDiv);
    
    const downloadBtn = document.createElement('button');
    downloadBtn.className = 'download-btn';
    downloadBtn.innerHTML = '<i class="fas fa-download"></i>';
    downloadBtn.onclick = () => downloadFile(file.filename);
    
    item.appendChild(infoDiv);
    item.appendChild(downloadBtn);
    
    return item;
}

function downloadFile(filename) {
    window.open(`/api/download/${filename}`, '_blank');
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function formatDate(dateString) {
    const date = new Date(dateString);
    return date.toLocaleString('zh-CN');
}

// =============== 设置管理 ===============
function openSettings() {
    document.getElementById('settings-modal').style.display = 'block';
}

function closeSettings() {
    document.getElementById('settings-modal').style.display = 'none';
}

async function loadSettings() {
    try {
        const response = await fetch('/api/config');
        const config = await response.json();
        
        document.getElementById('api-key').value = config.api_key || '';
        document.getElementById('base-url').value = config.base_url || '';
        document.getElementById('model').value = config.model || 'gpt-4o-mini';
    } catch (error) {
        console.error('加载设置失败:', error);
    }
}

async function saveSettings() {
    const config = {
        api_key: document.getElementById('api-key').value,
        base_url: document.getElementById('base-url').value,
        model: document.getElementById('model').value
    };
    
    try {
        const response = await fetch('/api/config', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(config)
        });
        
        const result = await response.json();
        
        if (result.success) {
            showNotification('设置保存成功', 'success');
            closeSettings();
        } else {
            throw new Error(result.detail || '保存失败');
        }
    } catch (error) {
        console.error('保存设置失败:', error);
        showNotification('保存设置失败: ' + error.message, 'error');
    }
}

// =============== 工具函数 ===============
function showLoading(message = '处理中...') {
    const loading = document.getElementById('loading');
    loading.querySelector('p').textContent = message;
    loading.style.display = 'flex';
}

function hideLoading() {
    document.getElementById('loading').style.display = 'none';
}

function showNotification(message, type = 'info') {
    // 创建通知元素
    const notification = document.createElement('div');
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 12px 20px;
        border-radius: 8px;
        color: white;
        font-weight: 500;
        z-index: 10000;
        animation: slideIn 0.3s ease;
        max-width: 300px;
        word-wrap: break-word;
    `;
    
    // 根据类型设置颜色
    switch (type) {
        case 'success':
            notification.style.background = '#10b981';
            break;
        case 'error':
            notification.style.background = '#ef4444';
            break;
        case 'warning':
            notification.style.background = '#f59e0b';
            break;
        default:
            notification.style.background = '#3b82f6';
    }
    
    notification.textContent = message;
    document.body.appendChild(notification);
    
    // 3秒后自动移除
    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => {
            document.body.removeChild(notification);
        }, 300);
    }, 3000);
}

function clearSession() {
    if (currentSessionId) {
        fetch(`/api/session/${currentSessionId}`, {
            method: 'DELETE'
        }).catch(console.error);
        currentSessionId = null;
    }
}

// =============== 表格复制功能 ===============
function copyTableData(tableId) {
    try {
        const table = document.getElementById(tableId);
        if (!table) {
            showNotification('表格未找到', 'error');
            return;
        }

        // 获取表格数据
        const headers = [];
        const rows = [];

        // 获取表头
        const headerCells = table.querySelectorAll('thead th');
        headerCells.forEach(cell => {
            headers.push(cell.textContent.trim());
        });

        // 获取数据行
        const dataRows = table.querySelectorAll('tbody tr');
        dataRows.forEach(row => {
            const rowData = [];
            const cells = row.querySelectorAll('td');
            cells.forEach(cell => {
                // 优先使用data-value属性（完整值），否则使用显示文本
                const value = cell.getAttribute('data-value') || cell.textContent.trim();
                rowData.push(value);
            });
            rows.push(rowData);
        });

        // 生成制表符分隔的文本（Excel友好格式）
        let copyText = headers.join('\t') + '\n';
        rows.forEach(row => {
            copyText += row.join('\t') + '\n';
        });

        // 复制到剪贴板
        navigator.clipboard.writeText(copyText).then(() => {
            showNotification(`已复制 ${rows.length} 行数据到剪贴板`, 'success');
            // 临时高亮复制按钮
            highlightCopyButton(tableId, 'copy-table-btn');
        }).catch(err => {
            console.error('复制失败:', err);
            // 回退方案：使用旧的复制方法
            fallbackCopyToClipboard(copyText, `已复制 ${rows.length} 行数据到剪贴板`);
        });

    } catch (error) {
        console.error('复制表格数据失败:', error);
        showNotification('复制失败', 'error');
    }
}

function copyTableAsCSV(tableId) {
    try {
        const table = document.getElementById(tableId);
        if (!table) {
            showNotification('表格未找到', 'error');
            return;
        }

        // 获取表格数据
        const headers = [];
        const rows = [];

        // 获取表头
        const headerCells = table.querySelectorAll('thead th');
        headerCells.forEach(cell => {
            headers.push(escapeCSVField(cell.textContent.trim()));
        });

        // 获取数据行
        const dataRows = table.querySelectorAll('tbody tr');
        dataRows.forEach(row => {
            const rowData = [];
            const cells = row.querySelectorAll('td');
            cells.forEach(cell => {
                // 优先使用data-value属性（完整值），否则使用显示文本
                const value = cell.getAttribute('data-value') || cell.textContent.trim();
                rowData.push(escapeCSVField(value));
            });
            rows.push(rowData);
        });

        // 生成CSV格式文本
        let csvText = headers.join(',') + '\n';
        rows.forEach(row => {
            csvText += row.join(',') + '\n';
        });

        // 复制到剪贴板
        navigator.clipboard.writeText(csvText).then(() => {
            showNotification(`已复制 ${rows.length} 行CSV数据到剪贴板`, 'success');
            // 临时高亮复制按钮
            highlightCopyButton(tableId, 'copy-csv-btn');
        }).catch(err => {
            console.error('复制CSV失败:', err);
            // 回退方案：使用旧的复制方法
            fallbackCopyToClipboard(csvText, `已复制 ${rows.length} 行CSV数据到剪贴板`);
        });

    } catch (error) {
        console.error('复制CSV数据失败:', error);
        showNotification('复制CSV失败', 'error');
    }
}

function escapeCSVField(field) {
    // CSV字段转义：如果包含逗号、引号或换行符，则用引号包围并转义内部引号
    const str = String(field);
    if (str.includes(',') || str.includes('"') || str.includes('\n') || str.includes('\r')) {
        return '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
}

function highlightCopyButton(tableId, buttonClass) {
    // 查找对应的复制按钮并添加高亮效果
    const tableContainer = document.getElementById(tableId).closest('.table-container');
    const button = tableContainer.querySelector(`.${buttonClass}`);
    if (button) {
        button.classList.add('copy-success');
        setTimeout(() => {
            button.classList.remove('copy-success');
        }, 1500);
    }
}

function fallbackCopyToClipboard(text, successMessage) {
    // 回退复制方案（适用于较老的浏览器）
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.position = 'fixed';
    textArea.style.left = '-999999px';
    textArea.style.top = '-999999px';
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    
    try {
        const successful = document.execCommand('copy');
        if (successful) {
            showNotification(successMessage, 'success');
        } else {
            showNotification('复制失败，请手动选择复制', 'error');
        }
    } catch (err) {
        console.error('回退复制方法也失败:', err);
        showNotification('复制失败，请手动选择复制', 'error');
    } finally {
        document.body.removeChild(textArea);
    }
}

// 添加CSS动画
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from {
            transform: translateX(100%);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
    
    @keyframes slideOut {
        from {
            transform: translateX(0);
            opacity: 1;
        }
        to {
            transform: translateX(100%);
            opacity: 0;
        }
    }
    
    @keyframes copySuccess {
        0% { 
            background: var(--success-color);
            transform: scale(1);
        }
        50% { 
            background: var(--success-color);
            transform: scale(1.05);
        }
        100% { 
            background: var(--success-color);
            transform: scale(1);
        }
    }
    
    .copy-success {
        animation: copySuccess 0.6s ease;
        color: white !important;
    }
`;
document.head.appendChild(style);
