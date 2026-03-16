/**
 * AI 报表生成器 - 前端脚本
 * 聊天式 Pandas 数据分析界面
 */

'use strict';

// ===== Agent 管理 =====
const agents = {
    list: [],
    currentId: null,
};

function createAgentState() {
    return {
        tables: [],           // 已加载的表格
        messages: [],         // 对话消息
        history: [],          // 历史查询
        currentCode: '',      // 当前生成的代码
    };
}

// ===== 全局状态 =====
const state = {
    ...createAgentState(),
    isLoading: false,     // 加载状态
    backendConnected: false,
    llmConnected: false,
};

// ===== DOM 元素缓存 =====
const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => document.querySelectorAll(selector);

const elements = {
    chatMessages: $('#chat-messages'),
    emptyState: $('#empty-state'),
    userInput: $('#user-input'),
    btnSend: $('#btn-send'),
    quickActions: $('#quick-actions'),
    tableList: $('#table-list'),
    statsCard: $('#stats-card'),
    historyList: $('#history-list'),
    loadingOverlay: $('#loading-overlay'),
    loadingText: $('#loading-text'),
    toastContainer: $('#toast-container'),
    // 状态栏
    statusBackend: $('#status-backend'),
    statusBackendText: $('#status-backend-text'),
    statusLlm: $('#status-llm'),
    statusLlmText: $('#status-llm-text'),
    footerTables: $('#footer-tables'),
    // 统计
    statTables: $('#stat-tables'),
    statRows: $('#stat-rows'),
    statCols: $('#stat-cols'),
};

// ===== 工具函数 =====
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// 尝试将数据解析为图表数据
function tryParseChartData(data) {
    // 放宽限制：2-30行都可以生成图表
    if (!Array.isArray(data) || data.length < 2 || data.length > 30) {
        return null;
    }
    
    const columns = Object.keys(data[0]);
    if (columns.length < 2) return null;
    
    // 找到第一个字符串列作为标签，第一个数值列作为值
    let labelCol = null;
    let valueCol = null;
    
    for (const col of columns) {
        const sample = data[0][col];
        if (labelCol === null && (typeof sample === 'string' || col.toLowerCase().includes('名') || col.toLowerCase().includes('name'))) {
            labelCol = col;
        }
        if (valueCol === null && typeof sample === 'number') {
            valueCol = col;
        }
    }
    
    // 如果没找到明确的标签列，用第一列；没找到数值列，用第二列
    if (!labelCol) labelCol = columns[0];
    if (!valueCol) {
        for (const col of columns) {
            if (col !== labelCol) {
                // 检查是否有数值
                const hasNumber = data.some(row => typeof row[col] === 'number');
                if (hasNumber) {
                    valueCol = col;
                    break;
                }
            }
        }
    }
    
    if (!valueCol) return null;
    
    // 提取数据
    const labels = data.map(row => String(row[labelCol] || ''));
    const values = data.map(row => {
        const v = row[valueCol];
        return typeof v === 'number' ? v : parseFloat(v) || 0;
    });
    
    // 检查是否有有效数值
    if (values.every(v => v === 0)) return null;
    
    return { labels, values, labelCol, valueCol };
}

// 渲染条形图
function renderCharts(chartData, msgId, preferredType = 'bar') {
    const { labels, values, labelCol, valueCol } = chartData;
    const defaultType = preferredType || 'bar';  // 用户偏好的图表类型
    const maxValue = Math.max(...values);
    const total = values.reduce((a, b) => a + b, 0);
    
    // 格式化数值
    const formatValue = (v) => {
        if (v >= 1000000) return (v / 1000000).toFixed(1) + 'M';
        if (v >= 1000) return (v / 1000).toFixed(1) + 'K';
        if (v < 1 && v > 0) return v.toFixed(2);
        return v.toFixed(0);
    };
    
    // 颜色
    const colors = ['#2d5af0', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899', '#84cc16', '#14b8a6', '#f97316'];
    
    // 柱状图
    const barsHtml = labels.slice(0, 15).map((label, i) => {
        const percent = maxValue > 0 ? (values[i] / maxValue * 100) : 0;
        const color = colors[i % colors.length];
        return `
            <div class="chart-bar-row">
                <div class="chart-label" title="${escapeHtml(label)}">${escapeHtml(label.length > 10 ? label.slice(0, 10) + '...' : label)}</div>
                <div class="chart-bar-container">
                    <div class="chart-bar" style="width: ${percent}%; background: ${color};"></div>
                </div>
                <div class="chart-value">${formatValue(values[i])}</div>
            </div>
        `;
    }).join('');
    
    // 饼图（CSS conic-gradient）
    let pieGradient = '';
    let currentAngle = 0;
    const pieData = labels.slice(0, 8).map((label, i) => {
        const percent = total > 0 ? (values[i] / total * 100) : 0;
        const startAngle = currentAngle;
        currentAngle += percent * 3.6; // 360度 / 100%
        return { label, value: values[i], percent, color: colors[i % colors.length], startAngle, endAngle: currentAngle };
    });
    
    // 如果还有其他数据，归为"其他"
    if (labels.length > 8) {
        const otherValue = values.slice(8).reduce((a, b) => a + b, 0);
        const otherPercent = total > 0 ? (otherValue / total * 100) : 0;
        pieData.push({ label: '其他', value: otherValue, percent: otherPercent, color: '#6b7280', startAngle: currentAngle, endAngle: currentAngle + otherPercent * 3.6 });
    }
    
    // 构建饼图渐变
    const gradientParts = pieData.map((d, i) => {
        const start = i === 0 ? 0 : pieData.slice(0, i).reduce((sum, p) => sum + p.percent, 0);
        const end = start + d.percent;
        return `${d.color} ${start}% ${end}%`;
    });
    pieGradient = gradientParts.join(', ');
    
    // 饼图图例
    const legendHtml = pieData.map(d => `
        <div class="pie-legend-item">
            <span class="legend-color" style="background: ${d.color}"></span>
            <span class="legend-label">${escapeHtml(d.label.length > 8 ? d.label.slice(0, 8) + '..' : d.label)}</span>
            <span class="legend-value">${d.percent.toFixed(1)}%</span>
        </div>
    `).join('');
    
    // 根据用户偏好决定默认显示哪个图表
    const barActive = defaultType === 'bar' ? 'active' : '';
    const pieActive = defaultType === 'pie' ? 'active' : '';
    
    return `
        <div class="charts-container" id="chart-${msgId}">
            <div class="chart-tabs">
                <button class="chart-tab ${barActive}" onclick="switchChartTab('${msgId}', 'bar')">📊 柱状图</button>
                <button class="chart-tab ${pieActive}" onclick="switchChartTab('${msgId}', 'pie')">🥧 饼图</button>
            </div>
            <div class="chart-panel bar-chart ${barActive}" data-type="bar">
                <div class="chart-title">${escapeHtml(valueCol || '数值')} 分布${labels.length > 15 ? '（前15项）' : ''}</div>
                <div class="chart-bars">
                    ${barsHtml}
                </div>
            </div>
            <div class="chart-panel pie-chart ${pieActive}" data-type="pie">
                <div class="chart-title">${escapeHtml(valueCol || '数值')} 占比</div>
                <div class="pie-container">
                    <div class="pie-chart-svg" style="background: conic-gradient(${pieGradient});"></div>
                    <div class="pie-legend">
                        ${legendHtml}
                    </div>
                </div>
            </div>
        </div>
    `;
}

// 兼容旧函数名
function renderBarChart(chartData, msgId) {
    return renderCharts(chartData, msgId);
}

// 切换图表类型
window.switchChartTab = function(msgId, type) {
    const container = document.getElementById(`chart-${msgId}`);
    if (!container) return;
    
    // 切换 tab
    container.querySelectorAll('.chart-tab').forEach(tab => {
        tab.classList.toggle('active', tab.textContent.includes(type === 'bar' ? '柱状' : '饼'));
    });
    
    // 切换面板
    container.querySelectorAll('.chart-panel').forEach(panel => {
        panel.classList.toggle('active', panel.dataset.type === type);
    });
};

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function formatTime(date = new Date()) {
    // 处理时间戳数字或字符串
    if (typeof date === 'number' || typeof date === 'string') {
        date = new Date(date);
    }
    // 检查是否是有效的 Date 对象
    if (!(date instanceof Date) || isNaN(date.getTime())) {
        return '';
    }
    return date.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' });
}

function showLoading(text = 'AI 正在分析...') {
    state.isLoading = true;
    elements.loadingText.textContent = text;
    elements.loadingOverlay.style.display = 'flex';
}

function hideLoading() {
    state.isLoading = false;
    elements.loadingOverlay.style.display = 'none';
}

function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    elements.toastContainer.appendChild(toast);
    
    setTimeout(() => {
        toast.style.animation = 'toastIn 0.3s ease reverse';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// ===== Agent 管理 =====
function initAgents() {
    // 从 localStorage 恢复或创建默认 Agent
    const saved = localStorage.getItem('agents');
    if (saved) {
        try {
            const data = JSON.parse(saved);
            agents.list = data.list || [];
            agents.currentId = data.currentId;
        } catch (e) {
            console.error('Failed to load agents:', e);
        }
    }
    
    if (agents.list.length === 0) {
        // 创建默认 Agent
        const defaultAgent = {
            id: generateId(),
            name: '默认 Agent',
            createdAt: Date.now(),
            lastUsed: Date.now(),
            state: createAgentState()
        };
        agents.list.push(defaultAgent);
        agents.currentId = defaultAgent.id;
    }
    
    // 恢复当前 Agent 的状态
    loadAgentState(agents.currentId);
    updateAgentUI();
    
    // 绑定搜索事件
    const searchInput = document.getElementById('agent-search');
    if (searchInput) {
        searchInput.addEventListener('input', filterAgents);
    }
}

function saveAgents() {
    // 先保存当前 Agent 的状态
    const currentAgent = agents.list.find(a => a.id === agents.currentId);
    if (currentAgent) {
        currentAgent.lastUsed = Date.now();
        currentAgent.state = {
            tables: state.tables,
            messages: state.messages,
            history: state.history,
            currentCode: state.currentCode,
        };
    }
    
    localStorage.setItem('agents', JSON.stringify({
        list: agents.list,
        currentId: agents.currentId
    }));
}

function loadAgentState(agentId) {
    const agent = agents.list.find(a => a.id === agentId);
    if (agent && agent.state) {
        state.tables = agent.state.tables || [];
        state.messages = agent.state.messages || [];
        state.history = agent.state.history || [];
        state.currentCode = agent.state.currentCode || '';
    }
}

// 格式化相对时间
function formatRelativeTime(timestamp) {
    if (!timestamp) return '';
    const now = Date.now();
    const diff = now - timestamp;
    const minute = 60 * 1000;
    const hour = 60 * minute;
    const day = 24 * hour;
    
    if (diff < minute) return '刚刚';
    if (diff < hour) return `${Math.floor(diff / minute)}分钟前`;
    if (diff < day) return `${Math.floor(diff / hour)}小时前`;
    if (diff < 7 * day) return `${Math.floor(diff / day)}天前`;
    return new Date(timestamp).toLocaleDateString('zh-CN');
}

function updateAgentUI(filter = '') {
    const currentAgent = agents.list.find(a => a.id === agents.currentId);
    const nameEl = document.getElementById('current-agent-name');
    if (nameEl && currentAgent) {
        nameEl.textContent = currentAgent.name;
    }
    
    // 按最后使用时间排序
    const sortedAgents = [...agents.list].sort((a, b) => (b.lastUsed || 0) - (a.lastUsed || 0));
    
    // 过滤
    const filteredAgents = filter 
        ? sortedAgents.filter(a => a.name.toLowerCase().includes(filter.toLowerCase()))
        : sortedAgents;
    
    // 更新下拉列表
    const listEl = document.getElementById('agent-list');
    if (listEl) {
        if (filteredAgents.length === 0) {
            listEl.innerHTML = `<div class="agent-empty">没有找到匹配的 Agent</div>`;
        } else {
            listEl.innerHTML = filteredAgents.map(agent => `
                <div class="agent-menu-item ${agent.id === agents.currentId ? 'active' : ''}" 
                     onclick="switchAgent('${agent.id}')">
                    <span class="agent-item-icon">${agent.id === agents.currentId ? '✓' : '🤖'}</span>
                    <div class="agent-item-info">
                        <div class="agent-item-name">${escapeHtml(agent.name)}</div>
                        <div class="agent-item-time">${formatRelativeTime(agent.lastUsed || agent.createdAt)}</div>
                    </div>
                    <div class="agent-actions">
                        <span class="agent-rename" onclick="renameAgent('${agent.id}', event)" title="重命名">✏️</span>
                        ${agents.list.length > 1 ? `<span class="agent-delete" onclick="deleteAgent('${agent.id}', event)" title="删除">×</span>` : ''}
                    </div>
                </div>
            `).join('');
        }
    }
}

function filterAgents() {
    const searchInput = document.getElementById('agent-search');
    const filter = searchInput ? searchInput.value.trim() : '';
    updateAgentUI(filter);
}

// ===== 自定义 Prompt 对话框 =====
let promptResolve = null;

function showPromptModal(title, defaultValue = '') {
    return new Promise((resolve) => {
        promptResolve = resolve;
        const modal = document.getElementById('prompt-modal');
        const titleEl = document.getElementById('prompt-title');
        const inputEl = document.getElementById('prompt-input');
        
        if (titleEl) titleEl.textContent = title;
        if (inputEl) {
            inputEl.value = defaultValue;
            setTimeout(() => {
                inputEl.focus();
                inputEl.select();
            }, 50);
        }
        if (modal) modal.style.display = 'flex';
        
        // ESC 关闭
        const escHandler = (e) => {
            if (e.key === 'Escape') {
                closePromptModal();
                document.removeEventListener('keydown', escHandler);
            } else if (e.key === 'Enter') {
                confirmPromptModal();
                document.removeEventListener('keydown', escHandler);
            }
        };
        document.addEventListener('keydown', escHandler);
    });
}

window.closePromptModal = function() {
    const modal = document.getElementById('prompt-modal');
    if (modal) modal.style.display = 'none';
    if (promptResolve) {
        promptResolve(null);
        promptResolve = null;
    }
};

window.confirmPromptModal = function() {
    const inputEl = document.getElementById('prompt-input');
    const value = inputEl ? inputEl.value.trim() : '';
    const modal = document.getElementById('prompt-modal');
    if (modal) modal.style.display = 'none';
    if (promptResolve) {
        promptResolve(value || null);
        promptResolve = null;
    }
};

// ===== 关闭对话框 =====
function showCloseDialog() {
    const dialog = document.getElementById('close-dialog');
    if (dialog) {
        dialog.style.display = 'flex';
    }
}

function hideCloseDialog() {
    const dialog = document.getElementById('close-dialog');
    if (dialog) {
        dialog.style.display = 'none';
    }
}

window.handleCloseChoice = function(choice) {
    hideCloseDialog();
    window.electronAPI.closeDialogResult(choice);
};

// 监听主进程发来的显示关闭对话框事件
if (window.electronAPI?.onShowCloseDialog) {
    window.electronAPI.onShowCloseDialog(() => {
        showCloseDialog();
    });
}

window.createNewAgent = async function() {
    // 保存当前 Agent 状态
    saveAgents();
    
    // 清空后端表格（新会话从空开始）
    try {
        await apiRequest('DELETE', '/api/tables');
    } catch (e) {
        console.error('清空表格失败:', e);
    }
    
    // 创建新 Agent（临时名称，后续根据对话自动生成）
    const newAgent = {
        id: generateId(),
        name: '新对话',
        createdAt: Date.now(),
        lastUsed: Date.now(),
        autoNamed: false,  // 标记是否已自动命名
        state: createAgentState()
    };
    
    agents.list.push(newAgent);
    agents.currentId = newAgent.id;
    
    // 清空当前状态
    Object.assign(state, createAgentState());
    
    // 通知后端创建新工作区
    try {
        await apiRequest('POST', '/api/workspace/create', { name: newAgent.name });
    } catch (e) {
        console.error('Failed to create workspace:', e);
    }
    
    saveAgents();
    updateAgentUI();
    updateTablesUI();
    renderAllMessages();
    
    showToast('已创建新对话', 'success');
};

// 根据用户消息自动生成 Agent 名称
async function autoNameAgent(userMessage) {
    const currentAgent = agents.list.find(a => a.id === agents.currentId);
    if (!currentAgent || currentAgent.autoNamed) return;
    
    // 截取前100个字符用于生成名称
    const content = userMessage.slice(0, 100);
    
    try {
        const result = await apiRequest('POST', '/api/generate-title', { 
            content: content,
            max_length: 8
        });
        
        if (result && result.title) {
            currentAgent.name = result.title;
            currentAgent.autoNamed = true;
            saveAgents();
            updateAgentUI();
        }
    } catch (e) {
        // 如果API失败，用消息前8个字作为名称
        currentAgent.name = content.slice(0, 8) + (content.length > 8 ? '...' : '');
        currentAgent.autoNamed = true;
        saveAgents();
        updateAgentUI();
    }
}

// 重命名 Agent
window.renameAgent = async function(agentId, event) {
    if (event) event.stopPropagation();
    
    const agent = agents.list.find(a => a.id === agentId);
    if (!agent) return;
    
    const newName = await showPromptModal('重命名', agent.name);
    if (!newName || newName === agent.name) return;
    
    agent.name = newName;
    agent.autoNamed = true;  // 用户手动命名后不再自动更新
    saveAgents();
    updateAgentUI();
    
    showToast('已重命名', 'success');
};

window.switchAgent = async function(agentId) {
    if (agentId === agents.currentId) {
        return;
    }
    
    // 保存当前 Agent 状态
    saveAgents();
    
    // 清空后端表格（每个会话独立）
    try {
        await apiRequest('DELETE', '/api/tables');
    } catch (e) {
        console.error('清空表格失败:', e);
    }
    
    // 切换
    agents.currentId = agentId;
    loadAgentState(agentId);
    
    // 如果新 Agent 有表格记录，需要提示用户重新上传
    if (state.tables.length > 0) {
        // 清空前端表格状态（因为后端已清空）
        state.tables = [];
        showToast('请重新上传数据表', 'info');
    }
    
    saveAgents();
    updateAgentUI();
    updateTablesUI();
    renderAllMessages();
    
    const agent = agents.list.find(a => a.id === agentId);
    showToast(`已切换到: ${agent?.name || 'Agent'}`, 'info');
};

window.deleteAgent = async function(agentId, event) {
    if (event) event.stopPropagation();
    
    if (agents.list.length <= 1) {
        showToast('至少保留一个 Agent', 'warning');
        return;
    }
    
    const agent = agents.list.find(a => a.id === agentId);
    
    // 使用自定义确认（直接删除，因为有撤销能力或可重建）
    agents.list = agents.list.filter(a => a.id !== agentId);
    
    // 如果删除的是当前 Agent，切换到第一个
    if (agentId === agents.currentId) {
        agents.currentId = agents.list[0].id;
        loadAgentState(agents.currentId);
    }
    
    saveAgents();
    updateAgentUI();
    updateTablesUI();
    renderAllMessages();
    
    showToast(`已删除: ${agent?.name || 'Agent'}`, 'info');
};

function renderAllMessages() {
    // 清空消息区域
    const msgContainer = elements.chatMessages;
    msgContainer.innerHTML = '';
    
    // 重新添加空状态
    const emptyState = document.createElement('div');
    emptyState.className = 'empty-state';
    emptyState.id = 'empty-state';
    emptyState.style.display = state.messages.length > 0 ? 'none' : 'flex';
    emptyState.innerHTML = `
        <div class="empty-icon">🤖</div>
        <h3>开始对话</h3>
        <p>上传数据表后，用自然语言描述你的需求</p>
    `;
    msgContainer.appendChild(emptyState);
    
    // 渲染所有消息
    state.messages.forEach(msg => renderMessage(msg));
}

// ===== API 调用 =====
async function apiRequest(method, endpoint, data = null, isFormData = false) {
    try {
        const result = await window.electronAPI.apiRequest(method, endpoint, data, isFormData);
        return result;
    } catch (error) {
        console.error('API Error:', error);
        return { success: false, error: error.message };
    }
}

// ===== 状态检查 =====
async function checkStatus(retryCount = 0) {
    const maxRetries = 15; // 增加重试次数，最多等待30秒
    try {
        const result = await apiRequest('GET', '/api/health');
        state.backendConnected = result.success && result.data?.status === 'ok';
        
        elements.statusBackend.className = 'status-dot ' + (state.backendConnected ? 'online' : 'offline');
        elements.statusBackendText.textContent = state.backendConnected ? '已连接' : '未连接';
        
        if (state.backendConnected) {
            state.llmConnected = true;
            elements.statusLlm.className = 'status-dot online';
            elements.statusLlmText.textContent = '已连接';
        } else if (retryCount < maxRetries) {
            // 后端未就绪，2秒后重试
            elements.statusBackendText.textContent = `连接中(${retryCount + 1}/${maxRetries})`;
            setTimeout(() => checkStatus(retryCount + 1), 2000);
        } else {
            // 达到最大重试次数，显示错误提示
            showConnectionError();
        }
    } catch (e) {
        state.backendConnected = false;
        elements.statusBackend.className = 'status-dot offline';
        elements.statusBackendText.textContent = retryCount < maxRetries ? `连接中(${retryCount + 1}/${maxRetries})` : '连接失败';
        
        // 重试
        if (retryCount < maxRetries) {
            setTimeout(() => checkStatus(retryCount + 1), 2000);
        } else {
            showConnectionError();
        }
    }
}

// 显示连接错误提示
function showConnectionError() {
    const errorMsg = `
        <div class="connection-error-message">
            <h4>⚠️ 后端服务连接失败</h4>
            <p>可能的原因：</p>
            <ul>
                <li>端口 5000 被其他程序占用</li>
                <li>防火墙阻止了本地连接</li>
                <li>后端服务启动失败</li>
            </ul>
            <p>解决方案：</p>
            <ul>
                <li>关闭占用端口的程序后重启应用</li>
                <li>检查防火墙设置允许本地连接</li>
                <li>尝试以管理员身份运行</li>
            </ul>
            <button class="btn btn-primary" onclick="location.reload()">重新连接</button>
        </div>
    `;
    addMessage('system', errorMsg, 'error');
}

// ===== 主题切换 =====
function initTheme() {
    // 固定使用深色主题
    document.documentElement.setAttribute('data-theme', 'dark');
}

// ===== 侧边栏收起 =====
let sidebarCollapsed = localStorage.getItem('sidebarCollapsed') === 'true';

function toggleSidebar() {
    sidebarCollapsed = !sidebarCollapsed;
    document.body.classList.toggle('sidebar-collapsed', sidebarCollapsed);
    localStorage.setItem('sidebarCollapsed', sidebarCollapsed);
}

// ===== 设置管理 =====
const DEFAULT_SHORTCUTS = [
    { id: 'send', name: '发送消息', icon: '📤', keys: ['Ctrl', 'Enter'], action: 'sendMessage' },
    { id: 'upload', name: '上传文件', icon: '📁', keys: ['Ctrl', 'U'], action: 'uploadFile' },
    { id: 'export', name: '导出数据', icon: '💾', keys: ['Ctrl', 'E'], action: 'exportData' },
    { id: 'newAgent', name: '新建对话', icon: '➕', keys: ['Ctrl', 'N'], action: 'createNewAgent' },
    { id: 'toggleSidebar', name: '收起侧栏', icon: '📋', keys: ['Ctrl', 'B'], action: 'toggleSidebar' },
];

let settings = {
    shortcutsEnabled: false,
    shortcuts: JSON.parse(JSON.stringify(DEFAULT_SHORTCUTS)),
    backendPort: 5000,
};

function loadSettings() {
    const saved = localStorage.getItem('appSettings');
    if (saved) {
        try {
            const parsed = JSON.parse(saved);
            settings = {
                shortcutsEnabled: parsed.shortcutsEnabled ?? false,
                shortcuts: parsed.shortcuts || JSON.parse(JSON.stringify(DEFAULT_SHORTCUTS)),
                backendPort: parsed.backendPort || 5000,
            };
        } catch (e) {
            console.error('Failed to load settings:', e);
        }
    }
}

function saveSettings() {
    localStorage.setItem('appSettings', JSON.stringify(settings));
}

function openSettings() {
    const modal = document.getElementById('settings-modal');
    modal.style.display = 'flex';
    renderShortcutsList();
    document.getElementById('shortcuts-enabled').checked = settings.shortcutsEnabled;
    document.getElementById('backend-port').value = settings.backendPort || 5000;
}

// 应用端口设置并重启
window.applyPortAndRestart = async function() {
    const portInput = document.getElementById('backend-port');
    const port = parseInt(portInput.value, 10);
    
    if (isNaN(port) || port < 1024 || port > 65535) {
        showToast('端口号必须在 1024-65535 之间', 'error');
        return;
    }
    
    if (port === settings.backendPort) {
        showToast('端口未更改', 'info');
        return;
    }
    
    // 确认重启
    if (!confirm(`确定要将端口修改为 ${port} 并重启应用吗？`)) {
        return;
    }
    
    settings.backendPort = port;
    saveSettings();
    
    // 通知主进程保存端口配置并重启
    if (window.electronAPI && window.electronAPI.setBackendPort) {
        await window.electronAPI.setBackendPort(port);
    }
    
    showToast(`端口已设置为 ${port}，正在重启...`, 'success');
    
    // 触发应用重启
    if (window.electronAPI && window.electronAPI.restartApp) {
        setTimeout(() => {
            window.electronAPI.restartApp();
        }, 1000);
    }
}

// 诊断端口问题
window.runPortDiagnostics = async function() {
    const resultDiv = document.getElementById('diagnostics-result');
    resultDiv.style.display = 'block';
    resultDiv.innerHTML = '<span class="diag-info">正在诊断...</span>';
    
    if (window.electronAPI && window.electronAPI.diagnosePort) {
        const result = await window.electronAPI.diagnosePort();
        
        let html = '';
        html += `<span class="diag-info">═══ 端口诊断报告 ═══</span>\n\n`;
        html += `<span class="diag-info">📁 配置文件:</span> ${result.configPath}\n`;
        html += `<span class="diag-info">⚙️ 配置的端口:</span> ${result.configuredPort}\n`;
        html += `<span class="diag-info">🔌 当前使用端口:</span> ${result.currentPort}\n\n`;
        
        html += `<span class="diag-info">═══ 端口占用检查 ═══</span>\n`;
        for (const check of result.portChecks) {
            const icon = check.inUse ? '❌' : '✅';
            const cls = check.inUse ? 'diag-error' : 'diag-ok';
            html += `<span class="${cls}">${icon} 端口 ${check.port}: ${check.inUse ? '被占用' : '可用'}</span>`;
            if (check.process) {
                html += ` <span class="diag-warn">(${check.process})</span>`;
            }
            html += '\n';
        }
        
        html += `\n<span class="diag-info">═══ 后台进程 ═══</span>\n`;
        if (result.backendProcesses.length > 0) {
            for (const proc of result.backendProcesses) {
                html += `<span class="diag-warn">⚠️ ${proc}</span>\n`;
            }
        } else {
            html += `<span class="diag-ok">✅ 没有发现 report-backend.exe 进程</span>\n`;
        }
        
        html += `\n<span class="diag-info">═══ 建议 ═══</span>\n`;
        html += `<span class="diag-info">${result.suggestion}</span>`;
        
        resultDiv.innerHTML = html;
    } else {
        resultDiv.innerHTML = '<span class="diag-error">诊断功能不可用</span>';
    }
}

// 强制清理后台进程
window.forceCleanupBackend = async function() {
    if (!confirm('确定要强制清理所有后台进程吗？\n\n这将杀死所有 report-backend.exe 进程。')) {
        return;
    }
    
    showToast('正在清理后台进程...', 'info');
    
    if (window.electronAPI && window.electronAPI.forceCleanup) {
        const result = await window.electronAPI.forceCleanup();
        if (result.success) {
            showToast(`清理完成！${result.message}`, 'success');
            
            // 询问是否重启
            if (confirm('后台进程已清理。是否立即重启应用？')) {
                window.electronAPI.restartApp();
            }
        } else {
            showToast(`清理失败: ${result.message}`, 'error');
        }
    } else {
        showToast('功能不可用', 'error');
    }
}

function closeSettings() {
    document.getElementById('settings-modal').style.display = 'none';
    saveSettings();
}

function toggleShortcuts() {
    settings.shortcutsEnabled = document.getElementById('shortcuts-enabled').checked;
    saveSettings();
    renderShortcutsList();
    showToast(settings.shortcutsEnabled ? '快捷键已启用' : '快捷键已禁用', 'success');
}

function renderShortcutsList() {
    const list = document.getElementById('shortcuts-list');
    list.innerHTML = settings.shortcuts.map(shortcut => `
        <div class="shortcut-item ${settings.shortcutsEnabled ? '' : 'disabled'}">
            <div class="shortcut-info">
                <div class="shortcut-icon">${shortcut.icon}</div>
                <div class="shortcut-name">${shortcut.name}</div>
            </div>
            <div class="shortcut-keys">
                <div class="shortcut-key-editor" 
                     data-shortcut-id="${shortcut.id}" 
                     onclick="startRecordingShortcut('${shortcut.id}')"
                     title="点击编辑快捷键">
                    ${shortcut.keys.map(key => `<span class="key-badge">${key}</span>`).join('<span class="key-separator">+</span>')}
                </div>
            </div>
        </div>
    `).join('');
}

let recordingShortcutId = null;

function startRecordingShortcut(id) {
    if (!settings.shortcutsEnabled) {
        showToast('请先启用快捷键功能', 'warning');
        return;
    }
    
    // 如果已经在录制，先取消
    if (recordingShortcutId) {
        stopRecordingShortcut();
    }
    
    recordingShortcutId = id;
    const editor = document.querySelector(`[data-shortcut-id="${id}"]`);
    if (editor) {
        editor.classList.add('recording');
        editor.innerHTML = '<span class="key-badge" style="background: rgba(245, 158, 11, 0.3);">按下快捷键...</span>';
    }
    
    // 添加键盘监听
    document.addEventListener('keydown', recordShortcut);
    document.addEventListener('click', handleRecordClick);
}

function handleRecordClick(e) {
    // 如果点击的不是编辑器，取消录制
    const editor = e.target.closest('.shortcut-key-editor');
    if (!editor || editor.dataset.shortcutId !== recordingShortcutId) {
        stopRecordingShortcut();
    }
}

function recordShortcut(e) {
    e.preventDefault();
    e.stopPropagation();
    
    if (!recordingShortcutId) return;
    
    // 忽略单独的修饰键
    if (['Control', 'Alt', 'Shift', 'Meta'].includes(e.key)) {
        return;
    }
    
    // 构建快捷键组合
    const keys = [];
    if (e.ctrlKey) keys.push('Ctrl');
    if (e.altKey) keys.push('Alt');
    if (e.shiftKey) keys.push('Shift');
    
    // 处理按键名称
    let keyName = e.key;
    if (keyName === ' ') keyName = 'Space';
    else if (keyName === 'ArrowUp') keyName = '↑';
    else if (keyName === 'ArrowDown') keyName = '↓';
    else if (keyName === 'ArrowLeft') keyName = '←';
    else if (keyName === 'ArrowRight') keyName = '→';
    else if (keyName === 'Escape') keyName = 'Esc';
    else keyName = keyName.toUpperCase();
    
    keys.push(keyName);
    
    // 更新快捷键
    const shortcut = settings.shortcuts.find(s => s.id === recordingShortcutId);
    if (shortcut) {
        shortcut.keys = keys;
        saveSettings();
        showToast(`快捷键已更新为 ${keys.join(' + ')}`, 'success');
    }
    
    stopRecordingShortcut();
}

function stopRecordingShortcut() {
    recordingShortcutId = null;
    document.removeEventListener('keydown', recordShortcut);
    document.removeEventListener('click', handleRecordClick);
    renderShortcutsList();
}

function resetSettings() {
    settings = {
        shortcutsEnabled: false,
        shortcuts: JSON.parse(JSON.stringify(DEFAULT_SHORTCUTS)),
    };
    saveSettings();
    document.getElementById('shortcuts-enabled').checked = false;
    renderShortcutsList();
    showToast('已恢复默认设置', 'success');
}

// 处理快捷键
function handleGlobalShortcuts(e) {
    if (!settings.shortcutsEnabled) return;
    if (recordingShortcutId) return; // 正在录制时不触发
    
    // 检查是否匹配任何快捷键
    for (const shortcut of settings.shortcuts) {
        const keys = shortcut.keys;
        
        // 检查修饰键
        const needsCtrl = keys.includes('Ctrl');
        const needsAlt = keys.includes('Alt');
        const needsShift = keys.includes('Shift');
        
        if (needsCtrl !== e.ctrlKey) continue;
        if (needsAlt !== e.altKey) continue;
        if (needsShift !== e.shiftKey) continue;
        
        // 检查主键
        const mainKey = keys.find(k => !['Ctrl', 'Alt', 'Shift'].includes(k));
        let pressedKey = e.key;
        if (pressedKey === ' ') pressedKey = 'Space';
        else if (pressedKey === 'ArrowUp') pressedKey = '↑';
        else if (pressedKey === 'ArrowDown') pressedKey = '↓';
        else if (pressedKey === 'ArrowLeft') pressedKey = '←';
        else if (pressedKey === 'ArrowRight') pressedKey = '→';
        else if (pressedKey === 'Escape') pressedKey = 'Esc';
        else pressedKey = pressedKey.toUpperCase();
        
        if (mainKey === pressedKey) {
            e.preventDefault();
            executeShortcutAction(shortcut.action);
            return;
        }
    }
}

function executeShortcutAction(action) {
    switch (action) {
        case 'sendMessage':
            sendMessage();
            break;
        case 'uploadFile':
            uploadFile();
            break;
        case 'exportData':
            const lastDataMsg = [...state.messages].reverse().find(m => m.data && Array.isArray(m.data) && m.data.length > 0);
            if (lastDataMsg) {
                window.exportData?.(lastDataMsg.id, 'bar');
            } else {
                showToast('没有可导出的数据', 'warning');
            }
            break;
        case 'createNewAgent':
            window.createNewAgent?.();
            break;
        case 'toggleSidebar':
            toggleSidebar();
            break;
    }
}

// 暴露到全局
window.openSettings = openSettings;
window.closeSettings = closeSettings;
window.toggleShortcuts = toggleShortcuts;
window.startRecordingShortcut = startRecordingShortcut;
window.resetSettings = resetSettings;

function initSidebar() {
    if (sidebarCollapsed) {
        document.body.classList.add('sidebar-collapsed');
    }
}

// ===== 工作区状态 =====

// ===== 表格管理 =====
function updateTablesUI() {
    const tableList = elements.tableList;
    
    if (state.tables.length === 0) {
        tableList.innerHTML = '<div class="empty-table-hint">暂无数据表，请上传文件</div>';
        elements.statsCard.style.display = 'none';
        elements.footerTables.textContent = '0';
        return;
    }
    
    tableList.innerHTML = state.tables.map(table => {
        const sheetsCount = table.sheets?.length || 1;
        const sheetsInfo = sheetsCount > 1 ? `📑${sheetsCount}` : '';
        const alias = table.alias || '';
        const nickname = table.nickname || '';
        const displayName = nickname || table.name;
        const truncatedName = displayName.length > 12 ? displayName.substring(0, 12) + '...' : displayName;
        
        return `
            <div class="table-item" data-table="${escapeHtml(table.name)}">
                <div class="table-main">
                    <span class="table-alias" title="引用时可使用: ${alias}表">${alias}</span>
                    <span class="table-name" title="${escapeHtml(table.name)}">${escapeHtml(truncatedName)}</span>
                    ${nickname ? '' : `<button class="nickname-btn" onclick="editNickname(event, '${escapeHtml(table.name)}')" title="设置昵称">✏️</button>`}
                </div>
                <div class="table-meta">
                    <span class="table-info">${table.row_count}行 ${table.columns.length}列 ${sheetsInfo}</span>
                    <button class="table-action-btn" onclick="showTableActions('${escapeHtml(table.name)}')" title="操作">⋮</button>
                </div>
            </div>
        `;
    }).join('');
    
    // 统计
    elements.statsCard.style.display = 'block';
    elements.statTables.textContent = state.tables.length;
    elements.statRows.textContent = state.tables.reduce((sum, t) => sum + t.row_count, 0);
    elements.statCols.textContent = state.tables.reduce((sum, t) => sum + t.columns.length, 0);
    elements.footerTables.textContent = state.tables.length;
    
    // 绑定点击事件
    tableList.querySelectorAll('.table-item').forEach(item => {
        item.addEventListener('click', (e) => {
            // 如果点击的是按钮，不触发
            if (e.target.tagName === 'BUTTON') return;
            
            const tableName = item.dataset.table;
            const table = state.tables.find(t => t.name === tableName);
            const alias = table?.alias || '';
            elements.userInput.value = `查看 ${alias}表 的数据`;
            elements.userInput.focus();
        });
    });
}

// 编辑表昵称
window.editNickname = async function(event, tableName) {
    if (event) event.stopPropagation();
    
    const table = state.tables.find(t => t.name === tableName);
    if (!table) return;
    
    // 使用现有的 prompt 对话框
    const result = await showPromptModal(
        `设置 ${table.alias}表 的昵称`, 
        table.nickname || ''
    );
    
    if (result === null) return;
    
    // 调用 API 更新昵称
    try {
        const apiResult = await apiRequest('POST', '/api/update-nickname', {
            table_name: tableName,
            nickname: result
        });
        
        if (apiResult.success) {
            table.nickname = result;
            updateTablesUI();
            showToast(`已设置昵称: ${result || '(已清除)'}`, 'success');
        } else {
            showToast('设置昵称失败', 'error');
        }
    } catch (e) {
        showToast('设置昵称失败: ' + e.message, 'error');
    }
};

async function uploadFile() {
    const file = await window.electronAPI.selectFile();
    if (!file) return;
    
    showLoading('上传文件中...');
    
    try {
        const result = await apiRequest('POST', '/api/upload', {
            file: { path: file.path, name: file.name }
        }, true);
        
        if (result.success && result.data) {
            const tableData = result.data.table || result.data.meta;
            const tableInfo = {
                name: result.data.table_name,
                alias: tableData.alias || result.data.table?.alias || '',
                nickname: tableData.nickname || '',
                columns: tableData.columns,
                row_count: tableData.row_count,
                sheets: tableData.sheets || result.data.meta?.sheets || ['Sheet1'],
                current_sheet: tableData.current_sheet || null,
                header_row: tableData.header_row || 1,
                file_path: tableData.file_path || result.data.meta?.file_path,
            };
            
            // 检查是否已存在
            const existIndex = state.tables.findIndex(t => t.name === tableInfo.name);
            if (existIndex >= 0) {
                state.tables[existIndex] = tableInfo;
            } else {
                state.tables.push(tableInfo);
            }
            
            updateTablesUI();
            showToast(`已加载: ${tableInfo.name}`, 'success');
            
            // 隐藏空状态
            if (elements.emptyState) {
                elements.emptyState.style.display = 'none';
            }
        } else {
            showToast(result.error || '上传失败', 'error');
        }
    } catch (error) {
        showToast('上传失败: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

function loadSampleData() {
    showLoading('加载示例数据...');
    
    setTimeout(() => {
        state.tables = [
            { 
                name: '销售数据', 
                columns: ['日期', '销售额', '数量', '区域', '产品', '客户'], 
                row_count: 500 
            },
            { 
                name: '产品信息', 
                columns: ['产品ID', '产品名称', '类别', '单价', '库存'], 
                row_count: 50 
            },
        ];
        
        updateTablesUI();
        hideLoading();
        showToast('已加载示例数据', 'success');
        
        if (elements.emptyState) {
            elements.emptyState.style.display = 'none';
        }
    }, 800);
}

function clearChat() {
    state.messages = [];
    elements.chatMessages.innerHTML = '';
    
    // 显示空状态
    elements.chatMessages.innerHTML = `
        <div class="empty-state" id="empty-state">
            <div class="empty-icon">🤖</div>
            <h3>开始对话</h3>
            <p>上传数据表后，用自然语言描述你的需求</p>
            <div class="example-questions">
                <span class="example-tag" data-question="按区域统计销售额总和，降序排列">📊 按区域统计销售额</span>
                <span class="example-tag" data-question="找出销售额最高的前10个产品">🏆 TOP10 产品</span>
                <span class="example-tag" data-question="计算每月销售趋势">📈 月度趋势分析</span>
                <span class="example-tag" data-question="对比各部门业绩">📋 部门业绩对比</span>
            </div>
        </div>
    `;
    
    bindExampleTags();
    showToast('对话已清空', 'info');
}

// ===== 消息渲染 =====
function addMessage(role, content, extra = {}) {
    const msg = {
        id: generateId(),
        role,
        content,
        time: new Date(),
        ...extra
    };
    
    state.messages.push(msg);
    renderMessage(msg);
    scrollToBottom();
    
    return msg;
}

function renderMessage(msg) {
    // 隐藏空状态
    const emptyState = elements.chatMessages.querySelector('.empty-state');
    if (emptyState) {
        emptyState.style.display = 'none';
    }
    
    const msgEl = document.createElement('div');
    msgEl.className = `message-item ${msg.role}`;
    msgEl.id = `msg-${msg.id}`;
    
    const avatar = msg.role === 'user' ? '👤' : '🤖';
    
    let contentHtml;
    
    // 检查是否是原始 HTML（如执行计划）
    if (msg.isHtml) {
        // 直接使用 HTML，不转义
        contentHtml = `<div class="message-text">${msg.content}</div>`;
    } else {
        // 简单的 Markdown 渲染（加粗、换行）
        let textContent = escapeHtml(msg.content)
            .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')  // **加粗**
            .replace(/\n/g, '<br>');  // 换行
        
        contentHtml = `<div class="message-text">${textContent}</div>`;
    }
    
    // 代码块
    if (msg.code) {
        contentHtml += `
            <div class="message-code">
                <div class="code-header">
                    <span class="code-lang">Python / Pandas</span>
                    <button class="copy-btn" onclick="copyCode('${msg.id}')">复制代码</button>
                </div>
                <pre class="code-content" id="code-${msg.id}">${escapeHtml(msg.code)}</pre>
            </div>
        `;
    }
    
    // 数据表格和图表
    if (msg.data !== undefined && msg.data !== null) {
        // 判断数据类型
        if (Array.isArray(msg.data) && msg.data.length === 0) {
            // 空数组
            contentHtml += `
                <div class="message-empty">
                    📭 查询结果为空（没有符合条件的数据）
                </div>
            `;
        } else if (Array.isArray(msg.data) && msg.data.length > 0) {
            // DataFrame 数组格式
            const columns = Object.keys(msg.data[0]);
            const rows = msg.data.slice(0, 20);
            
            // 只在用户要求图表时显示（由后端 chart 字段控制）
            if (msg.chart?.show_chart) {
                const chartData = tryParseChartData(msg.data);
                if (chartData) {
                    // 根据用户要求的类型渲染图表
                    contentHtml += renderCharts(chartData, msg.id, msg.chart.chart_type);
                }
            }
            
            // 表格（支持排序和调整列宽）
            const tableId = `table-${msg.id}`;
            
            // Sheet 标签栏（拆分结果或多 Sheet 表格）
            let sheetTabsHtml = '';
            
            // 优先检查拆分结果
            if (msg.splitResult && msg.splitResult.sheets && msg.splitResult.sheets.length > 1) {
                const splitInfo = msg.splitResult;
                sheetTabsHtml = `
                    <div class="sheet-tabs split-tabs" data-split-id="${escapeHtml(splitInfo.id)}">
                        <span class="tabs-label">📑 按 ${escapeHtml(splitInfo.splitColumn)} 拆分 (${splitInfo.sheets.length}组):</span>
                        ${splitInfo.sheets.map(sheet => `
                            <button class="sheet-tab ${sheet.name === splitInfo.currentSheet ? 'active' : ''}" 
                                    data-sheet="${escapeHtml(sheet.name)}"
                                    onclick="switchSplitSheet('${msg.id}', '${escapeHtml(splitInfo.id)}', '${escapeHtml(sheet.name)}')">
                                ${escapeHtml(sheet.name)}
                                <span class="sheet-count">${sheet.row_count}</span>
                            </button>
                        `).join('')}
                    </div>
                `;
            } else if (msg.sourceTable && msg.sourceTable.sheets && msg.sourceTable.sheets.length > 1) {
                // 原有的多 Sheet 表格
                const sheets = msg.sourceTable.sheets;
                const currentSheet = msg.sourceTable.current_sheet || sheets[0];
                sheetTabsHtml = `
                    <div class="sheet-tabs" data-table-name="${escapeHtml(msg.sourceTable.name)}">
                        ${sheets.map(sheet => `
                            <button class="sheet-tab ${sheet === currentSheet ? 'active' : ''}" 
                                    data-sheet="${escapeHtml(sheet)}"
                                    onclick="switchSheetInMessage('${msg.id}', '${escapeHtml(msg.sourceTable.name)}', '${escapeHtml(sheet)}')">
                                ${escapeHtml(sheet)}
                            </button>
                        `).join('')}
                    </div>
                `;
            }
            
            contentHtml += `
                <div class="message-table">
                    ${sheetTabsHtml}
                    <div class="table-wrapper" id="wrapper-${tableId}">
                        <table class="data-table sortable-table" id="${tableId}" data-msg-id="${msg.id}">
                            <thead>
                                <tr>${columns.map((col, idx) => `
                                    <th class="sortable-th" data-col="${escapeHtml(col)}" data-idx="${idx}" onclick="sortTable('${tableId}', ${idx})">
                                        <span class="th-content">${escapeHtml(col)}</span>
                                        <span class="sort-icon">⇅</span>
                                        <div class="resize-handle" onmousedown="initResize(event, this)"></div>
                                    </th>
                                `).join('')}</tr>
                            </thead>
                            <tbody>
                                ${rows.map(row => `
                                    <tr>${columns.map(col => `<td>${escapeHtml(String(row[col] ?? ''))}</td>`).join('')}</tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                    <div class="table-footer">
                        <span class="table-info-inline">共 ${msg.totalRows || msg.data.length} 行${msg.data.length > 20 ? '（显示前20行）' : ''}</span>
                        <a class="export-link" onclick="exportData('${msg.id}', 'none')">📥 导出Excel</a>
                    </div>
                </div>
            `;
            
            // 存储数据用于导出
            window.exportDataStore = window.exportDataStore || {};
            window.exportDataStore[msg.id] = msg.data;
        } else if (typeof msg.data === 'object' && !Array.isArray(msg.data)) {
            // Series 或其他对象格式
            const entries = Object.entries(msg.data).slice(0, 20);
            if (entries.length > 0) {
                // 只在用户要求图表时显示
                if (msg.chart?.show_chart) {
                    const numericEntries = entries.filter(([k, v]) => typeof v === 'number' && !isNaN(v));
                    if (numericEntries.length >= 2 && numericEntries.length <= 30) {
                        const chartData = {
                            labels: numericEntries.map(([k]) => k),
                            values: numericEntries.map(([, v]) => v)
                        };
                        contentHtml += renderCharts(chartData, msg.id, msg.chart.chart_type);
                    }
                }
                
                contentHtml += `
                    <div class="message-table">
                        <table class="data-table">
                            <thead>
                                <tr><th>索引</th><th>值</th></tr>
                            </thead>
                            <tbody>
                                ${entries.map(([k, v]) => `
                                    <tr><td>${escapeHtml(String(k))}</td><td>${escapeHtml(String(v ?? ''))}</td></tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                    <div class="table-info">
                        共 ${Object.keys(msg.data).length} 项数据
                    </div>
                `;
            }
        } else if (typeof msg.data === 'string' || typeof msg.data === 'number') {
            // 单值结果
            contentHtml += `
                <div class="message-result">
                    <span class="result-label">结果：</span>
                    <span class="result-value">${escapeHtml(String(msg.data))}</span>
                </div>
            `;
        }
    }
    
    // 下载按钮（如果有下载链接）
    if (msg.download) {
        contentHtml += `
            <div class="message-download">
                <div class="download-info">
                    <span class="download-icon">📁</span>
                    <span class="download-desc">${escapeHtml(msg.download.description)}</span>
                </div>
                <button class="download-btn" onclick="downloadFile('${escapeHtml(msg.download.url)}', '${escapeHtml(msg.download.fileName)}')">
                    ⬇️ 下载文件
                </button>
            </div>
        `;
    }
    
    // 反馈按钮（仅助手消息）
    if (msg.role === 'assistant' && msg.status === 'success') {
        contentHtml += `
            <div class="message-feedback">
                <span>这个结果有帮助吗？</span>
                <button class="feedback-btn" onclick="setFeedback('${msg.id}', 'good')">👍 有帮助</button>
                <button class="feedback-btn" onclick="setFeedback('${msg.id}', 'bad')">👎 需改进</button>
            </div>
        `;
    }
    
    msgEl.innerHTML = `
        <div class="message-avatar ${msg.role}">${avatar}</div>
        <div class="message-content">${contentHtml}</div>
    `;
    
    elements.chatMessages.appendChild(msgEl);
}

function renderTypingIndicator() {
    const typingEl = document.createElement('div');
    typingEl.className = 'message-item assistant';
    typingEl.id = 'typing-indicator';
    typingEl.innerHTML = `
        <div class="message-avatar assistant">🤖</div>
        <div class="message-content">
            <div class="typing-indicator">
                <span></span>
                <span></span>
                <span></span>
            </div>
        </div>
    `;
    elements.chatMessages.appendChild(typingEl);
    scrollToBottom();
}

function removeTypingIndicator() {
    const typing = $('#typing-indicator');
    if (typing) typing.remove();
}

function scrollToBottom() {
    elements.chatMessages.scrollTop = elements.chatMessages.scrollHeight;
}

// ===== 全局函数 =====
window.copyCode = function(msgId) {
    const codeEl = $(`#code-${msgId}`);
    if (codeEl) {
        navigator.clipboard.writeText(codeEl.textContent);
        showToast('代码已复制', 'success');
    }
};

window.downloadFile = function(url, fileName) {
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    showToast(`开始下载: ${fileName}`, 'success');
};

window.setFeedback = function(msgId, type) {
    const msg = state.messages.find(m => m.id === msgId);
    if (!msg) return;
    
    msg.feedback = type;
    
    // 更新按钮样式
    const msgEl = $(`#msg-${msgId}`);
    if (msgEl) {
        const btns = msgEl.querySelectorAll('.feedback-btn');
        btns.forEach(btn => {
            btn.classList.remove('active-good', 'active-bad');
        });
        
        if (type === 'good') {
            btns[0].classList.add('active-good');
            showToast('感谢您的反馈！', 'success');
        } else {
            btns[1].classList.add('active-bad');
            showToast('感谢反馈，我们会继续改进', 'info');
        }
    }
};

// 显示表格操作菜单
window.showTableActions = function(tableName) {
    // 移除已存在的菜单
    const existingMenu = document.querySelector('.table-action-menu');
    if (existingMenu) existingMenu.remove();
    
    const table = state.tables.find(t => t.name === tableName);
    if (!table) return;
    
    const alias = table.alias || '';
    const nickname = table.nickname || '';
    const displayTitle = nickname ? `${alias}: ${nickname}` : `${alias}表`;
    const otherTables = state.tables.filter(t => t.name !== tableName);
    
    let menuHtml = `
        <div class="table-action-menu" id="table-menu-${tableName}">
            <div class="menu-header">${escapeHtml(displayTitle)}</div>
            <div class="menu-item" onclick="viewTableDetails('${escapeHtml(tableName)}')">
                📊 查看详情
            </div>
            <div class="menu-item" onclick="editNickname(event, '${escapeHtml(tableName)}'); closeTableMenu();">
                ✏️ ${nickname ? '修改昵称' : '设置昵称'}
            </div>
            <div class="menu-divider"></div>
            <div class="menu-item danger" onclick="removeTable('${escapeHtml(tableName)}')">
                🗑️ 移除此表
            </div>
        </div>
    `;
    
    document.body.insertAdjacentHTML('beforeend', menuHtml);
    
    // 点击其他地方关闭菜单
    setTimeout(() => {
        document.addEventListener('click', closeTableMenu, { once: true });
    }, 10);
};

function closeTableMenu() {
    const menu = document.querySelector('.table-action-menu');
    if (menu) menu.remove();
}

// 拆分导出对话框
window.showSplitToSheetsDialog = async function(tableName) {
    closeTableMenu();
    
    const table = state.tables.find(t => t.name === tableName);
    if (!table) return;
    
    const alias = table.alias || '';
    const nickname = table.nickname || '';
    const displayTitle = nickname ? `${alias}: ${nickname}` : `${alias}表`;
    
    // 获取列信息
    showToast('正在获取列信息...', 'info');
    const result = await apiRequest('GET', `/api/table-columns/${tableName}`);
    
    if (!result.success) {
        showToast('获取列信息失败', 'error');
        return;
    }
    
    const columns = result.data.columns;
    const suitableColumns = columns.filter(c => c.suitable_for_split);
    
    let columnsHtml = columns.map(col => {
        const suitableClass = col.suitable_for_split ? 'suitable' : 'not-suitable';
        const suitableIcon = col.suitable_for_split ? '✓' : '⚠';
        const tooltip = col.suitable_for_split 
            ? `${col.unique_count} 个不同值，适合拆分` 
            : col.unique_count <= 1 ? '值太少，不适合拆分' : '值太多（超过100），不适合拆分';
        return `
            <div class="split-column-option ${suitableClass}" 
                 onclick="selectSplitColumn('${escapeHtml(col.name)}')"
                 data-column="${escapeHtml(col.name)}"
                 title="${tooltip}">
                <span class="column-name">${escapeHtml(col.name)}</span>
                <span class="column-info">${suitableIcon} ${col.unique_count}个值</span>
            </div>
        `;
    }).join('');
    
    const dialogHtml = `
        <div class="modal-overlay" id="split-dialog" onclick="closeSplitDialog(event)">
            <div class="modal-content split-dialog-content" onclick="event.stopPropagation()">
                <div class="modal-header">
                    <h3>📑 按字段拆分导出</h3>
                    <button class="modal-close" onclick="closeSplitDialog()">×</button>
                </div>
                <div class="modal-body">
                    <p class="split-tip">选择一个字段，将 <strong>${escapeHtml(displayTitle)}</strong> 按该字段的不同值拆分成多个 Sheet 导出</p>
                    <div class="split-columns-list" id="split-columns-list">
                        ${columnsHtml}
                    </div>
                    <div class="split-selected" id="split-selected" style="display: none;">
                        <span>已选择: </span>
                        <strong id="selected-column-name"></strong>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" onclick="closeSplitDialog()">取消</button>
                    <button class="btn btn-primary" id="btn-do-split" onclick="doSplitExport('${escapeHtml(tableName)}')" disabled>
                        导出拆分文件
                    </button>
                </div>
            </div>
        </div>
    `;
    
    document.body.insertAdjacentHTML('beforeend', dialogHtml);
};

let selectedSplitColumn = null;

window.selectSplitColumn = function(columnName) {
    selectedSplitColumn = columnName;
    
    // 更新选中状态
    document.querySelectorAll('.split-column-option').forEach(el => {
        el.classList.remove('selected');
        if (el.dataset.column === columnName) {
            el.classList.add('selected');
        }
    });
    
    // 显示已选择
    document.getElementById('split-selected').style.display = 'flex';
    document.getElementById('selected-column-name').textContent = columnName;
    document.getElementById('btn-do-split').disabled = false;
};

window.closeSplitDialog = function(event) {
    if (event && event.target !== event.currentTarget) return;
    const dialog = document.getElementById('split-dialog');
    if (dialog) dialog.remove();
    selectedSplitColumn = null;
};

window.doSplitExport = async function(tableName) {
    if (!selectedSplitColumn) {
        showToast('请选择拆分字段', 'warning');
        return;
    }
    
    const table = state.tables.find(t => t.name === tableName);
    const alias = table?.alias || '';
    const nickname = table?.nickname || '';
    const displayTitle = nickname ? `${alias}_${nickname}` : `${alias}表`;
    
    showToast('正在拆分导出...', 'info');
    
    try {
        const result = await apiRequest('POST', '/api/split-to-sheets', {
            table_name: tableName,
            split_column: selectedSplitColumn,
            file_name: `${displayTitle}_按${selectedSplitColumn}拆分`
        });
        
        if (result.success && result.data.download_url) {
            // 自动下载
            const downloadUrl = `http://127.0.0.1:5000${result.data.download_url}`;
            const a = document.createElement('a');
            a.href = downloadUrl;
            a.download = result.data.file_name;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            
            showToast(`✅ 已拆分成 ${result.data.sheet_count} 个 Sheet`, 'success');
            closeSplitDialog();
        } else {
            showToast('拆分导出失败: ' + (result.data?.detail || result.error || '未知错误'), 'error');
        }
    } catch (e) {
        showToast('拆分导出失败: ' + e.message, 'error');
    }
};

// 查看表详情（可编辑）
window.viewTableDetails = function(tableName) {
    closeTableMenu();
    const table = state.tables.find(t => t.name === tableName);
    if (!table) return;
    
    const alias = table.alias || '';
    const nickname = table.nickname || '';
    const sheets = table.sheets || ['Sheet1'];
    const currentSheet = table.current_sheet || sheets[0];
    const headerRow = table.header_row || 1;
    
    let detailHtml = `
        <div class="modal-overlay" onclick="closeModal(event)">
            <div class="modal-content modal-editable" onclick="event.stopPropagation()">
                <div class="modal-header">
                    <div class="modal-title-with-alias">
                        <span class="modal-alias">${alias}</span>
                        <h3>${escapeHtml(nickname || tableName)}</h3>
                    </div>
                    <button class="modal-close" onclick="closeModal()">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="detail-section">
                        <div class="detail-label">表格标识</div>
                        <div class="setting-row">
                            <label>别名:</label>
                            <span class="alias-badge">${alias}表</span>
                            <span class="setting-hint">（对话时使用"${alias}表"引用此表）</span>
                        </div>
                        <div class="setting-row">
                            <label>昵称:</label>
                            <input type="text" id="nickname-input" class="setting-input" value="${escapeHtml(nickname)}" placeholder="可选，如: 收入表" />
                        </div>
                        <div class="setting-row">
                            <label>原表名:</label>
                            <span class="original-name" title="${escapeHtml(tableName)}">${escapeHtml(tableName)}</span>
                        </div>
                    </div>
                    
                    <div class="detail-section">
                        <div class="detail-label">数据源设置</div>
                        
                        <div class="setting-row">
                            <label>选择 Sheet:</label>
                            <select id="sheet-select" class="setting-select">
                                ${sheets.map(s => `<option value="${escapeHtml(s)}" ${s === currentSheet ? 'selected' : ''}>${escapeHtml(s)}</option>`).join('')}
                            </select>
                        </div>
                        
                        <div class="setting-row">
                            <label>表头行号:</label>
                            <input type="number" id="header-row-input" class="setting-input" value="${headerRow}" min="1" max="20" />
                            <span class="setting-hint">（第几行是列名）</span>
                        </div>
                    </div>
                    
                    <div class="detail-section">
                        <div class="detail-label">当前数据信息</div>
                        <div class="detail-row">行数: <strong>${table.row_count}</strong></div>
                        <div class="detail-row">列数: <strong>${table.columns.length}</strong></div>
                    </div>
                    
                    <div class="detail-section">
                        <div class="detail-label">识别的列名</div>
                        <div class="columns-list" id="columns-preview">
                            ${table.columns.map(col => `<span class="column-tag">${escapeHtml(col)}</span>`).join('')}
                        </div>
                    </div>
                    
                    <div class="modal-actions">
                        <button class="btn-secondary" onclick="closeModal()">取消</button>
                        <button class="btn-primary" onclick="reloadTableWithSettings('${escapeHtml(tableName)}')">
                            🔄 应用设置
                        </button>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    document.body.insertAdjacentHTML('beforeend', detailHtml);
};

// 使用新设置重新加载表格
window.reloadTableWithSettings = async function(tableName) {
    const table = state.tables.find(t => t.name === tableName);
    if (!table || !table.file_path) {
        showToast('无法重新加载：缺少文件路径', 'error');
        return;
    }
    
    const sheetSelect = document.getElementById('sheet-select');
    const headerRowInput = document.getElementById('header-row-input');
    const nicknameInput = document.getElementById('nickname-input');
    
    const newSheet = sheetSelect?.value;
    const newHeaderRow = parseInt(headerRowInput?.value) || 1;
    const newNickname = nicknameInput?.value?.trim() || '';
    
    showToast('正在应用设置...', 'info');
    
    try {
        // 先更新昵称
        if (newNickname !== (table.nickname || '')) {
            await apiRequest('POST', '/api/update-nickname', {
                table_name: tableName,
                nickname: newNickname
            });
            table.nickname = newNickname;
        }
        
        // 如果 sheet 或表头行改变了，重新加载
        const sheetChanged = newSheet !== table.current_sheet;
        const headerChanged = newHeaderRow !== table.header_row;
        
        if (sheetChanged || headerChanged) {
            const result = await apiRequest('POST', '/api/reload-table', {
                table_name: tableName,
                sheet_name: newSheet,
                header_row: newHeaderRow
            });
            
            if (result.success && result.data.success) {
                // 更新本地状态
                const idx = state.tables.findIndex(t => t.name === tableName);
                if (idx >= 0) {
                    state.tables[idx] = {
                        ...state.tables[idx],
                        ...result.data.table,
                        nickname: newNickname,
                        current_sheet: newSheet,
                        header_row: newHeaderRow
                    };
                }
            } else {
                showToast('加载失败: ' + (result.data?.detail || result.error), 'error');
                return;
            }
        }
        
        updateTablesUI();
        closeModal();
        showToast(`已更新「${table.alias}表」设置`, 'success');
    } catch (e) {
        showToast('操作失败: ' + e.message, 'error');
    }
};

// 原来的简单重载逻辑（兼容）
async function simpleReloadTable(tableName, newSheet, newHeaderRow) {
    try {
        const result = await apiRequest('POST', '/api/reload-table', {
            table_name: tableName,
            sheet_name: newSheet,
            header_row: newHeaderRow
        });
        
        if (result.success && result.data.success) {
            // 更新本地状态
            const idx = state.tables.findIndex(t => t.name === tableName);
            if (idx >= 0) {
                state.tables[idx] = {
                    ...state.tables[idx],
                    ...result.data.table,
                    current_sheet: newSheet,
                    header_row: newHeaderRow
                };
            }
            
            updateTablesUI();
            closeModal();
            showToast(`已重新加载「${tableName}」`, 'success');
        } else {
            showToast('加载失败: ' + (result.data?.detail || result.error), 'error');
        }
    } catch (e) {
        showToast('加载失败: ' + e.message, 'error');
    }
}

// 提示合并到 Sheet
window.promptMergeToSheet = function(sourceTable, targetTable) {
    closeTableMenu();
    
    const sheetName = prompt(`将「${sourceTable}」合并到「${targetTable}」的新 Sheet\n\n请输入新 Sheet 的名称:`, sourceTable);
    
    if (sheetName) {
        mergeToSheet(sourceTable, targetTable, sheetName);
    }
};

// 执行合并
async function mergeToSheet(sourceTable, targetTable, sheetName) {
    showToast('正在合并...', 'info');
    
    try {
        const result = await apiRequest('POST', '/api/merge-to-sheet', {
            source_table: sourceTable,
            target_table: targetTable,
            new_sheet_name: sheetName,
            include_chart: false,
        });
        
        if (result.success && result.data.success) {
            // 触发下载
            const downloadUrl = `http://127.0.0.1:5000${result.data.download_url}`;
            const a = document.createElement('a');
            a.href = downloadUrl;
            a.download = result.data.file_name;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            
            showToast(`合并成功！新 Sheet: ${result.data.sheet_name}`, 'success');
        } else {
            showToast('合并失败: ' + (result.data?.detail || result.error || '未知错误'), 'error');
        }
    } catch (error) {
        console.error('Merge error:', error);
        showToast('合并失败: ' + error.message, 'error');
    }
}

// 移除表
window.removeTable = async function(tableName) {
    closeTableMenu();
    
    const table = state.tables.find(t => t.name === tableName);
    const displayName = table ? `${table.alias}表` : tableName;
    
    if (confirm(`确定要移除「${displayName}」吗？`)) {
        try {
            // 调用后端删除 API
            await apiRequest('DELETE', `/api/tables/${encodeURIComponent(tableName)}`);
            
            state.tables = state.tables.filter(t => t.name !== tableName);
            updateTablesUI();
            showToast(`已移除「${displayName}」`, 'info');
        } catch (e) {
            showToast('移除失败: ' + e.message, 'error');
        }
    }
};

// 关闭模态框
window.closeModal = function(event) {
    if (!event || event.target.classList.contains('modal-overlay')) {
        const modal = document.querySelector('.modal-overlay');
        if (modal) modal.remove();
    }
};

// 导出数据到 Excel
window.exportData = async function(msgId, chartType) {
    const data = window.exportDataStore?.[msgId];
    if (!data || !Array.isArray(data) || data.length === 0) {
        showToast('没有可导出的数据', 'warning');
        return;
    }
    
    showToast('正在导出...', 'info');
    
    try {
        const result = await apiRequest('POST', '/api/export', {
            data: data,
            file_name: '数据分析结果',
            include_chart: chartType !== 'none',
            chart_type: chartType === 'none' ? 'bar' : chartType,
        });
        
        if (result.success && result.data.success) {
            // 触发下载
            const downloadUrl = `http://127.0.0.1:5000${result.data.download_url}`;
            
            // 创建隐藏的 a 标签下载
            const a = document.createElement('a');
            a.href = downloadUrl;
            a.download = result.data.file_name;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            
            showToast(`导出成功: ${result.data.file_name}`, 'success');
        } else {
            showToast('导出失败: ' + (result.error || '未知错误'), 'error');
        }
    } catch (error) {
        console.error('Export error:', error);
        showToast('导出失败: ' + error.message, 'error');
    }
};

// ===== 发送消息 =====
async function sendMessage() {
    const text = elements.userInput.value.trim();
    if (!text || state.isLoading) return;
    
    if (state.tables.length === 0) {
        showToast('请先上传数据表', 'warning');
        return;
    }
    
    // 添加用户消息
    addMessage('user', text);
    elements.userInput.value = '';
    elements.userInput.style.height = 'auto'; // 重置输入框高度
    
    // 添加到历史
    addToHistory(text);
    
    // 如果是新对话的第一条消息，自动生成名称
    autoNameAgent(text);
    
    // 显示加载状态
    renderTypingIndicator();
    state.isLoading = true;
    
    try {
        // 先分析任务是否需要分步执行
        const analyzeResult = await apiRequest('POST', '/api/analyze-task', {
            user_description: text
        });
        
        if (analyzeResult.success && analyzeResult.data?.is_multi_step) {
            // 多步骤任务，显示执行计划
            removeTypingIndicator();
            showTaskPlan(analyzeResult.data, text);
            return;
        }
        
        // 单步任务，直接执行（typing indicator 已经显示，不需要重复）
        await executeSimpleTask(text, true);
        
    } catch (error) {
        removeTypingIndicator();
        addMessage('assistant', `❌ 发生错误: ${error.message}`, { status: 'error' });
    } finally {
        state.isLoading = false;
    }
}

// 执行简单任务（单步）
async function executeSimpleTask(text, skipTypingIndicator = false) {
    // 如果调用方已经显示了 typing indicator，就不再重复显示
    if (!skipTypingIndicator) {
        renderTypingIndicator();
    }
    
    try {
        const result = await apiRequest('POST', '/api/smart-execute', {
            user_description: text
        });
        
        removeTypingIndicator();
        
        if (!result.success) {
            addMessage('assistant', `❌ 请求失败: ${result.error || '未知错误'}`, { status: 'error' });
            return;
        }
        
        const data = result.data;
        
        if (data.status === 'success') {
            // 成功
            const resultData = data.result?.data || [];
            
            // 检查是否包含下载链接（合并/导出操作）
            let downloadInfo = null;
            if (Array.isArray(resultData) && resultData.length > 0 && resultData[0]['下载链接']) {
                downloadInfo = {
                    url: `http://127.0.0.1:5000${resultData[0]['下载链接']}`,
                    fileName: resultData[0]['文件名'] || '导出文件.xlsx',
                    description: resultData[0]['说明'] || '文件已生成'
                };
            }
            
            // 构建回复消息（包含分析过程）
            const analysisText = data.analysis ? `💡 **分析思路：** ${data.analysis}\n\n` : '';
            const messageText = `${analysisText}✅ 分析完成！以下是结果：`;
            
            addMessage('assistant', messageText, {
                code: data.code,
                data: resultData,
                totalRows: data.result?.total_rows,
                status: 'success',
                download: downloadInfo,
                chart: data.chart,  // 图表意图（由后端根据用户问题检测）
                sourceTable: data.sourceTable,  // 原表信息（用于 Sheet 切换）
                splitResult: data.splitResult  // 拆分结果（用于显示分组标签）
            });
            
        } else if (data.status === 'need_clarification') {
            // 需要澄清
            let clarifyMsg = `⚠️ ${data.message}\n\n`;
            if (data.options && data.options.length > 0) {
                clarifyMsg += '可选项：' + data.options.slice(0, 10).join('、');
            }
            addMessage('assistant', clarifyMsg, { status: 'clarification' });
            
        } else {
            // 错误
            addMessage('assistant', `❌ ${data.message || '执行失败'}`, { 
                code: data.code,
                status: 'error' 
            });
        }
        
    } catch (error) {
        removeTypingIndicator();
        throw error;
    }
}

// 显示任务执行计划
function showTaskPlan(planData, originalQuery) {
    const planId = generateId();
    const steps = planData.steps || [];
    
    // 构建步骤 HTML
    const stepsHtml = steps.map((step, idx) => {
        const inputStr = step.input.map(i => {
            if (i.startsWith('step_')) return `[${i}]`;
            return `${i}表`;
        }).join(' + ');
        
        const isLast = idx === steps.length - 1;
        
        return `
            <div class="plan-step" data-step-id="${step.id}" data-status="pending">
                <div class="step-number">${step.id}</div>
                <div class="step-content">
                    <div class="step-desc">${escapeHtml(step.description)}</div>
                    <div class="step-meta">
                        <span class="step-input">📥 ${escapeHtml(inputStr)}</span>
                        <span class="step-arrow">→</span>
                        <span class="step-output">📤 ${isLast ? '最终结果' : step.output}</span>
                    </div>
                </div>
                <div class="step-status">
                    <span class="status-icon">⏳</span>
                </div>
            </div>
        `;
    }).join('');
    
    const planHtml = `
        <div class="task-plan" id="plan-${planId}">
            <div class="plan-header">
                <div class="plan-title">
                    <span class="plan-icon">📋</span>
                    <span>执行计划（共${steps.length}步）</span>
                </div>
                <div class="plan-message">${escapeHtml(planData.message || '')}</div>
            </div>
            <div class="plan-steps">
                ${stepsHtml}
            </div>
            <div class="plan-actions">
                <button class="btn-execute-plan" onclick="executePlan('${planId}', ${JSON.stringify(steps).replace(/"/g, '&quot;')}, '${escapeHtml(originalQuery)}')">
                    ▶️ 开始执行
                </button>
                <button class="btn-cancel-plan" onclick="cancelPlan('${planId}')">
                    ✖️ 取消
                </button>
                <button class="btn-direct-execute" onclick="directExecute('${escapeHtml(originalQuery)}')">
                    ⚡ 直接执行（不分步）
                </button>
            </div>
        </div>
    `;
    
    addMessage('assistant', planHtml, { isHtml: true, status: 'plan' });
}

// 执行计划（逐步）
async function executePlan(planId, steps, originalQuery) {
    const planEl = document.getElementById(`plan-${planId}`);
    if (!planEl) return;
    
    // 禁用按钮
    const executeBtn = planEl.querySelector('.btn-execute-plan');
    const cancelBtn = planEl.querySelector('.btn-cancel-plan');
    const directBtn = planEl.querySelector('.btn-direct-execute');
    if (executeBtn) executeBtn.disabled = true;
    if (cancelBtn) cancelBtn.disabled = true;
    if (directBtn) directBtn.disabled = true;
    
    // 清除之前的中间结果
    await apiRequest('POST', '/api/clear-step-results');
    
    // 逐步执行
    for (let i = 0; i < steps.length; i++) {
        const step = steps[i];
        const stepEl = planEl.querySelector(`[data-step-id="${step.id}"]`);
        
        if (stepEl) {
            stepEl.dataset.status = 'running';
            stepEl.querySelector('.status-icon').textContent = '⏳';
            stepEl.querySelector('.status-icon').classList.add('spinning');
        }
        
        try {
            const result = await apiRequest('POST', '/api/execute-step', {
                step_id: step.id,
                step_description: step.description,
                input_sources: step.input,
                output_id: step.output
            });
            
            if (result.success && result.data?.success) {
                // 步骤成功
                if (stepEl) {
                    stepEl.dataset.status = 'completed';
                    stepEl.querySelector('.status-icon').textContent = '✅';
                    stepEl.querySelector('.status-icon').classList.remove('spinning');
                    
                    // 添加结果预览（可折叠）
                    const resultData = result.data.result?.data || [];
                    if (resultData.length > 0) {
                        const previewHtml = `
                            <div class="step-result-preview">
                                <button class="toggle-preview" onclick="toggleStepResult(this)">
                                    📊 查看结果（${result.data.result.total_rows}行）
                                </button>
                                <div class="step-result-table" style="display:none;">
                                    ${renderMiniTable(resultData.slice(0, 5), result.data.result.columns)}
                                    ${resultData.length > 5 ? `<div class="more-rows">... 还有 ${result.data.result.total_rows - 5} 行</div>` : ''}
                                    <button class="btn-save-step" onclick="saveStepAsTable('${step.output}')">
                                        💾 保存为新表
                                    </button>
                                </div>
                            </div>
                        `;
                        stepEl.querySelector('.step-content').insertAdjacentHTML('beforeend', previewHtml);
                    }
                }
                
                // 如果是最后一步，显示最终结果
                if (i === steps.length - 1) {
                    const finalData = result.data.result?.data || [];
                    addMessage('assistant', '✅ 所有步骤执行完成！以下是最终结果：', {
                        data: finalData,
                        totalRows: result.data.result?.total_rows,
                        status: 'success'
                    });
                }
                
            } else {
                // 步骤失败
                if (stepEl) {
                    stepEl.dataset.status = 'failed';
                    stepEl.querySelector('.status-icon').textContent = '❌';
                    stepEl.querySelector('.status-icon').classList.remove('spinning');
                    
                    // 显示错误
                    const errorHtml = `<div class="step-error">${escapeHtml(result.data?.error || '执行失败')}</div>`;
                    stepEl.querySelector('.step-content').insertAdjacentHTML('beforeend', errorHtml);
                }
                
                addMessage('assistant', `❌ 步骤 ${step.id} 执行失败: ${result.data?.error || '未知错误'}`, { status: 'error' });
                break;
            }
            
        } catch (error) {
            if (stepEl) {
                stepEl.dataset.status = 'failed';
                stepEl.querySelector('.status-icon').textContent = '❌';
                stepEl.querySelector('.status-icon').classList.remove('spinning');
            }
            addMessage('assistant', `❌ 步骤 ${step.id} 发生错误: ${error.message}`, { status: 'error' });
            break;
        }
    }
    
    state.isLoading = false;
}

// 取消计划
function cancelPlan(planId) {
    const planEl = document.getElementById(`plan-${planId}`);
    if (planEl) {
        planEl.classList.add('cancelled');
        const actionsEl = planEl.querySelector('.plan-actions');
        if (actionsEl) {
            actionsEl.innerHTML = '<div class="plan-cancelled">已取消</div>';
        }
    }
}

// 直接执行（不分步）
async function directExecute(query) {
    state.isLoading = true;
    await executeSimpleTask(query);
    state.isLoading = false;
}

// 切换步骤结果显示
function toggleStepResult(btn) {
    const tableEl = btn.nextElementSibling;
    if (tableEl) {
        const isVisible = tableEl.style.display !== 'none';
        tableEl.style.display = isVisible ? 'none' : 'block';
        btn.textContent = isVisible ? btn.textContent.replace('收起', '查看') : btn.textContent.replace('查看', '收起');
    }
}

// 保存步骤结果为新表
async function saveStepAsTable(outputId) {
    const nickname = await showPromptModal('保存为新表', '请输入表的名称（可选）：', '');
    
    try {
        const result = await apiRequest('POST', `/api/save-step-as-table?output_id=${outputId}&nickname=${encodeURIComponent(nickname || '')}`);
        
        if (result.success && result.data?.success) {
            showToast(`已保存为 ${result.data.alias}表`, 'success');
            await loadTables();  // 刷新表列表
        } else {
            showToast(result.data?.error || '保存失败', 'error');
        }
    } catch (error) {
        showToast(`保存失败: ${error.message}`, 'error');
    }
}

// 渲染迷你表格（用于步骤结果预览）
function renderMiniTable(data, columns) {
    if (!data || data.length === 0) return '';
    
    const headers = columns.slice(0, 5);
    const headerHtml = headers.map(col => `<th>${escapeHtml(col)}</th>`).join('');
    
    const rowsHtml = data.map(row => {
        const cells = headers.map(col => {
            const val = row[col];
            return `<td>${escapeHtml(String(val ?? ''))}</td>`;
        }).join('');
        return `<tr>${cells}</tr>`;
    }).join('');
    
    return `
        <table class="mini-table">
            <thead><tr>${headerHtml}${columns.length > 5 ? '<th>...</th>' : ''}</tr></thead>
            <tbody>${rowsHtml}</tbody>
        </table>
    `;
}

// 暴露任务计划相关函数到全局
window.executePlan = executePlan;
window.cancelPlan = cancelPlan;
window.directExecute = directExecute;
window.toggleStepResult = toggleStepResult;
window.saveStepAsTable = saveStepAsTable;
window.toggleSidebar = toggleSidebar;

// ===== 表格排序 =====
const tableSortState = {};  // 记录每个表格每列的排序状态

// 切换拆分结果的 Sheet
window.switchSplitSheet = async function(msgId, splitId, sheetName) {
    const wrapper = document.getElementById(`wrapper-table-${msgId}`);
    if (!wrapper) return;
    
    // 更新标签状态
    const tabs = wrapper.parentElement.querySelectorAll('.sheet-tab');
    tabs.forEach(tab => {
        tab.classList.toggle('active', tab.dataset.sheet === sheetName);
    });
    
    // 显示加载状态
    wrapper.innerHTML = '<div class="sheet-loading">正在加载...</div>';
    
    try {
        const result = await apiRequest('POST', '/api/split-sheet-data', {
            split_id: splitId,
            sheet_name: sheetName,
            limit: 20
        });
        
        if (result.success && result.data) {
            const { columns, data, total_rows } = result.data;
            const tableId = `table-${msgId}`;
            
            wrapper.innerHTML = `
                <table class="data-table sortable-table" id="${tableId}" data-msg-id="${msgId}">
                    <thead>
                        <tr>${columns.map((col, idx) => `
                            <th class="sortable-th" data-col="${escapeHtml(col)}" data-idx="${idx}" onclick="sortTable('${tableId}', ${idx})">
                                <span class="th-content">${escapeHtml(col)}</span>
                                <span class="sort-icon">⇅</span>
                                <div class="resize-handle" onmousedown="initResize(event, this)"></div>
                            </th>
                        `).join('')}</tr>
                    </thead>
                    <tbody>
                        ${data.map(row => `
                            <tr>${columns.map(col => `<td>${escapeHtml(String(row[col] ?? ''))}</td>`).join('')}</tr>
                        `).join('')}
                    </tbody>
                </table>
            `;
            
            // 更新底部信息
            const footer = wrapper.parentElement.querySelector('.table-footer .table-info-inline');
            if (footer) {
                footer.textContent = `共 ${total_rows} 行${data.length < total_rows ? '（显示前20行）' : ''}`;
            }
            
            // 更新导出数据
            window.exportDataStore = window.exportDataStore || {};
            window.exportDataStore[msgId] = data;
            
        } else {
            wrapper.innerHTML = `<div class="sheet-error">加载失败: ${result.error || '未知错误'}</div>`;
        }
    } catch (error) {
        wrapper.innerHTML = `<div class="sheet-error">加载失败: ${error.message}</div>`;
    }
};

// 切换消息中表格的 Sheet
window.switchSheetInMessage = async function(msgId, tableName, sheetName) {
    const wrapper = document.getElementById(`wrapper-table-${msgId}`);
    if (!wrapper) return;
    
    // 更新标签状态
    const tabs = wrapper.parentElement.querySelectorAll('.sheet-tab');
    tabs.forEach(tab => {
        tab.classList.toggle('active', tab.dataset.sheet === sheetName);
    });
    
    // 显示加载状态
    wrapper.innerHTML = '<div class="sheet-loading">正在加载...</div>';
    
    try {
        // 调用后端加载指定 Sheet 的数据
        const result = await apiRequest('POST', '/api/preview-sheet', {
            table_name: tableName,
            sheet_name: sheetName,
            limit: 20
        });
        
        if (result.success && result.data) {
            const { columns, data, total_rows } = result.data;
            const tableId = `table-${msgId}`;
            
            // 重新渲染表格
            wrapper.innerHTML = `
                <table class="data-table sortable-table" id="${tableId}" data-msg-id="${msgId}">
                    <thead>
                        <tr>${columns.map((col, idx) => `
                            <th class="sortable-th" data-col="${escapeHtml(col)}" data-idx="${idx}" onclick="sortTable('${tableId}', ${idx})">
                                <span class="th-content">${escapeHtml(col)}</span>
                                <span class="sort-icon">⇅</span>
                                <div class="resize-handle" onmousedown="initResize(event, this)"></div>
                            </th>
                        `).join('')}</tr>
                    </thead>
                    <tbody>
                        ${data.map(row => `
                            <tr>${columns.map(col => `<td>${escapeHtml(String(row[col] ?? ''))}</td>`).join('')}</tr>
                        `).join('')}
                    </tbody>
                </table>
            `;
            
            // 更新底部信息
            const footer = wrapper.parentElement.querySelector('.table-footer .table-info-inline');
            if (footer) {
                footer.textContent = `共 ${total_rows} 行${data.length > 20 ? '（显示前20行）' : ''}`;
            }
            
            // 更新导出数据
            window.exportDataStore = window.exportDataStore || {};
            window.exportDataStore[msgId] = data;
            
        } else {
            wrapper.innerHTML = `<div class="sheet-error">加载失败: ${result.error || '未知错误'}</div>`;
        }
    } catch (error) {
        wrapper.innerHTML = `<div class="sheet-error">加载失败: ${error.message}</div>`;
    }
};

window.sortTable = function(tableId, colIdx) {
    const table = document.getElementById(tableId);
    if (!table) return;
    
    const tbody = table.querySelector('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    const headers = table.querySelectorAll('th');
    
    // 获取当前排序状态
    const stateKey = `${tableId}-${colIdx}`;
    const currentOrder = tableSortState[stateKey] || 'none';
    const newOrder = currentOrder === 'asc' ? 'desc' : 'asc';
    tableSortState[stateKey] = newOrder;
    
    // 更新所有列的排序图标
    headers.forEach((th, idx) => {
        const icon = th.querySelector('.sort-icon');
        if (icon) {
            if (idx === colIdx) {
                icon.textContent = newOrder === 'asc' ? '↑' : '↓';
                th.classList.add('sorted');
            } else {
                icon.textContent = '⇅';
                th.classList.remove('sorted');
            }
        }
    });
    
    // 排序
    rows.sort((a, b) => {
        const aVal = a.cells[colIdx]?.textContent.trim() || '';
        const bVal = b.cells[colIdx]?.textContent.trim() || '';
        
        // 尝试数字排序
        const aNum = parseFloat(aVal.replace(/,/g, ''));
        const bNum = parseFloat(bVal.replace(/,/g, ''));
        
        if (!isNaN(aNum) && !isNaN(bNum)) {
            return newOrder === 'asc' ? aNum - bNum : bNum - aNum;
        }
        
        // 字符串排序
        return newOrder === 'asc' 
            ? aVal.localeCompare(bVal, 'zh-CN')
            : bVal.localeCompare(aVal, 'zh-CN');
    });
    
    // 重新插入排序后的行
    rows.forEach(row => tbody.appendChild(row));
};

// ===== 列宽调整 =====
let resizing = null;

window.initResize = function(e, handle) {
    e.preventDefault();
    e.stopPropagation();
    
    const th = handle.parentElement;
    const startX = e.pageX;
    const startWidth = th.offsetWidth;
    
    resizing = { th, startX, startWidth };
    
    document.addEventListener('mousemove', doResize);
    document.addEventListener('mouseup', stopResize);
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
};

function doResize(e) {
    if (!resizing) return;
    
    const { th, startX, startWidth } = resizing;
    const diff = e.pageX - startX;
    const newWidth = Math.max(50, startWidth + diff);
    th.style.width = newWidth + 'px';
    th.style.minWidth = newWidth + 'px';
}

function stopResize() {
    resizing = null;
    document.removeEventListener('mousemove', doResize);
    document.removeEventListener('mouseup', stopResize);
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
}

// ===== 历史记录 =====
function addToHistory(text) {
    state.history.unshift({
        text,
        time: new Date()
    });
    
    // 最多保留10条
    if (state.history.length > 10) {
        state.history.pop();
    }
    
    updateHistoryUI();
}

// 从后端加载执行历史
async function loadExecutionHistory() {
    try {
        const result = await apiRequest('GET', '/api/history');
        if (result.success && result.data?.history) {
            return result.data.history;
        }
    } catch (e) {
        console.error('加载执行历史失败:', e);
    }
    return [];
}

// ===== 执行历史弹框 =====
window.openExecutionHistoryModal = async function() {
    const modal = document.getElementById('execution-history-modal');
    const listEl = document.getElementById('execution-history-list');
    
    if (!modal || !listEl) return;
    
    modal.style.display = 'flex';
    listEl.innerHTML = '<div class="loading-hint">加载中...</div>';
    
    // 加载历史记录
    const history = await loadExecutionHistory();
    
    if (history.length === 0) {
        listEl.innerHTML = `
            <div class="empty-execution-history">
                <div class="empty-icon">📭</div>
                <div class="empty-text">暂无执行历史</div>
                <div class="empty-hint">执行数据分析后，记录会自动保存在这里</div>
            </div>
        `;
        document.getElementById('btn-batch-execute').disabled = true;
        return;
    }
    
    document.getElementById('btn-batch-execute').disabled = false;
    
    // 渲染历史记录（三列布局：查询、分析、代码）
    listEl.innerHTML = history.map(record => `
        <div class="execution-history-row" data-id="${record.id}">
            <div class="history-row-header">
                <span class="history-row-time">${escapeHtml(record.timestamp || '')}</span>
                <button class="history-row-delete" onclick="deleteExecutionHistoryItem(${record.id}, event)" title="删除此记录">×</button>
            </div>
            <div class="history-row-content">
                <div class="history-col history-col-query">
                    <div class="history-col-label">💬 查询</div>
                    <div class="history-col-value">${escapeHtml(record.query)}</div>
                </div>
                <div class="history-col history-col-analysis">
                    <div class="history-col-label">💡 分析</div>
                    <div class="history-col-value">${escapeHtml(record.analysis || '无')}</div>
                </div>
                <div class="history-col history-col-code">
                    <div class="history-col-label">📝 代码</div>
                    <pre class="history-col-code-content">${escapeHtml(record.code)}</pre>
                </div>
            </div>
        </div>
    `).join('');
};

window.closeExecutionHistoryModal = function() {
    const modal = document.getElementById('execution-history-modal');
    if (modal) modal.style.display = 'none';
};

// 删除单条执行历史
window.deleteExecutionHistoryItem = async function(recordId, event) {
    if (event) event.stopPropagation();
    
    try {
        const result = await apiRequest('DELETE', `/api/history/${recordId}`);
        if (result.success) {
            // 移除该行
            const row = document.querySelector(`.execution-history-row[data-id="${recordId}"]`);
            if (row) {
                row.style.animation = 'fadeOut 0.3s ease';
                setTimeout(() => {
                    row.remove();
                    // 检查是否还有记录
                    const listEl = document.getElementById('execution-history-list');
                    if (listEl && !listEl.querySelector('.execution-history-row')) {
                        listEl.innerHTML = `
                            <div class="empty-execution-history">
                                <div class="empty-icon">📭</div>
                                <div class="empty-text">暂无执行历史</div>
                            </div>
                        `;
                        document.getElementById('btn-batch-execute').disabled = true;
                    }
                }, 300);
            }
            showToast('已删除', 'info');
            updateHistoryUI();
        }
    } catch (e) {
        showToast('删除失败', 'error');
    }
};

// 清空全部执行历史
window.clearAllExecutionHistory = async function() {
    if (!confirm('确定要清空所有执行历史吗？')) return;
    
    try {
        const result = await apiRequest('DELETE', '/api/history');
        if (result.success) {
            const listEl = document.getElementById('execution-history-list');
            if (listEl) {
                listEl.innerHTML = `
                    <div class="empty-execution-history">
                        <div class="empty-icon">📭</div>
                        <div class="empty-text">暂无执行历史</div>
                    </div>
                `;
            }
            document.getElementById('btn-batch-execute').disabled = true;
            showToast('已清空执行历史', 'info');
            updateHistoryUI();
        }
    } catch (e) {
        showToast('清空失败', 'error');
    }
};

// 批量执行历史
window.batchExecuteHistory = async function() {
    closeExecutionHistoryModal();
    
    if (state.tables.length === 0) {
        showToast('请先上传数据表', 'warning');
        return;
    }
    
    // 添加用户消息
    addMessage('user', '历史执行处理');
    
    // 显示加载
    renderTypingIndicator();
    state.isLoading = true;
    
    try {
        const result = await apiRequest('POST', '/api/smart-execute', {
            user_description: '历史执行处理'
        });
        
        removeTypingIndicator();
        
        if (!result.success) {
            addMessage('assistant', `❌ 执行失败: ${result.error || '未知错误'}`, { status: 'error' });
            return;
        }
        
        const data = result.data;
        
        if (data.status === 'success') {
            // 构建执行日志显示
            let logHtml = '';
            if (data.historyExecution) {
                const { logs, total_steps, success_count, error_count } = data.historyExecution;
                logHtml = `<div class="history-execution-result">
                    <div class="execution-summary">
                        执行完成：共 ${total_steps} 步，成功 ${success_count}，失败 ${error_count}
                    </div>
                    <div class="execution-log-list">
                        ${logs.map(log => `
                            <div class="execution-log-item ${log.status}">
                                <span class="log-status">${log.status === 'success' ? '✅' : '❌'}</span>
                                <span class="log-step">步骤 ${log.step}</span>
                                <span class="log-query">${escapeHtml(log.query)}</span>
                                ${log.result_rows ? `<span class="log-rows">(${log.result_rows}行)</span>` : ''}
                            </div>
                        `).join('')}
                    </div>
                </div>`;
            }
            
            const analysisText = data.analysis || '';
            
            addMessage('assistant', `${analysisText}${logHtml}`, {
                isHtml: true,
                code: data.code,
                data: data.result?.data || [],
                totalRows: data.result?.total_rows,
                status: 'success',
            });
        } else if (data.status === 'need_clarification') {
            addMessage('assistant', `⚠️ ${data.message}\n\n${data.analysis || ''}`, { status: 'clarification' });
        } else {
            addMessage('assistant', `❌ ${data.message || '执行失败'}`, { status: 'error' });
        }
    } catch (error) {
        removeTypingIndicator();
        addMessage('assistant', `❌ 发生错误: ${error.message}`, { status: 'error' });
    } finally {
        state.isLoading = false;
    }
};

// 删除执行历史记录（兼容旧函数名）
window.deleteHistoryRecord = async function(recordId, event) {
    return deleteExecutionHistoryItem(recordId, event);
};

// 清空所有执行历史（兼容旧函数名）
window.clearAllHistory = async function() {
    return clearAllExecutionHistory();
};

async function updateHistoryUI() {
    // 从后端加载执行历史
    const executionHistory = await loadExecutionHistory();
    
    if (executionHistory.length === 0 && state.history.length === 0) {
        elements.historyList.innerHTML = '<div class="empty-history-hint">暂无历史记录</div>';
        return;
    }
    
    let html = '';
    
    // 显示执行历史简要信息
    if (executionHistory.length > 0) {
        html += `
            <div class="history-section">
                <div class="history-section-header">
                    <span class="history-section-title">📋 执行历史 (${executionHistory.length})</span>
                    <button class="history-view-all-btn" onclick="openExecutionHistoryModal()">查看全部</button>
                </div>
        `;
        
        html += executionHistory.slice(0, 5).map(h => `
            <div class="history-item execution-history-item" data-id="${h.id}" onclick="openExecutionHistoryModal()">
                <div class="history-text" title="${escapeHtml(h.query)}">${escapeHtml(h.query.length > 25 ? h.query.slice(0, 25) + '...' : h.query)}</div>
                <div class="history-time">${escapeHtml(h.timestamp ? h.timestamp.split(' ')[1] : '')}</div>
            </div>
        `).join('');
        
        html += '</div>';
    }
    
    // 显示本地查询历史
    if (state.history.length > 0) {
        html += `
            <div class="history-section">
                <div class="history-section-header">
                    <span class="history-section-title">🕒 查询历史</span>
                </div>
        `;
        
        html += state.history.map(h => `
            <div class="history-item local-history-item" data-text="${escapeHtml(h.text)}">
                <div class="history-text">${escapeHtml(h.text.length > 25 ? h.text.slice(0, 25) + '...' : h.text)}</div>
                <div class="history-time">${formatTime(h.time)}</div>
            </div>
        `).join('');
        
        html += '</div>';
    }
    
    elements.historyList.innerHTML = html;
    
    // 绑定点击事件（本地历史，点击填入输入框）
    elements.historyList.querySelectorAll('.local-history-item').forEach(item => {
        item.addEventListener('click', () => {
            elements.userInput.value = item.dataset.text;
            elements.userInput.focus();
        });
    });
}

// ===== 事件绑定 =====
function bindExampleTags() {
    $$('.example-tag').forEach(tag => {
        tag.addEventListener('click', () => {
            elements.userInput.value = tag.dataset.question;
            elements.userInput.focus();
        });
    });
}

function bindEvents() {
    // 窗口控制按钮
    document.getElementById('btn-minimize')?.addEventListener('click', () => {
        window.electronAPI.windowMinimize();
    });
    
    document.getElementById('btn-maximize')?.addEventListener('click', () => {
        window.electronAPI.windowMaximize();
    });
    
    document.getElementById('btn-close')?.addEventListener('click', () => {
        window.electronAPI.windowClose();
    });
    
    // 发送按钮
    elements.btnSend.addEventListener('click', sendMessage);
    
    // 回车发送，Shift+Enter 换行
    elements.userInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            sendMessage();
        }
    });
    
    // 自动调整输入框高度
    elements.userInput.addEventListener('input', () => {
        const textarea = elements.userInput;
        textarea.style.height = 'auto';
        textarea.style.height = Math.min(textarea.scrollHeight, 180) + 'px';
    });
    
    // 全局快捷键（由设置控制）
    document.addEventListener('keydown', handleGlobalShortcuts);
    
    // 快捷操作
    elements.quickActions.addEventListener('click', (e) => {
        const tag = e.target.closest('.quick-tag');
        if (!tag) return;
        
        const action = tag.dataset.action;
        switch (action) {
            case 'upload':
                uploadFile();
                break;
            case 'history':
                openExecutionHistoryModal();
                break;
            case 'clear':
                clearChat();
                break;
        }
    });
    
    // 上传按钮
    $('#btn-upload-table')?.addEventListener('click', uploadFile);
    
    // 导出报表按钮
    elements.btnExportReport?.addEventListener('click', exportReport);
    
    // 示例问题
    bindExampleTags();
    
    // 模式切换
    $$('.mode-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            $$('.mode-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
        });
    });
}

// ===== 初始化 =====
async function init() {
    console.log('AI 报表生成器前端初始化...');
    
    // 初始化主题、侧边栏和设置
    initTheme();
    initSidebar();
    loadSettings();
    
    // 清空后端表格（确保每次启动从 A 开始）
    try {
        await apiRequest('DELETE', '/api/tables');
    } catch (e) {
        console.log('初始化时清空表格失败（后端可能未就绪）');
    }
    
    // 初始化 Agent 系统
    initAgents();
    
    // 清空当前 Agent 的表格记录（与后端同步）
    state.tables = [];
    
    bindEvents();
    
    // 检查后端状态
    checkStatus();
    setInterval(checkStatus, 30000);
    
    // 检查工作区状态
    
    // 初始化 UI
    updateTablesUI();
    renderAllMessages();
    updateHistoryUI();
}

document.addEventListener('DOMContentLoaded', init);
