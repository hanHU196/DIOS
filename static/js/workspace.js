// 工作区主逻辑
let currentFileManager = null;
// ========== 功能内进度条管理（美化版） ==========
let functionProgress = {
    container: null,
    fillBar: null,
    percentText: null,
    statusText: null,
    detailText: null,
    stepsContainer: null,
    stepItems: []
};

// 初始化功能内进度条（每次切换功能时调用）
function initFunctionProgress() {
    functionProgress.container = document.getElementById('functionProgress');
    if (functionProgress.container) {
        functionProgress.fillBar = functionProgress.container.querySelector('.progress-bar-fill');
        functionProgress.percentText = document.getElementById('progressPercent');
        functionProgress.statusText = document.getElementById('progressStatus');
        functionProgress.detailText = document.getElementById('progressDetail');
        functionProgress.stepsContainer = document.getElementById('progressSteps');
        
        // 获取所有步骤项
        if (functionProgress.stepsContainer) {
            functionProgress.stepItems = functionProgress.stepsContainer.querySelectorAll('.progress-step-item');
        }
        
        // 调试日志
        console.log('进度条初始化完成', {
            container: !!functionProgress.container,
            fillBar: !!functionProgress.fillBar,
            percentText: !!functionProgress.percentText,
            statusText: !!functionProgress.statusText,
            detailText: !!functionProgress.detailText,
            stepsContainer: !!functionProgress.stepsContainer,
            stepItemsCount: functionProgress.stepItems?.length || 0
        });
    } else {
        console.error('找不到 #functionProgress 容器');
    }
}
// 显示进度条
function showFunctionProgress() {
    if (functionProgress.container) {
        functionProgress.container.style.display = 'block';
    }
}

// 更新进度条（支持步骤状态）
function updateFunctionProgress(percent, status, detail = '', stepIndex = -1) {
    if (!functionProgress.container) {
        console.error('进度条容器不存在');
        return;
    }
    
    functionProgress.container.style.display = 'block';
    
    // 更新百分比
    if (functionProgress.fillBar) {
        functionProgress.fillBar.style.width = percent + '%';
    }
    if (functionProgress.percentText) {
        functionProgress.percentText.textContent = percent + '%';
    }
    
    // 更新状态文字
    if (functionProgress.statusText) {
        functionProgress.statusText.textContent = status;
    }
    
    // 更新详情
    if (functionProgress.detailText) {
        if (detail) {
            functionProgress.detailText.textContent = detail;
            functionProgress.detailText.style.display = 'block';
        } else {
            functionProgress.detailText.style.display = 'none';
        }
    }
    
    // 更新步骤指示器
    if (stepIndex >= 0 && functionProgress.stepItems) {
        functionProgress.stepItems.forEach((step, idx) => {
            step.classList.remove('active', 'completed', 'pending');
            if (idx < stepIndex) {
                step.classList.add('completed');
            } else if (idx === stepIndex) {
                step.classList.add('active');
            } else {
                step.classList.add('pending');
            }
        });
    }
    
    console.log(`进度更新: ${percent}% - ${status}`);
}
// 隐藏进度条
function hideFunctionProgress(delay = 0) {
    if (delay > 0) {
        setTimeout(() => {
            if (functionProgress.container) {
                functionProgress.container.style.display = 'none';
            }
        }, delay);
    } else if (functionProgress.container) {
        functionProgress.container.style.display = 'none';
    }
}

// 进度条完成状态
function completeFunctionProgress(status = '处理完成', detail = '') {
    updateFunctionProgress(100, status, detail, 3);
    if (functionProgress.fillBar) {
        functionProgress.fillBar.style.background = 'var(--success)';
    }
    if (functionProgress.stepItems) {
        functionProgress.stepItems.forEach(step => {
            step.classList.remove('active', 'pending');
            step.classList.add('completed');
        });
    }
    hideFunctionProgress(2500);
}

// 进度条错误状态
function errorFunctionProgress(errorMsg) {
    updateFunctionProgress(100, '处理失败', errorMsg);
    if (functionProgress.fillBar) {
        functionProgress.fillBar.style.background = 'var(--danger)';
    }
    if (functionProgress.statusText) {
        functionProgress.statusText.style.color = 'var(--danger)';
    }
    hideFunctionProgress(3500);
}
// 重置进度条
function resetFunctionProgress() {
    if (functionProgress.fillBar) {
        functionProgress.fillBar.style.background = 'var(--accent-solid)';
        functionProgress.fillBar.style.width = '0%';
    }
    if (functionProgress.percentText) {
        functionProgress.percentText.textContent = '0%';
    }
    if (functionProgress.statusText) {
        functionProgress.statusText.style.color = 'var(--text-secondary)';
        functionProgress.statusText.textContent = '准备就绪';
    }
    if (functionProgress.detailText) {
        functionProgress.detailText.textContent = '';
        functionProgress.detailText.style.display = 'none';
    }
    if (functionProgress.stepItems) {
        functionProgress.stepItems.forEach((step) => {
            step.classList.remove('active', 'completed', 'pending');
            step.classList.add('pending');
        });
    }
    if (functionProgress.container) {
        functionProgress.container.style.display = 'none';
    }
}
// 重置进度条颜色
function resetFunctionProgressColor() {
    if (functionProgress.fillBar) {
        functionProgress.fillBar.style.background = 'var(--accent-solid)';
    }
    if (functionProgress.statusText) {
        functionProgress.statusText.style.color = 'var(--text-secondary)';
    }
}

let currentFunction = null;
let currentChart = null;

// ========== 历史文档管理（对接后端数据库） ==========
let historyDocuments = [];

// 从后端加载历史文档
async function loadHistoryFromServer() {
    try {
        const response = await fetch('/api/history/list');
        const data = await response.json();
        if (data.success) {
            historyDocuments = data.data || [];
            renderHistoryList();
            updateHistoryCount();
        } else {
            console.error('加载历史失败:', data.error);
        }
    } catch (error) {
        console.error('加载历史失败:', error);
    }
}

// 在文件开头添加标记
let isImportingFromHistory = false;

// 修改 addToHistory 函数，检查标记
async function addToHistory(file) {
    // 如果是历史导入操作，跳过保存
    if (isImportingFromHistory) {
        console.log('跳过历史导入文件的保存');
        return;
    }
    
    try {
        const formData = new FormData();
        formData.append('files', file);
        
        const response = await fetch('/api/history/add', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        if (data.success) {
            await loadHistoryFromServer();
        }
    } catch (error) {
        console.error('保存到历史失败:', error);
    }
}

// 修改 importFromHistory 函数（使用专属通道下载真实文件）
async function importFromHistory(items) {
    const ids = items.map(item => item._id);
    try {
        const response = await fetch('/api/history/import', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ids: ids })
        });
        
        const data = await response.json();
        if (data.success && data.files) {
            // 设置标记，避免导入时又被重新存一遍历史
            isImportingFromHistory = true;
            
            for (const fileInfo of data.files) {
                // 【修复核心】：去请求后端的专属下载通道，而不是去读死目录
                const fileResponse = await fetch(`/api/history/download/${fileInfo.id}`);
                
                if (!fileResponse.ok) {
                    throw new Error(`无法获取文件流: ${fileInfo.name}`);
                }
                
                // 将真实的二进制流包装回 File 对象
                const blob = await fileResponse.blob();
                const file = new File([blob], fileInfo.name, { 
                    type: fileResponse.headers.get('content-type') || 'application/octet-stream' 
                });
                
                if (currentFileManager) {
                    currentFileManager.addFiles([file]);
                }
            }
            
            // 延迟释放锁定
            setTimeout(() => {
                isImportingFromHistory = false;
            }, 500);
        }
    } catch (error) {
        console.error('导入失败:', error);
        alert('导入失败: ' + error.message);
        isImportingFromHistory = false;
    }
}
// 删除历史文档（单个或多个）
async function deleteHistoryItem(id) {
    try {
        const response = await fetch('/api/history/delete', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ ids: [id] })
        });
        
        const data = await response.json();
        if (data.success) {
            await loadHistoryFromServer();
        }
    } catch (error) {
        console.error('删除失败:', error);
        alert('删除失败: ' + error.message);
    }
}

// 清空所有历史文档
async function clearAllHistory() {
    if (!confirm('确定要清空所有历史文档吗？此操作不可恢复。')) {
        return;
    }
    
    try {
        const response = await fetch('/api/history/clear', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            }
        });
        
        const data = await response.json();
        if (data.success) {
            await loadHistoryFromServer();
        }
    } catch (error) {
        console.error('清空失败:', error);
        alert('清空失败: ' + error.message);
    }
}

// 搜索历史文档
async function searchHistory(keyword) {
    try {
        const response = await fetch('/api/history/search', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ keyword: keyword })
        });
        
        const data = await response.json();
        if (data.success) {
            historyDocuments = data.data || [];
            renderHistoryList();
            updateHistoryCount();
        }
    } catch (error) {
        console.error('搜索失败:', error);
    }
}

// 更新历史文档数量显示
function updateHistoryCount() {
    const countSpan = document.getElementById('historyCount');
    if (countSpan) {
        countSpan.textContent = historyDocuments.length;
    }
}

// 格式化文件大小
function formatFileSizeForHistory(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 格式化日期
function formatDate(timestamp) {
    if (!timestamp) return '未知';
    const date = new Date(timestamp);
    const now = new Date();
    const diff = now - date;
    
    if (diff < 60 * 1000) return '刚刚';
    if (diff < 60 * 60 * 1000) return `${Math.floor(diff / (60 * 1000))}分钟前`;
    if (diff < 24 * 60 * 60 * 1000) return `${Math.floor(diff / (60 * 60 * 1000))}小时前`;
    if (diff < 7 * 24 * 60 * 60 * 1000) return `${Math.floor(diff / (24 * 60 * 60 * 1000))}天前`;
    return date.toLocaleDateString();
}

// 获取文件图标
function getHistoryFileIcon(fileName) {
    const ext = fileName.split('.').pop().toLowerCase();
    const icons = {
        txt: 'file-text',
        docx: 'file-text',
        md: 'file-text',
        xlsx: 'table',
        xls: 'table',
        pdf: 'file',
        jpg: 'image',
        png: 'image',
        zip: 'archive'
    };
    return icons[ext] || 'file';
}

// 渲染历史文档列表
function renderHistoryList() {
    const historyList = document.getElementById('historyList');
    const historyFooter = document.getElementById('historyFooter');
    
    if (!historyList) return;
    
    if (historyDocuments.length === 0) {
        historyList.innerHTML = `
            <div class="history-empty">
                <i data-lucide="archive"></i>
                <p>暂无历史文档</p>
                <small>上传文档后会自动保存到历史记录</small>
            </div>
        `;
        if (historyFooter) historyFooter.style.display = 'none';
        lucide.createIcons();
        return;
    }
    
    if (historyFooter) historyFooter.style.display = 'flex';
    
    historyList.innerHTML = '';
    historyDocuments.forEach(doc => {
        const item = document.createElement('div');
        item.className = 'history-item';
        item.innerHTML = `
            <div class="history-item-checkbox">
                <input type="checkbox" class="history-checkbox" data-id="${doc._id}">
            </div>
            <div class="history-item-info">
                <div class="history-item-icon">
                    <i data-lucide="${getHistoryFileIcon(doc.name)}" style="width: 24px; height: 24px;"></i>
                </div>
                <div class="history-item-details">
                    <div class="history-item-name" title="${doc.name}">${doc.name}</div>
                    <div class="history-item-meta">
                        <span>${formatFileSizeForHistory(doc.size)}</span>
                        <span>${formatDate(doc.timestamp)}</span>
                    </div>
                </div>
            </div>
            <div class="history-item-actions">
                <button class="import-btn" data-id="${doc._id}" title="导入">
                    <i data-lucide="import"></i>
                </button>
                <button class="delete-btn" data-id="${doc._id}" title="删除">
                    <i data-lucide="trash-2"></i>
                </button>
            </div>
        `;
        historyList.appendChild(item);
    });
    
    lucide.createIcons();
    
    // 绑定复选框事件
    document.querySelectorAll('.history-checkbox').forEach(cb => {
        cb.addEventListener('change', updateSelectedCount);
    });
    
    // 绑定导入按钮事件
    document.querySelectorAll('.import-btn').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            const id = btn.dataset.id;
            const doc = historyDocuments.find(d => d._id === id);
            if (doc) {
                await importFromHistory([doc]);
            }
        });
    });
    
    // 绑定删除按钮事件
    document.querySelectorAll('.delete-btn').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            const id = btn.dataset.id;
            await deleteHistoryItem(id);
        });
    });
}

// 更新选中数量
function updateSelectedCount() {
    const checkboxes = document.querySelectorAll('.history-checkbox:checked');
    const selectedCountSpan = document.getElementById('selectedCount');
    if (selectedCountSpan) {
        selectedCountSpan.textContent = checkboxes.length;
    }
}

// 获取选中的历史文档
function getSelectedHistoryItems() {
    const checkboxes = document.querySelectorAll('.history-checkbox:checked');
    const selectedIds = Array.from(checkboxes).map(cb => cb.dataset.id);
    return historyDocuments.filter(doc => selectedIds.includes(doc._id));
}

// 全选/取消全选
function toggleSelectAll() {
    const selectAllCheckbox = document.getElementById('selectAllHistory');
    const checkboxes = document.querySelectorAll('.history-checkbox');
    checkboxes.forEach(cb => {
        cb.checked = selectAllCheckbox.checked;
    });
    updateSelectedCount();
}

// 导入选中的文档
async function importSelected() {
    const selected = getSelectedHistoryItems();
    if (selected.length === 0) {
        alert('请先选择要导入的文档');
        return;
    }
    await importFromHistory(selected);
    // 清空复选框选中状态
    document.querySelectorAll('.history-checkbox').forEach(cb => {
        cb.checked = false;
    });
    updateSelectedCount();
    const selectAll = document.getElementById('selectAllHistory');
    if (selectAll) selectAll.checked = false;
}

// 初始化历史文档模块
async function initHistoryModule() {
    await loadHistoryFromServer();
    
    // 绑定全局事件
    const refreshBtn = document.getElementById('refreshHistoryBtn');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', () => loadHistoryFromServer());
    }
    
    const clearAllBtn = document.getElementById('clearAllHistoryBtn');
    if (clearAllBtn) {
        clearAllBtn.addEventListener('click', () => clearAllHistory());
    }
    
    const selectAllCheckbox = document.getElementById('selectAllHistory');
    if (selectAllCheckbox) {
        selectAllCheckbox.addEventListener('change', toggleSelectAll);
    }
    
    const importBtn = document.getElementById('importSelectedBtn');
    if (importBtn) {
        importBtn.addEventListener('click', importSelected);
    }
}

// ========== 历史文档侧边栏控制 ==========
function openHistorySidebar() {
    const sidebar = document.getElementById('historySidebar');
    const overlay = document.getElementById('historyOverlay');
    if (sidebar) sidebar.classList.add('open');
    if (overlay) overlay.classList.add('open');
    // 打开时刷新列表
    loadHistoryFromServer();
}

function closeHistorySidebar() {
    const sidebar = document.getElementById('historySidebar');
    const overlay = document.getElementById('historyOverlay');
    if (sidebar) sidebar.classList.remove('open');
    if (overlay) overlay.classList.remove('open');
}

// 绑定侧边栏事件
function bindHistoryEvents() {
    const toggleBtn = document.getElementById('historyToggleBtn');
    const closeBtn = document.getElementById('historySidebarClose');
    const overlay = document.getElementById('historyOverlay');
    const searchInput = document.getElementById('historySearchInput');
    
    if (toggleBtn) {
        toggleBtn.addEventListener('click', openHistorySidebar);
    }
    if (closeBtn) {
        closeBtn.addEventListener('click', closeHistorySidebar);
    }
    if (overlay) {
        overlay.addEventListener('click', closeHistorySidebar);
    }
    if (searchInput) {
        let debounceTimer;
        searchInput.addEventListener('input', (e) => {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(() => {
                const keyword = e.target.value.trim();
                if (keyword) {
                    searchHistory(keyword);
                } else {
                    loadHistoryFromServer();
                }
            }, 300);
        });
    }
}

// 删除选中的历史文档
async function deleteSelectedHistory() {
    const checkboxes = document.querySelectorAll('.history-checkbox:checked');
    if (checkboxes.length === 0) {
        alert('请先选择要删除的文档');
        return;
    }
    
    if (!confirm(`确定要删除选中的 ${checkboxes.length} 个文档吗？`)) {
        return;
    }
    
    const ids = Array.from(checkboxes).map(cb => cb.dataset.id);
    try {
        const response = await fetch('/api/history/delete', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ids: ids })
        });
        
        const data = await response.json();
        if (data.success) {
            await loadHistoryFromServer();
            // 清空全选状态
            const selectAll = document.getElementById('selectAllHistory');
            if (selectAll) selectAll.checked = false;
        }
    } catch (error) {
        console.error('删除失败:', error);
        alert('删除失败: ' + error.message);
    }
}

// 修改 renderHistoryList 函数，适配侧边栏
function renderHistoryList() {
    const historyList = document.getElementById('historyList');
    const historyFooter = document.getElementById('historyFooter');
    const historyBadge = document.getElementById('historyBadge');
    const historyCount = document.getElementById('historyCount');
    
    if (!historyList) return;
    
    // 更新徽章数量
    if (historyBadge) historyBadge.textContent = historyDocuments.length;
    if (historyCount) historyCount.textContent = historyDocuments.length;
    
    if (historyDocuments.length === 0) {
        historyList.innerHTML = `
            <div class="history-empty">
                <i data-lucide="archive"></i>
                <p>暂无历史文档</p>
                <small>上传文档后会自动保存</small>
            </div>
        `;
        if (historyFooter) historyFooter.style.display = 'none';
        lucide.createIcons();
        return;
    }
    
    if (historyFooter) historyFooter.style.display = 'flex';
    
    historyList.innerHTML = '';
    historyDocuments.forEach(doc => {
        const item = document.createElement('div');
        item.className = 'history-item';
        item.innerHTML = `
            <div class="history-item-checkbox">
                <input type="checkbox" class="history-checkbox" data-id="${doc._id}">
            </div>
            <div class="history-item-icon">
                <i data-lucide="${getHistoryFileIcon(doc.name)}"></i>
            </div>
            <div class="history-item-info">
                <div class="history-item-name" title="${doc.name}">${doc.name}</div>
                <div class="history-item-meta">
                    <span>${formatFileSizeForHistory(doc.size)}</span>
                    <span>${formatDate(doc.timestamp)}</span>
                </div>
            </div>
            <div class="history-item-actions">
                <button class="import-single-btn" data-id="${doc._id}" title="导入">
                    <i data-lucide="import"></i>
                </button>
                <button class="delete-single-btn" data-id="${doc._id}" title="删除">
                    <i data-lucide="trash-2"></i>
                </button>
            </div>
        `;
        historyList.appendChild(item);
    });
    
    lucide.createIcons();
    
    // 绑定复选框事件
    document.querySelectorAll('.history-checkbox').forEach(cb => {
        cb.addEventListener('change', updateSelectedCount);
    });
    
    // 绑定单个导入按钮
    document.querySelectorAll('.import-single-btn').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            e.stopPropagation();
            const id = btn.dataset.id;
            const doc = historyDocuments.find(d => d._id === id);
            if (doc) {
                await importFromHistory([doc]);
            }
        });
    });
    
    // 绑定单个删除按钮
    document.querySelectorAll('.delete-single-btn').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            e.stopPropagation();
            const id = btn.dataset.id;
            await deleteHistoryItem(id);
        });
    });
}

// 更新选中数量
function updateSelectedCount() {
    const checkboxes = document.querySelectorAll('.history-checkbox:checked');
    const selectedCountSpan = document.getElementById('selectedCount');
    if (selectedCountSpan) {
        selectedCountSpan.textContent = checkboxes.length;
    }
}

// 全选/取消全选
function toggleSelectAll() {
    const selectAllCheckbox = document.getElementById('selectAllHistory');
    if (!selectAllCheckbox) return;
    
    const checkboxes = document.querySelectorAll('.history-checkbox');
    checkboxes.forEach(cb => {
        cb.checked = selectAllCheckbox.checked;
    });
    updateSelectedCount();
}

// 导入选中的文档
async function importSelected() {
    const selected = getSelectedHistoryItems();
    if (selected.length === 0) {
        alert('请先选择要导入的文档');
        return;
    }
    await importFromHistory(selected);
    // 清空复选框选中状态
    document.querySelectorAll('.history-checkbox').forEach(cb => {
        cb.checked = false;
    });
    updateSelectedCount();
    const selectAll = document.getElementById('selectAllHistory');
    if (selectAll) selectAll.checked = false;
}

// 初始化历史文档模块
async function initHistoryModule() {
    await loadHistoryFromServer();
    bindHistoryEvents();
    
    // 绑定全选事件
    const selectAllCheckbox = document.getElementById('selectAllHistory');
    if (selectAllCheckbox) {
        selectAllCheckbox.addEventListener('change', toggleSelectAll);
    }
    
    // 绑定导入选中按钮
    const importBtn = document.getElementById('importSelectedBtn');
    if (importBtn) {
        importBtn.addEventListener('click', importSelected);
    }
    
    // 绑定删除选中按钮
    const deleteBtn = document.getElementById('deleteSelectedBtn');
    if (deleteBtn) {
        deleteBtn.addEventListener('click', deleteSelectedHistory);
    }
}
// 页面跳转
document.addEventListener('DOMContentLoaded', () => {
    const backHomeBtn = document.getElementById('backHomeBtn');
    if (backHomeBtn) {
        backHomeBtn.addEventListener('click', () => {
            window.location.href = '/';
        });
    }
    
    initFileManager();
    initHistoryModule();  // 新增
    bindFunctionCards();
});

// 文件管理器初始化
function initFileManager() {
    currentFileManager = new FileManager();
    currentFileManager.init({
        fileInputId: 'globalFileInput',
        listContainerId: 'globalFileList',
        listBodyId: 'globalFileListBody',
        countSpanId: 'globalFileCount',
        onChange: (files) => {
            // 如果是历史导入操作，跳过自动保存
            if (isImportingFromHistory) {
                console.log('跳过历史导入文件的自动保存');
                if (currentFunction && files.length > 0) {
                    loadFunctionUI(currentFunction);
                }
                return;
            }
            
            // 正常上传：自动保存到历史记录
            files.forEach(file => {
                addToHistory(file);
            });
            
            if (currentFunction && files.length > 0) {
                loadFunctionUI(currentFunction);
            } else if (files.length === 0) {
                document.getElementById('resultContent').innerHTML = `
                    <div style="text-align: center; padding: 60px; color: var(--warning);">
                        <i data-lucide="alert-triangle" style="width: 32px; height: 32px; margin-bottom: 12px;"></i><br>
                        请先上传文档
                    </div>
                `;
                lucide.createIcons();
            }
        }
    });
    currentFileManager.setupDragDrop('workspaceDropArea');
}
// 覆盖 FileManager 的 render 方法，使用 Lucide 线性图标
function overrideFileManagerRender() {
    if (!currentFileManager) return;
    
    const originalRender = currentFileManager.render.bind(currentFileManager);
    currentFileManager.render = function() {
        if (!this.fileListContainer || !this.fileListBody) return;
        
        if (this.files.length === 0) {
            this.fileListContainer.classList.add('hidden');
            return;
        }
        
        this.fileListContainer.classList.remove('hidden');
        if (this.countSpan) {
            this.countSpan.textContent = this.files.length;
        }
        
        this.fileListBody.innerHTML = '';
        this.files.forEach((file, idx) => {
            const item = document.createElement('div');
            item.className = 'file-item';
            // 使用 Utils.getFileIcon 获取 Lucide 图标
            item.innerHTML = `
                <div class="file-info">
                    ${Utils.getFileIcon(file.name)}
                    <div>
                        <div style="font-size: 0.85rem; font-weight: 500;">${file.name}</div>
                        <div style="font-size: 0.7rem; color: var(--text-placeholder);">${Utils.formatFileSize(file.size)}</div>
                    </div>
                </div>
                <button class="file-remove" data-index="${idx}">
                    <i data-lucide="x" style="width: 14px; height: 14px;"></i>
                </button>
            `;
            this.fileListBody.appendChild(item);
        });
        
        // 重新初始化 Lucide 图标
        if (typeof lucide !== 'undefined') {
            lucide.createIcons();
        }
        
        this.fileListBody.querySelectorAll('.file-remove').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = parseInt(btn.dataset.index);
                this.removeFile(idx);
            });
        });
    };
    
    // 重新渲染一次
    currentFileManager.render();
}

// 绑定功能卡片
function bindFunctionCards() {
    document.querySelectorAll('.function-card').forEach(card => {
        card.addEventListener('click', () => {
            document.querySelectorAll('.function-card').forEach(c => c.classList.remove('active'));
            card.classList.add('active');
            loadFunctionUI(card.dataset.function);
        });
    });
}

// 加载功能 UI
function loadFunctionUI(func) {
    currentFunction = func;
    const resultContent = document.getElementById('resultContent');
    const resultTitle = document.getElementById('resultTitle');
    
    const titles = {
        extract: '智能文档提取',
        fill: '表格智能填写',
        format: '文档格式调整',
        qa: '智能问答',
        graph: '知识图谱',
        analyze: '文档深度分析'
    };
    resultTitle.innerHTML = `<i data-lucide="${getTitleIcon(func)}" style="width: 18px; height: 18px; margin-right: 6px;"></i> ${titles[func] || '执行结果'}`;
    lucide.createIcons();
    
    const files = currentFileManager.getFiles();
    if (files.length === 0) {
        resultContent.innerHTML = '<div style="text-align: center; padding: 60px; color: var(--warning);"><i data-lucide="alert-triangle" style="width: 32px; height: 32px; margin-bottom: 12px;"></i><br>请先上传文档</div>';
        lucide.createIcons();
        return;
    }
    
    // 根据功能加载对应 UI
    switch(func) {
        case 'extract': renderExtractUI(resultContent); break;
        case 'fill': renderFillUI(resultContent); break;
        case 'format': renderFormatUI(resultContent); break;
        case 'qa': renderQaUI(resultContent); break;
        case 'graph': renderGraphUI(resultContent); break;
        case 'analyze': renderAnalyzeUI(resultContent); break;
    }
}

function getTitleIcon(func) {
    const icons = {
        extract: 'file-search',
        fill: 'table',
        format: 'paintbrush',
        qa: 'message-circle',
        graph: 'network',
        analyze: 'bar-chart-3'
    };
    return icons[func] || 'file-text';
}

// 智能提取 UI
function renderExtractUI(container) {
    container.innerHTML = `
           <!-- 美化版进度条 -->
        <div id="functionProgress" class="progress-container-enhanced" style="display: none;">
            <div class="progress-header">
                <div class="progress-title">
                    <i data-lucide="loader" class="progress-icon spin"></i>
                    <span id="progressStatus">准备就绪</span>
                </div>
                <div class="progress-percent-wrapper">
                    <span id="progressPercent">0%</span>
                </div>
            </div>
            
            <div class="progress-bar-wrapper">
                <div class="progress-bar-bg">
                    <div class="progress-bar-fill" style="width: 0%"></div>
                </div>
            </div>
            
            <div id="progressDetail" class="progress-detail" style="display: none;"></div>
            
            <div id="progressSteps" class="progress-steps">
                <div class="progress-step-item">
                    <span class="step-icon">1</span>
                    <span class="step-label">解析指令</span>
                </div>
                <div class="progress-step-item">
                    <span class="step-icon">2</span>
                    <span class="step-label">AI 提取</span>
                </div>
                <div class="progress-step-item">
                    <span class="step-icon">3</span>
                    <span class="step-label">生成结果</span>
                </div>
            </div>
        </div>
        
        <div class="form-group">
            <label class="form-label">提取指令</label>
            <input type="text" id="extractCommand" class="input" placeholder="例如：甲方、乙方、合同金额">
            <div style="font-size: 12px; color: var(--text-placeholder); margin-top: 6px;">
                <i data-lucide="lightbulb" style="width: 12px; height: 12px; display: inline;"></i> 支持格式：提取甲方、乙方、金额
            </div>
        </div>
        <button class="btn-primary" id="executeExtract">
            开始提取
        </button>
        <div id="extractResultArea" style="margin-top: 20px;"></div>
        <div id="extractError" class="error-message hidden"></div>
    `;
    lucide.createIcons();
    
    // 初始化进度条
    initFunctionProgress();
    resetFunctionProgressColor();
    
    document.getElementById('executeExtract').addEventListener('click', async () => {
        const command = document.getElementById('extractCommand').value.trim();
        if (!command) {
            Utils.showError('extractError', '请输入提取指令');
            return;
        }
        
        const area = document.getElementById('extractResultArea');
        area.innerHTML = '';
        const errorEl = document.getElementById('extractError');
        errorEl.classList.add('hidden');
        
        // 显示进度条
    resetFunctionProgressColor();
    showFunctionProgress();

    // 使用 stepIndex 参数（0, 1, 2 分别对应三个步骤）
    updateFunctionProgress(10, '解析指令中', '正在识别提取字段...', 0);

    // 在 fetch 之前
    updateFunctionProgress(30, '读取文档', `正在处理 ${files.length} 个文档...`, 0);

    // fetch 请求时
    updateFunctionProgress(60, 'AI 提取中', '正在识别关键信息...', 1);

    // 处理完成时
    updateFunctionProgress(85, '生成结果', `已提取 ${records.length} 条记录`, 2);

    // 完成时
    completeFunctionProgress('提取完成', `成功提取 ${records.length} 条记录`);
        
        const files = currentFileManager.getFiles();
        const formData = new FormData();
        for (const file of files) {
            formData.append('documents', file);
        }
        formData.append('command', command);
        
        try {
            updateFunctionProgress(30, '正在处理', '读取文档内容...');
            
            const response = await fetch('/api/extract', {
                method: 'POST',
                body: formData
            });
            
            updateFunctionProgress(70, 'AI 提取中', '正在识别关键信息...');
            
            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || '提取失败');
            }
            
            const data = await response.json();
            const fields = data.fields || [];
            const records = data.data || [];
            
            updateFunctionProgress(90, '生成结果', '渲染数据表格...');
            
            if (records.length === 0) {
                area.innerHTML = '<div style="text-align: center; padding: 40px; color: var(--warning);"><i data-lucide="inbox" style="width: 32px; height: 32px; margin-bottom: 12px;"></i><br>未提取到数据，请检查文档内容或指令</div>';
                lucide.createIcons();
                completeFunctionProgress('完成', '未提取到数据');
                return;
            }
            
            let html = `<table class="data-table"><thead><tr>${fields.map(f => `<th>${f}</th>`).join('')}</tr></thead><tbody>`;
            records.forEach(r => {
                html += '<tr>' + fields.map(f => `<td contenteditable="true">${r[f] || '<span style="color: var(--text-placeholder);">无数据</span>'}</td>`).join('') + '</tr>';
            });
            html += `</tbody></table>
                <div style="margin-top: 16px;">
                    <button class="btn-secondary" id="exportTableBtn">
                        <i data-lucide="download" style="width: 14px; height: 14px;"></i> 导出 CSV
                    </button>
                </div>`;
            area.innerHTML = html;
            lucide.createIcons();
            
            completeFunctionProgress('提取完成', `成功提取 ${records.length} 条记录`);
            
            document.getElementById('exportTableBtn')?.addEventListener('click', () => {
                Utils.exportToCSV(records, fields, '提取结果');
            });
            
        } catch (error) {
            console.error('提取失败:', error);
            errorFunctionProgress(error.message);
            Utils.showError('extractError', error.message);
            area.innerHTML = `<div class="error-message"><i data-lucide="alert-circle" style="width: 16px; height: 16px;"></i> ${error.message}</div>`;
            lucide.createIcons();
        }
    });
}

// 表格填写 UI（支持 Word 和 Excel 模板，自动识别）
function renderFillUI(container) {
    
    container.innerHTML = `
    <!-- 美化版进度条 -->
        <div id="functionProgress" class="progress-container-enhanced" style="display: none;">
            <div class="progress-header">
                <div class="progress-title">
                    <i data-lucide="loader" class="progress-icon spin"></i>
                    <span id="progressStatus">准备就绪</span>
                </div>
                <div class="progress-percent-wrapper">
                    <span id="progressPercent">0%</span>
                </div>
            </div>
            
            <div class="progress-bar-wrapper">
                <div class="progress-bar-bg">
                    <div class="progress-bar-fill" style="width: 0%"></div>
                </div>
            </div>
            
            <div id="progressDetail" class="progress-detail" style="display: none;"></div>
            
            <div id="progressSteps" class="progress-steps">
                <div class="progress-step-item">
                    <span class="step-icon">1</span>
                    <span class="step-label">解析指令</span>
                </div>
                <div class="progress-step-item">
                    <span class="step-icon">2</span>
                    <span class="step-label">AI 提取</span>
                </div>
                <div class="progress-step-item">
                    <span class="step-icon">3</span>
                    <span class="step-label">生成文档</span>
                </div>
            </div>
        </div>

        <div class="form-group">
            <label class="form-label">上传模板文件</label>
            <div class="template-file-area" id="templateFileArea">
                <input type="file" id="templateFile" accept=".xlsx,.xls,.docx" class="file-input">
                <div class="template-file-placeholder" id="templateFilePlaceholder">
                    <i data-lucide="file-up" style="width: 32px; height: 32px;"></i>
                    <span>点击或拖拽模板文件至此</span>
                    <small>支持 .xlsx, .xls, .docx 格式</small>
                </div>
            </div>
            <div id="templateFileInfo" class="template-file-info hidden">
                <i data-lucide="file-check"></i>
                <span id="templateFileName">未选择文件</span>
                <button class="template-remove-btn" id="removeTemplateBtn">
                    <i data-lucide="x"></i>
                </button>
            </div>
        </div>
        <div class="form-group">
            <label class="form-label">字段指令</label>
            <input type="text" id="fillCommand" class="input" placeholder="甲方、乙方、金额">
            <div style="font-size: 12px; color: var(--text-placeholder); margin-top: 6px;">
                <i data-lucide="lightbulb" style="width: 12px; height: 12px;"></i> 支持：甲方、乙方、金额等字段
            </div>
        </div>
        <button class="btn-primary" id="executeFill">
             生成并下载
        </button>
        <div id="fillResultArea" style="margin-top: 20px;"></div>
        <div id="fillError" class="error-message hidden"></div>
    `;
    
    lucide.createIcons();
    // 初始化进度条
    initFunctionProgress();
    resetFunctionProgressColor();

    
    let currentTemplateFile = null;
    
    const templateInput = document.getElementById('templateFile');
    const templatePlaceholder = document.getElementById('templateFilePlaceholder');
    const templateFileInfo = document.getElementById('templateFileInfo');
    const templateFileName = document.getElementById('templateFileName');
    
    // 修复：完整的清空函数
    function clearTemplateFile() {
        console.log('清除模板文件'); // 调试日志
        currentTemplateFile = null;
        templateInput.value = '';
        
        // 关键修复：确保 UI 正确更新
        templateFileInfo.classList.add('hidden');
        templatePlaceholder.classList.remove('hidden');
        templateFileName.textContent = '未选择文件';
        
        // 强制重新渲染图标
        if (typeof lucide !== 'undefined') {
            lucide.createIcons();
        }
    }
    
    function updateTemplateDisplay(file) {
        currentTemplateFile = file;
        templateFileName.textContent = file.name;
        templatePlaceholder.classList.add('hidden');
        templateFileInfo.classList.remove('hidden');
    }
    
    // 文件选择
    templateInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            if (!file.name.match(/\.(xlsx|xls|docx)$/i)) {
                Utils.showError('fillError', '请上传 Excel (.xlsx, .xls) 或 Word (.docx) 文件');
                templateInput.value = '';
                return;
            }
            updateTemplateDisplay(file);
        }
    });
    
    // ========== 删除按钮：直接绑定事件 ==========
    setTimeout(() => {
        const removeBtn = document.getElementById('removeTemplateBtn');
        if (removeBtn) {
            // 移除所有已有事件，重新绑定
            const newRemoveBtn = removeBtn.cloneNode(true);
            removeBtn.parentNode.replaceChild(newRemoveBtn, removeBtn);
            newRemoveBtn.onclick = function(e) {
                e.preventDefault();
                e.stopPropagation();
                clearTemplateFile();
                return false;
            };
        }
    }, 50);
    
    // 拖拽上传模板
    const templateArea = document.getElementById('templateFileArea');
    // 移除旧的事件监听器（避免重复绑定）
    const newTemplateArea = templateArea.cloneNode(true);
    templateArea.parentNode.replaceChild(newTemplateArea, templateArea);
    
    newTemplateArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        newTemplateArea.style.borderColor = 'var(--accent-solid)';
        newTemplateArea.style.backgroundColor = 'var(--accent-soft)';
    });
    newTemplateArea.addEventListener('dragleave', () => {
        newTemplateArea.style.borderColor = '';
        newTemplateArea.style.backgroundColor = '';
    });
    newTemplateArea.addEventListener('drop', (e) => {
        e.preventDefault();
        newTemplateArea.style.borderColor = '';
        newTemplateArea.style.backgroundColor = '';
        const files = Array.from(e.dataTransfer.files);
        if (files.length > 0) {
            const file = files[0];
            if (!file.name.match(/\.(xlsx|xls|docx)$/i)) {
                Utils.showError('fillError', '请上传 Excel (.xlsx, .xls) 或 Word (.docx) 文件');
                return;
            }
            const dt = new DataTransfer();
            dt.items.add(file);
            templateInput.files = dt.files;
            updateTemplateDisplay(file);
        }
    });
    newTemplateArea.addEventListener('click', () => {
        templateInput.click();
    });
    
    // 执行填表
    const executeBtn = document.getElementById('executeFill');
    if (executeBtn) {
        const newExecuteBtn = executeBtn.cloneNode(true);
        executeBtn.parentNode.replaceChild(newExecuteBtn, executeBtn);
        newExecuteBtn.addEventListener('click', async () => {
            const command = document.getElementById('fillCommand').value.trim();
            
            if (!command) { 
                Utils.showError('fillError', '请输入字段指令'); 
                return; 
            }
            if (!currentTemplateFile) { 
                Utils.showError('fillError', '请上传模板文件'); 
                return; 
            }
            
            const files = currentFileManager.getFiles();
            if (files.length === 0) { 
                Utils.showError('fillError', '请先上传源文档'); 
                return; 
            }
            
            const area = document.getElementById('fillResultArea');
            const errorEl = document.getElementById('fillError');
            errorEl.classList.add('hidden');
            
            // 重置并显示进度条
            resetFunctionProgressColor();
            updateFunctionProgress(20, '正在准备', '解析填表指令...');
            
            const isWordTemplate = currentTemplateFile.name.match(/\.(docx)$/i);
            const templateType = isWordTemplate ? 'word' : 'excel';
            
            const formData = new FormData();
            for (const file of files) {
                formData.append('documents', file);
            }
            formData.append('template', currentTemplateFile);
            formData.append('command', command);
            formData.append('template_type', templateType);
            
            try {
                updateFunctionProgress(50, '处理中', 'AI 正在提取数据...');
                
                const response = await fetch('/api/fill', {
                    method: 'POST',
                    body: formData
                });
                
                updateFunctionProgress(80, '生成文档', '正在生成填表结果...');
                
                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.error || '生成失败');
                }
                
                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                const extension = isWordTemplate ? '.docx' : '.xlsx';
                
                completeFunctionProgress('生成完成', '文档已生成，正在下载...');
                
                area.innerHTML = `
                    <div style="text-align: center; padding: 20px;">
                        <a href="${url}" class="btn-secondary" download="filled_template${extension}" style="display: inline-flex; align-items: center; gap: 8px; padding: 10px 24px; text-decoration: none; border-radius: 40px;">
                            <i data-lucide="download" style="width: 14px; height: 14px;"></i> 下载${isWordTemplate ? '文档' : '表格'}文件
                        </a>
                    </div>
                `;
                lucide.createIcons();
                
            } catch (error) {
                console.error('填表失败:', error);
                errorFunctionProgress(error.message);
                Utils.showError('fillError', error.message);
                area.innerHTML = `<div class="error-message"><i data-lucide="alert-circle" style="width: 16px; height: 16px;"></i> ${error.message}</div>`;
                lucide.createIcons();
            }
        });
    }
}
// ========== 在 renderFillUI 函数外面，绑定全局事件委托（只绑定一次） ==========
// 使用全局事件委托处理删除按钮
if (!window._fillDeleteHandlerBound) {
    document.body.addEventListener('click', (e) => {
        // 检查点击的是否是模板删除按钮
        const deleteBtn = e.target.closest('#removeTemplateBtn');
        if (deleteBtn) {
            e.preventDefault();
            e.stopPropagation();
            
            // 获取当前激活的功能页面中的元素
            const templateInput = document.getElementById('templateFile');
            const templatePlaceholder = document.getElementById('templateFilePlaceholder');
            const templateFileInfo = document.getElementById('templateFileInfo');
            
            if (templateInput && templatePlaceholder && templateFileInfo) {
                templateInput.value = '';
                templatePlaceholder.classList.remove('hidden');
                templateFileInfo.classList.add('hidden');
            }
        }
    });
    window._fillDeleteHandlerBound = true;
}

// 格式调整 UI（根据文件类型动态显示）
function renderFormatUI(container) {
    const files = currentFileManager.getFiles();
    const fileCount = files.length;
    
    const hasWord = files.some(f => f.name.match(/\.(docx|doc)$/i));
    const hasExcel = files.some(f => f.name.match(/\.(xlsx|xls)$/i));
    const hasMixed = hasWord && hasExcel;
    
    let activePanel = 'word';
    if (hasExcel && !hasWord) {
        activePanel = 'excel';
    } else if (hasMixed) {
        activePanel = 'mixed';
    }
    
    container.innerHTML = `
        <div class="format-info-bar">
            <div class="format-doc-info">
                <i data-lucide="file-text"></i>
                <span>已选择 <strong>${fileCount}</strong> 个文档</span>
                ${fileCount > 0 && fileCount <= 3 ? `<span class="format-doc-names">${files.map(f => f.name).join(', ')}</span>` : ''}
            </div>
            ${hasMixed ? '<div class="format-warning">检测到混合文件类型，请选择要操作的文档类型</div>' : ''}
            ${fileCount === 0 ? '<div class="format-warning"><i data-lucide="alert-triangle"></i> 请先在上方上传文档</div>' : ''}
        </div>
        
        <!-- 文件类型切换标签 -->
        <div class="format-type-tabs" id="formatTypeTabs" style="${hasMixed ? 'display: flex' : 'display: none'}">
            <button class="format-tab ${activePanel === 'word' ? 'active' : ''}" data-type="word">
                Word 文档
            </button>
            <button class="format-tab ${activePanel === 'excel' ? 'active' : ''}" data-type="excel">
                Excel 表格
            </button>
        </div>
        
        <!-- Word 操作面板（包含应用范围） -->
        <div id="wordPanel" class="format-panel" style="${activePanel === 'word' ? 'display: block' : 'display: none'}">
            ${renderWordRangePanel()}
            ${renderWordFormatPanel()}
        </div>
        
        <!-- Excel 操作面板（包含应用范围） -->
        <div id="excelPanel" class="format-panel" style="${activePanel === 'excel' ? 'display: block' : 'display: none'}">
            ${renderExcelRangePanel()}
            ${renderExcelFormatPanel()}
        </div>
        
        <!-- 自定义指令区域 -->
        <div class="format-custom-section">
            <div class="format-custom-header">
                <span>自定义指令</span>
                <span class="custom-hint">（可选，会覆盖上面的选择）</span>
            </div>
            <textarea id="formatCommand" class="format-custom-input" rows="2" 
                placeholder='Word示例：第2段居中、加粗   Excel示例：第2行加粗、第3列居中'></textarea>
            <div class="format-custom-example">
                <span id="formatHint">Word：第2段居中、加粗 | Excel：第2行加粗、第3列居中、A1单元格红色</span>
            </div>
        </div>
        
        <!-- 指令预览 -->
        <div class="format-command-preview" id="formatCommandPreview">
            <span>生成的指令：</span>
            <code id="commandPreviewText">-</code>
        </div>
        
        <button class="btn-primary" id="executeFormat" ${fileCount === 0 ? 'disabled' : ''}>
            应用格式
        </button>
        <div id="formatProgressArea" style="margin-top: 20px;"></div>
        <div id="formatResultArea" style="margin-top: 20px;"></div>
        <div id="formatError" class="error-message hidden"></div>
    `;
    
    lucide.createIcons();
    
    let currentType = activePanel === 'excel' ? 'excel' : 'word';
    
    /// Word 范围切换逻辑
    const wordRangeRadios = document.querySelectorAll('input[name="wordRangeType"]');
    const wordParagraphGroup = document.getElementById('wordParagraphGroup');
    const wordKeywordGroup = document.getElementById('wordKeywordGroup');

    if (wordRangeRadios.length) {
        wordRangeRadios.forEach(radio => {
            radio.addEventListener('change', (e) => {
                const val = e.target.value;
                wordParagraphGroup.style.display = 'none';
                wordKeywordGroup.style.display = 'none';
                
                if (val === 'paragraph') wordParagraphGroup.style.display = 'block';
                else if (val === 'keyword') wordKeywordGroup.style.display = 'block';
                
                updateCommandPreview();
            });
        });
    }

    // Excel 范围切换逻辑
    const excelRangeRadios = document.querySelectorAll('input[name="excelRangeType"]');
    const excelRowGroup = document.getElementById('excelRowGroup');
    const excelColGroup = document.getElementById('excelColGroup');
    const excelCellGroup = document.getElementById('excelCellGroup');

    if (excelRangeRadios.length) {
        excelRangeRadios.forEach(radio => {
            radio.addEventListener('change', (e) => {
                const val = e.target.value;
                excelRowGroup.style.display = 'none';
                excelColGroup.style.display = 'none';
                excelCellGroup.style.display = 'none';
                
                if (val === 'row') excelRowGroup.style.display = 'block';
                else if (val === 'col') excelColGroup.style.display = 'block';
                else if (val === 'cell') excelCellGroup.style.display = 'block';
                
                updateCommandPreview();
            });
        });
    }

    // 标签切换逻辑
    const tabs = document.querySelectorAll('.format-tab');
    const wordPanel = document.getElementById('wordPanel');
    const excelPanel = document.getElementById('excelPanel');
    const hintSpan = document.getElementById('formatHint');
    
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            currentType = tab.dataset.type;
            
            if (currentType === 'word') {
                wordPanel.style.display = 'block';
                excelPanel.style.display = 'none';
                if (hintSpan) hintSpan.innerHTML = 'Word示例：第2段居中、加粗、字体宋体 | 第1-3段斜体';
            } else {
                wordPanel.style.display = 'none';
                excelPanel.style.display = 'block';
                if (hintSpan) hintSpan.innerHTML = 'Excel示例：第2行加粗、第3列居中、A1单元格字体红色 | 列宽设为20';
            }
            
            updateCommandPreview();
        });
    });
    
  function getWordRangeDescription() {
    const rangeType = document.querySelector('input[name="wordRangeType"]:checked')?.value;
    
    switch(rangeType) {
        case 'full': return '';
        case 'paragraph':
            const pVal = document.getElementById('wordParagraphValue')?.value.trim();
            if (!pVal) return '';
            // 支持：3、1-5、2,4,6 等格式
            return `第${pVal}段`;
        case 'keyword':
            const kVal = document.getElementById('wordKeywordValue')?.value.trim();
            return kVal ? `包含"${kVal}"的段落` : '';
        default: return '';
    }
}
    // 获取 Excel 范围描述
    function getExcelRangeDescription() {
        const rangeType = document.querySelector('input[name="excelRangeType"]:checked')?.value;
        
        switch(rangeType) {
            case 'all': return '';
            case 'row':
                const rowVal = document.getElementById('excelRowNumber')?.value.trim();
                if (!rowVal) return '';
                if (rowVal.includes('-')) return `第${rowVal}行`;
                if (rowVal.includes(',')) return `第${rowVal}行`;
                return `第${rowVal}行`;
            case 'col':
                const colVal = document.getElementById('excelColNumber')?.value.trim();
                if (!colVal) return '';
                if (colVal.includes('-')) return `第${colVal}列`;
                if (colVal.includes(',')) return `第${colVal}列`;
                return `第${colVal}列`;
            case 'cell':
                const cellVal = document.getElementById('excelCellValue')?.value.trim().toUpperCase();
                if (!cellVal) return '';
                return `${cellVal}单元格`;
            case 'header':
                return `表头`;
            default: return '';
        }
    }
    
    // 获取 Word 选中的格式
    function getWordSelectedFormats() {
        const formats = [];
        const align = document.getElementById('wordAlignSelect')?.value;
        const fontStyle = document.getElementById('wordFontStyleSelect')?.value;
        const font = document.getElementById('wordFontSelect')?.value;
        const fontSize = document.getElementById('wordFontSizeSelect')?.value;
        const paragraph = document.getElementById('wordParagraphSelect')?.value;
        const color = document.getElementById('wordColorSelect')?.value;
        
        if (align && align !== '') formats.push(align);
        if (fontStyle && fontStyle !== '') formats.push(fontStyle);
        if (font && font !== '') formats.push(font);
        if (fontSize && fontSize !== '') formats.push(fontSize);
        if (paragraph && paragraph !== '') formats.push(paragraph);
        if (color && color !== '') formats.push(color);
        
        return formats;
    }
    
    // 获取 Excel 选中的格式
    function getExcelSelectedFormats() {
        const formats = [];
        const align = document.getElementById('excelAlignSelect')?.value;
        const fontStyle = document.getElementById('excelFontStyleSelect')?.value;
        const font = document.getElementById('excelFontSelect')?.value;
        const fontSize = document.getElementById('excelFontSizeSelect')?.value;
        const color = document.getElementById('excelColorSelect')?.value;
        const colWidth = document.getElementById('excelColWidth')?.value;
        const rowHeight = document.getElementById('excelRowHeight')?.value;
        
        if (align && align !== '') formats.push(align);
        if (fontStyle && fontStyle !== '') formats.push(fontStyle);
        if (font && font !== '') formats.push(font);
        if (fontSize && fontSize !== '') formats.push(fontSize);
        if (color && color !== '') formats.push(color);
        if (colWidth && colWidth !== '') formats.push(`列宽${colWidth}`);
        if (rowHeight && rowHeight !== '') formats.push(`行高${rowHeight}`);
        
        return formats;
    }
    
    // 更新指令预览
    function updateCommandPreview() {
        const previewText = document.getElementById('commandPreviewText');
        const customCommand = document.getElementById('formatCommand')?.value.trim();
        
        let finalCommand = '';
        
        if (customCommand) {
            finalCommand = customCommand;
        } else {
            if (currentType === 'word') {
                const rangeDesc = getWordRangeDescription();
                const formats = getWordSelectedFormats();
                if (formats.length > 0) {
                    finalCommand = rangeDesc ? `${rangeDesc}${formats.join('、')}` : formats.join('、');
                }
            } else {
                const rangeDesc = getExcelRangeDescription();
                const formats = getExcelSelectedFormats();
                if (formats.length > 0) {
                    finalCommand = rangeDesc ? `${rangeDesc}${formats.join('、')}` : formats.join('、');
                }
            }
        }
        
        if (previewText) previewText.textContent = finalCommand || '-';
    }
    
    // 监听变化
    const wordSelects = document.querySelectorAll('#wordPanel .format-select, #wordPanel .range-input');
    const excelSelects = document.querySelectorAll('#excelPanel .format-select, #excelPanel .range-input, #excelPanel .format-number-input');
    wordSelects.forEach(select => select.addEventListener('change', updateCommandPreview));
    excelSelects.forEach(select => select.addEventListener('change', updateCommandPreview));
    const customInput = document.getElementById('formatCommand');
    if (customInput) customInput.addEventListener('input', updateCommandPreview);
    
    // 进度条和执行逻辑
    function updateProgress(percent, status, detail) {
        const progressArea = document.getElementById('formatProgressArea');
        if (!progressArea) return;
        if (percent === 0) {
            progressArea.innerHTML = '';
            return;
        }
        progressArea.innerHTML = `
            <div class="progress-container loading">
                <div class="progress-header"><div class="progress-title"><i data-lucide="loader"></i><span>格式调整中</span></div><div class="progress-stats"><span class="progress-percent">${percent}%</span></div></div>
                <div class="progress-bar-wrapper"><div class="progress-bar-bg"><div class="progress-bar-fill" style="width: ${percent}%"></div></div></div>
                <div class="progress-details"><div class="progress-status">${status}</div><div class="progress-step">${detail}</div></div>
            </div>
        `;
        lucide.createIcons();
    }
    
    // 执行格式调整
    const executeBtn = document.getElementById('executeFormat');
    if (executeBtn) {
        const newExecuteBtn = executeBtn.cloneNode(true);
        executeBtn.parentNode.replaceChild(newExecuteBtn, executeBtn);
        newExecuteBtn.addEventListener('click', async () => {
            const files = currentFileManager.getFiles();
            if (files.length === 0) {
                Utils.showError('formatError', '请先上传文档');
                return;
            }
            
            let command = document.getElementById('formatCommand')?.value.trim();
            if (!command) {
                if (currentType === 'word') {
                    const rangeDesc = getWordRangeDescription();
                    const formats = getWordSelectedFormats();
                    if (formats.length === 0) {
                        Utils.showError('formatError', '请选择格式指令或输入自定义指令');
                        return;
                    }
                    command = rangeDesc ? `${rangeDesc}${formats.join('、')}` : formats.join('、');
                } else {
                    const rangeDesc = getExcelRangeDescription();
                    const formats = getExcelSelectedFormats();
                    if (formats.length === 0) {
                        Utils.showError('formatError', '请选择格式指令或输入自定义指令');
                        return;
                    }
                    command = rangeDesc ? `${rangeDesc}${formats.join('、')}` : formats.join('、');
                }
            }
            
            const resultArea = document.getElementById('formatResultArea');
            resultArea.innerHTML = '';
            
            updateProgress(30, '正在处理', `应用指令: ${command.substring(0, 50)}...`);
            
            const formData = new FormData();
            for (const file of files) formData.append('documents', file);
            formData.append('command', command);
            
            try {
                const response = await fetch('/api/format/batch', { method: 'POST', body: formData });
                if (!response.ok) throw new Error('格式调整失败');
                
                updateProgress(100, '处理完成', '文档已生成');
                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                
                setTimeout(() => {
                    updateProgress(0, '', '');
                    resultArea.innerHTML = files.length === 1 ? `
                        <div class="format-result success">
                            <div><strong>处理完成！</strong><p style="font-size: 0.75rem;">指令：${command.length > 60 ? command.substring(0, 60) + '...' : command}</p><a href="${url}" class="download-link" download="formatted_${files[0].name}"><i data-lucide="download"></i> 下载调整后的文档</a></div>
                        </div>
                    ` : `
                        <div class="format-result success">
                            <div><strong>批量处理完成！</strong><p style="font-size: 0.75rem;">指令：${command.length > 60 ? command.substring(0, 60) + '...' : command}</p><span>共处理 ${files.length} 个文档</span><a href="${url}" class="download-link" download="formatted_documents.zip"><i data-lucide="download"></i> 下载全部（ZIP打包）</a></div>
                        </div>
                    `;
                    lucide.createIcons();
                }, 500);
            } catch (error) {
                updateProgress(0, '', '');
                Utils.showError('formatError', error.message);
            }
        });
    }
    
    updateCommandPreview();
}

function renderWordRangePanel() {
    return `
        <div class="format-range-section">
            <div class="format-range-header">
                <span>应用范围</span>
            </div>
            <div class="format-range-options">
                <label class="range-option">
                    <input type="radio" name="wordRangeType" value="full" checked>
                    <span>全文</span>
                </label>
                <label class="range-option">
                    <input type="radio" name="wordRangeType" value="paragraph">
                    <span>指定段落</span>
                </label>
                <label class="range-option">
                    <input type="radio" name="wordRangeType" value="keyword">
                    <span>关键词定位</span>
                </label>
            </div>
            
            <!-- 指定段落（支持单个、范围、多个） -->
            <div class="range-input-group" id="wordParagraphGroup" style="display: none;">
                <div class="range-input-wrapper">
                    <input type="text" id="wordParagraphValue" class="range-input" placeholder="段落：3 或 1-5 或 2,4,6">
                </div>
            </div>
            
            <!-- 关键词定位 -->
            <div class="range-input-group" id="wordKeywordGroup" style="display: none;">
                <div class="range-input-wrapper">
                    <input type="text" id="wordKeywordValue" class="range-input" placeholder="关键词，如：甲方、合同">
                </div>
            </div>
        </div>
    `;
}

function renderExcelRangePanel() {
    return `
        <div class="format-range-section">
            <div class="format-range-header">
                <span>应用范围</span>
            </div>
            <div class="format-range-options">
                <label class="range-option">
                    <input type="radio" name="excelRangeType" value="all" checked>
                    <span>全部</span>
                </label>
                <label class="range-option">
                    <input type="radio" name="excelRangeType" value="row">
                    <span>指定行</span>
                </label>
                <label class="range-option">
                    <input type="radio" name="excelRangeType" value="col">
                    <span>指定列</span>
                </label>
                <label class="range-option">
                    <input type="radio" name="excelRangeType" value="cell">
                    <span>单元格</span>
                </label>
                <label class="range-option">
                    <input type="radio" name="excelRangeType" value="header">
                    <span>表头</span>
                </label>
            </div>
            
            <div class="range-input-group" id="excelRowGroup" style="display: none;">
                <div class="range-input-wrapper">
                    <input type="text" id="excelRowNumber" class="range-input" placeholder="行号，如：3 或 1-5 或 2,4,6">
                </div>
            </div>
            
            <div class="range-input-group" id="excelColGroup" style="display: none;">
                <div class="range-input-wrapper">
                    <input type="text" id="excelColNumber" class="range-input" placeholder="列号，如：2 或 1-3">
                </div>
            </div>
            
            <div class="range-input-group" id="excelCellGroup" style="display: none;">
                <div class="range-input-wrapper">
                    <input type="text" id="excelCellValue" class="range-input" placeholder="单元格，如：A1、B3、C5">
                </div>
            </div>
        </div>
    `;
}

function renderWordFormatPanel() {
    return `
        <div class="format-presets-section">
            <div class="format-presets-header">
                <span>Word 格式指令</span>
                <button class="reset-formats-btn" id="resetFormatsBtn">
                    <i data-lucide="refresh-cw" style="width: 12px; height: 12px;"></i> 重置所有
                </button>
            </div>
            <div class="format-select-grid">
                <div class="format-select-item">
                    <label>对齐</label>
                    <select id="wordAlignSelect" class="format-select">
                        <option value="">不设置</option>
                        <option value="居中">居中</option>
                        <option value="左对齐">左对齐</option>
                        <option value="右对齐">右对齐</option>
                        <option value="两端对齐">两端对齐</option>
                    </select>
                </div>
                <div class="format-select-item">
                    <label>字体样式</label>
                    <select id="wordFontStyleSelect" class="format-select">
                        <option value="">不设置</option>
                        <option value="加粗">加粗</option>
                        <option value="斜体">斜体</option>
                        <option value="下划线">下划线</option>
                    </select>
                </div>
                <div class="format-select-item">
                    <label>字体</label>
                    <select id="wordFontSelect" class="format-select">
                        <option value="">不设置</option>
                        <option value="字体宋体">宋体</option>
                        <option value="字体黑体">黑体</option>
                        <option value="字体楷体">楷体</option>
                        <option value="字体微软雅黑">微软雅黑</option>
                    </select>
                </div>
                <div class="format-select-item">
                    <label>字号</label>
                    <select id="wordFontSizeSelect" class="format-select">
                        <option value="">不设置</option>
                        <option value="字号12">12</option>
                        <option value="字号14">14</option>
                        <option value="字号16">16</option>
                        <option value="字号18">18</option>
                        <option value="字号20">20</option>
                    </select>
                </div>
                <div class="format-select-item">
                    <label>段落</label>
                    <select id="wordParagraphSelect" class="format-select">
                        <option value="">不设置</option>
                        <option value="行间距1.5">行间距1.5</option>
                        <option value="行间距2">行间距2</option>
                        <option value="首行缩进">首行缩进</option>
                    </select>
                </div>
                <div class="format-select-item">
                    <label>颜色</label>
                    <select id="wordColorSelect" class="format-select">
                        <option value="">不设置</option>
                        <option value="字体颜色红色">红色</option>
                        <option value="字体颜色蓝色">蓝色</option>
                        <option value="字体颜色绿色">绿色</option>
                    </select>
                </div>
            </div>
        </div>
    `;
}

function renderExcelFormatPanel() {
    return `
        <div class="format-presets-section">
            <div class="format-presets-header">
                <span>Excel 格式指令</span>
                <button class="reset-formats-btn" id="resetFormatsBtn">
                    <i data-lucide="refresh-cw" style="width: 12px; height: 12px;"></i> 重置所有
                </button>
            </div>
            <div class="format-select-grid">
                <div class="format-select-item">
                    <label>对齐</label>
                    <select id="excelAlignSelect" class="format-select">
                        <option value="">不设置</option>
                        <option value="居中">居中</option>
                        <option value="左对齐">左对齐</option>
                        <option value="右对齐">右对齐</option>
                    </select>
                </div>
                <div class="format-select-item">
                    <label>字体样式</label>
                    <select id="excelFontStyleSelect" class="format-select">
                        <option value="">不设置</option>
                        <option value="加粗">加粗</option>
                        <option value="斜体">斜体</option>
                    </select>
                </div>
                <div class="format-select-item">
                    <label>字体</label>
                    <select id="excelFontSelect" class="format-select">
                        <option value="">不设置</option>
                        <option value="字体宋体">宋体</option>
                        <option value="字体黑体">黑体</option>
                        <option value="字体微软雅黑">微软雅黑</option>
                    </select>
                </div>
                <div class="format-select-item">
                    <label>字号</label>
                    <select id="excelFontSizeSelect" class="format-select">
                        <option value="">不设置</option>
                        <option value="字号10">10</option>
                        <option value="字号11">11</option>
                        <option value="字号12">12</option>
                        <option value="字号14">14</option>
                    </select>
                </div>
                <div class="format-select-item">
                    <label>颜色</label>
                    <select id="excelColorSelect" class="format-select">
                        <option value="">不设置</option>
                        <option value="字体颜色红色">红色</option>
                        <option value="字体颜色蓝色">蓝色</option>
                        <option value="字体颜色绿色">绿色</option>
                    </select>
                </div>
                <div class="format-select-item">
                    <label>列宽</label>
                    <input type="number" id="excelColWidth" class="format-number-input" placeholder="如：15" min="5" max="50" step="1">
                </div>
                <div class="format-select-item">
                    <label>行高</label>
                    <input type="number" id="excelRowHeight" class="format-number-input" placeholder="如：20" min="10" max="100" step="1">
                </div>
            </div>
        </div>
    `;
}


// 智能问答 UI
function renderQaUI(container) {
    const fileCount = currentFileManager.getFiles().length;
    container.innerHTML = `
        <div class="chat-messages" id="chatMessages">
            <div class="message assistant">
                <div class="message-avatar"><i data-lucide="bot" style="width: 18px; height: 18px;"></i></div>
                <div class="message-content">您好！已加载 ${fileCount} 个文档，请问有什么可以帮助您？</div>
            </div>
        </div>
        <div style="display: flex; gap: 12px;">
            <input type="text" id="qaInput" class="input" placeholder="输入问题..." onkeypress="if(event.key==='Enter')document.getElementById('sendQa').click()">
            <button class="btn-primary" id="sendQa">
                发送
            </button>
        </div>
        <div style="display: flex; gap: 8px; margin-top: 12px; flex-wrap: wrap;">
            <button class="btn-secondary quick-q" style="padding: 6px 12px;"><i data-lucide="file-text" style="width: 12px; height: 12px;"></i> 总结文档内容</button>
            <button class="btn-secondary quick-q" style="padding: 6px 12px;"><i data-lucide="list" style="width: 12px; height: 12px;"></i> 有哪些关键信息？</button>
            <button class="btn-secondary quick-q" style="padding: 6px 12px;"><i data-lucide="lightbulb" style="width: 12px; height: 12px;"></i> 文档的主要观点是什么？</button>
        </div>
    `;
    lucide.createIcons();
    
    const chatMessages = document.getElementById('chatMessages');
    const qaInput = document.getElementById('qaInput');
    const sendBtn = document.getElementById('sendQa');
    
    // 添加消息到聊天区域
    const addMessage = (role, content, isLoading = false) => {
        const msg = document.createElement('div');
        msg.className = `message ${role}`;
        const icon = role === 'user' ? 'user' : 'bot';
        
        if (isLoading) {
            msg.id = 'loadingMessage';
            msg.innerHTML = `
                <div class="message-avatar"><i data-lucide="${icon}" style="width: 18px; height: 18px;"></i></div>
                <div class="message-content" style="display: flex; align-items: center; gap: 8px;">
                    <i data-lucide="loader" class="spin" style="width: 18px; height: 18px; stroke: var(--accent-solid);"></i>
                    <span>正在思考中...</span>
                </div>
            `;
        } else {
            msg.innerHTML = `
                <div class="message-avatar"><i data-lucide="${icon}" style="width: 18px; height: 18px;"></i></div>
                <div class="message-content">${escapeHtml(content)}</div>
            `;
        }
        
        chatMessages.appendChild(msg);
        chatMessages.scrollTop = chatMessages.scrollHeight;
        lucide.createIcons();
        return msg;
    };
    
    // 移除加载消息
    const removeLoadingMessage = () => {
        const loadingMsg = document.getElementById('loadingMessage');
        if (loadingMsg) {
            loadingMsg.remove();
        }
    };
    
    // 发送问题
    const askQuestion = async (question) => {
        if (!question.trim()) return;
        
        // 禁用输入框和发送按钮
        qaInput.disabled = true;
        sendBtn.disabled = true;
        sendBtn.style.opacity = '0.6';
        sendBtn.style.cursor = 'not-allowed';
        
        // 添加用户消息
        addMessage('user', question);
        
        // 添加加载消息
        addMessage('assistant', '', true);
        
        // 清空输入框
        qaInput.value = '';
        
        const files = currentFileManager.getFiles();
        const formData = new FormData();
        for (const file of files) {
            formData.append('documents', file);
        }
        formData.append('question', question);
        
        try {
            const response = await fetch('/api/qa', {
                method: 'POST',
                body: formData
            });
            
            if (!response.ok) {
                throw new Error('问答服务异常');
            }
            
            const data = await response.json();
            
            // 移除加载消息
            removeLoadingMessage();
            
            // 添加 AI 回复
            addMessage('assistant', data.answer);
            
        } catch (error) {
            // 移除加载消息
            removeLoadingMessage();
            
            // 显示错误消息
            addMessage('assistant', `抱歉，服务暂时不可用：${error.message}`);
        } finally {
            // 恢复输入框和发送按钮
            qaInput.disabled = false;
            sendBtn.disabled = false;
            sendBtn.style.opacity = '';
            sendBtn.style.cursor = '';
            
            // 聚焦输入框
            qaInput.focus();
        }
    };
    
    // 绑定发送按钮
    sendBtn.addEventListener('click', () => {
        askQuestion(qaInput.value);
    });
    
    // 绑定快捷问题
    document.querySelectorAll('.quick-q').forEach(btn => {
        btn.addEventListener('click', () => {
            const question = btn.textContent.trim();
            qaInput.value = question;
            askQuestion(question);
        });
    });
}
// 进度条管理类（美化版）
class ProgressManager {
    constructor(containerId, options = {}) {
        this.container = document.getElementById(containerId);
        this.options = {
            steps: options.steps || [],
            title: options.title || '处理中',
            ...options
        };
        this.currentPercent = 0;
        this.currentStep = 0;
        this.startTime = null;
        this.timer = null;
    }
    
    show() {
        if (!this.container) return;
        this.container.classList.remove('hidden');
        this.container.classList.add('loading');
        this.startTime = Date.now();
        this.updateTime();
        this.reset();
    }
    
    hide() {
        if (!this.container) return;
        this.container.classList.add('hidden');
        if (this.timer) clearInterval(this.timer);
    }
    
    reset() {
        this.currentPercent = 0;
        this.currentStep = 0;
        this.update(0, '准备就绪', '等待开始');
        this.updateSteps(0);
    }
    
    update(percent, status, detail) {
        if (!this.container) return;
        
        this.currentPercent = Math.min(100, Math.max(0, percent));
        
        // 更新进度条宽度
        const fillEl = this.container.querySelector('.progress-bar-fill');
        if (fillEl) fillEl.style.width = `${this.currentPercent}%`;
        
        // 更新百分比显示
        const percentEl = this.container.querySelector('.progress-percent');
        if (percentEl) percentEl.textContent = `${Math.floor(this.currentPercent)}%`;
        
        // 更新状态文字
        const statusEl = this.container.querySelector('.progress-status span');
        if (statusEl && status) statusEl.textContent = status;
        
        // 更新详情
        const detailEl = this.container.querySelector('.progress-step span');
        if (detailEl && detail) detailEl.textContent = detail;
        
        // 根据进度更新步骤指示器
        const stepIndex = Math.floor(percent / (100 / this.options.steps.length));
        if (stepIndex !== this.currentStep && stepIndex < this.options.steps.length) {
            this.currentStep = stepIndex;
            this.updateSteps(stepIndex);
        }
        
        // 更新进度条颜色
        if (percent >= 100) {
            this.container.classList.remove('loading');
            this.container.classList.add('success');
        } else {
            this.container.classList.remove('success', 'error');
            this.container.classList.add('loading');
        }
    }
    
    updateSteps(activeIndex) {
        const stepsContainer = this.container.querySelector('.progress-steps');
        if (!stepsContainer) return;
        
        const steps = stepsContainer.querySelectorAll('.step-item');
        steps.forEach((step, idx) => {
            step.classList.remove('active', 'completed');
            if (idx < activeIndex) {
                step.classList.add('completed');
            } else if (idx === activeIndex) {
                step.classList.add('active');
            }
        });
    }
    
    updateTime() {
        if (this.timer) clearInterval(this.timer);
        this.timer = setInterval(() => {
            if (!this.startTime) return;
            const elapsed = Math.floor((Date.now() - this.startTime) / 1000);
            const minutes = Math.floor(elapsed / 60);
            const seconds = elapsed % 60;
            const timeStr = `${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
            
            const timeEl = this.container.querySelector('.progress-time');
            if (timeEl) timeEl.textContent = timeStr;
        }, 1000);
    }
    
    complete(status = '完成', detail = '处理成功') {
        this.update(100, status, detail);
        if (this.timer) clearInterval(this.timer);
        
        // 显示成功状态
        const statusContainer = this.container.querySelector('.progress-status');
        if (statusContainer) {
            statusContainer.innerHTML = `
                <i data-lucide="check-circle"></i>
                <span>${status}</span>
            `;
            lucide.createIcons();
        }
        
        setTimeout(() => this.hide(), 2000);
    }
    
    error(errorMsg) {
        this.container.classList.remove('loading');
        this.container.classList.add('error');
        
        const statusContainer = this.container.querySelector('.progress-status');
        if (statusContainer) {
            statusContainer.innerHTML = `
                <i data-lucide="alert-circle"></i>
                <span>处理失败</span>
            `;
            lucide.createIcons();
        }
        
        const detailEl = this.container.querySelector('.progress-step span');
        if (detailEl) detailEl.textContent = errorMsg;
        
        setTimeout(() => this.hide(), 3000);
    }
}

// 文档分析 UI（含企业版入口）- 保持原结果展示
function renderAnalyzeUI(container) {
    const files = currentFileManager.getFiles();
    const isMultiple = files.length > 1;
    
    container.innerHTML = `
        <div class="analyze-container">
            <!-- 普通分析（免费） -->
            <div class="analyze-section free-section">
                <div class="analyze-section-header">
                    <span>基础分析</span>
                </div>
                <div class="analyze-type-grid">
                    <label class="analyze-type-option">
                        <input type="checkbox" name="analyzeType" value="summary" checked>
                        <i data-lucide="file-text"></i>
                        <span>文档摘要</span>
                    </label>
                    <label class="analyze-type-option">
                        <input type="checkbox" name="analyzeType" value="keywords" checked>
                        <i data-lucide="hash"></i>
                        <span>关键词提取</span>
                    </label>
                    <label class="analyze-type-option">
                        <input type="checkbox" name="analyzeType" value="stats">
                        <i data-lucide="bar-chart-2"></i>
                        <span>统计信息</span>
                    </label>
                </div>
                ${isMultiple ? `
                <div class="single-doc-hint">
                    <span>当前选择了 ${files.length} 个文档，基础分析仅分析第一个文档。如需多文档对比分析，请升级企业版。</span>
                </div>
                ` : ''}
            </div>
            
            <!-- 企业版功能（展示但锁定） -->
            <div class="analyze-section enterprise-section">
                <div class="analyze-section-header">
                    <span class="enterprise-badge">企业版</span>
                    <span>商业智能分析</span>
                </div>
                <div class="enterprise-features">
                    <div class="feature-item locked" data-feature="contract">
                        <i data-lucide="file-text"></i>
                        <div class="feature-info">
                            <div class="feature-name">合同智能审阅</div>
                            <div class="feature-desc">自动识别风险条款、缺失条款、违约责任</div>
                        </div>
                        <span class="lock-badge"><i data-lucide="lock" style="width: 12px; height: 12px;"></i> </span>
                    </div>
                    <div class="feature-item locked" data-feature="bid">
                        <i data-lucide="trophy"></i>
                        <div class="feature-info">
                            <div class="feature-name">投标文件分析</div>
                            <div class="feature-desc">完整性检查、评分要点提取、报价分析</div>
                        </div>
                        <span class="lock-badge"><i data-lucide="lock" style="width: 12px; height: 12px;"></i> </span>
                    </div>
                    <div class="feature-item locked" data-feature="financial">
                        <i data-lucide="trending-up"></i>
                        <div class="feature-info">
                            <div class="feature-name">财报/研报分析</div>
                            <div class="feature-desc">财务指标提取、趋势分析、投资亮点</div>
                        </div>
                        <span class="lock-badge"><i data-lucide="lock" style="width: 12px; height: 12px;"></i> </span>
                    </div>
                    <div class="feature-item locked" data-feature="policy">
                        <i data-lucide="scroll"></i>
                        <div class="feature-info">
                            <div class="feature-name">政策文件解读</div>
                            <div class="feature-desc">政策要点提取、适用条件判断、申报指南</div>
                        </div>
                        <span class="lock-badge"><i data-lucide="lock" style="width: 12px; height: 12px;"></i> </span>
                    </div>
                    <div class="feature-item locked" data-feature="risk">
                        <i data-lucide="alert-triangle"></i>
                        <div class="feature-info">
                            <div class="feature-name">风险识别</div>
                            <div class="feature-desc">多维度风险检测、合规性检查</div>
                        </div>
                        <span class="lock-badge"><i data-lucide="lock" style="width: 12px; height: 12px;"></i> </span>
                    </div>
                    <div class="feature-item locked" data-feature="batch">
                        <i data-lucide="layers"></i>
                        <div class="feature-info">
                            <div class="feature-name">批量分析</div>
                            <div class="feature-desc">同时分析多个文档，生成对比报告</div>
                        </div>
                        <span class="lock-badge"><i data-lucide="lock" style="width: 12px; height: 12px;"></i> </span>
                    </div>
                </div>
                
                <!-- 升级提示 -->
                <div class="upgrade-prompt">
                    <i data-lucide="crown"></i>
                    <div class="upgrade-text">
                        <strong>解锁全部商业智能功能</strong>
                        <span>升级到企业版，获得合同审阅、风险识别、批量分析等高级功能</span>
                    </div>
                    <button class="upgrade-btn" id="upgradeToEnterprise">
                        了解企业版 <i data-lucide="arrow-right"></i>
                    </button>
                </div>
            </div>
            <!-- 分析按钮 -->
            <button class="btn-primary" id="executeAnalyze">
                开始分析
            </button>
            
            <div id="analyzeProgressArea" style="margin-top: 20px;"></div>
            <div id="analyzeResultArea" style="margin-top: 20px;"></div>
            <div id="analyzeError" class="error-message hidden"></div>
        </div>
    `;
    
    lucide.createIcons();
    
    // 企业版功能点击提示
    document.querySelectorAll('.feature-item.locked').forEach(item => {
        item.addEventListener('click', (e) => {
            e.stopPropagation();
            showEnterpriseModal(item.querySelector('.feature-name')?.textContent || '该功能');
        });
    });
    
    // 升级按钮
    const upgradeBtn = document.getElementById('upgradeToEnterprise');
    if (upgradeBtn) {
        upgradeBtn.addEventListener('click', () => {
            showEnterpriseModal();
        });
    }
    
    // 显示企业版弹窗
    function showEnterpriseModal(featureName = '') {
        const modal = document.createElement('div');
        modal.className = 'enterprise-modal';
        modal.innerHTML = `
            <div class="enterprise-modal-overlay"></div>
            <div class="enterprise-modal-content">
                <div class="enterprise-modal-header">
                    <i data-lucide="crown"></i>
                    <h3>解锁企业版功能</h3>
                    <button class="modal-close"><i data-lucide="x"></i></button>
                </div>
                <div class="enterprise-modal-body">
                    ${featureName ? `<p class="feature-highlight">「${featureName}」</p>` : ''}
                    <p>该功能仅限企业版用户使用。</p>
                    <div class="enterprise-benefits">
                        <div class="benefit-item">
                            <i data-lucide="check-circle"></i>
                            <span>合同智能审阅 - 自动识别风险条款</span>
                        </div>
                        <div class="benefit-item">
                            <i data-lucide="check-circle"></i>
                            <span>投标文件分析 - 完整性检查、评分提取</span>
                        </div>
                        <div class="benefit-item">
                            <i data-lucide="check-circle"></i>
                            <span>财报/研报分析 - 财务指标提取</span>
                        </div>
                        <div class="benefit-item">
                            <i data-lucide="check-circle"></i>
                            <span>批量文档处理 - 同时分析多个文档</span>
                        </div>
                        <div class="benefit-item">
                            <i data-lucide="check-circle"></i>
                            <span>API 接口调用 - 集成到您的系统</span>
                        </div>
                    </div>
                </div>
                <div class="enterprise-modal-footer">
                    <button class="btn-secondary" id="modalCloseBtn">暂不考虑</button>
                    <button class="btn-primary" id="modalContactBtn">联系销售</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
        lucide.createIcons();
        
        modal.querySelector('.enterprise-modal-overlay').addEventListener('click', () => modal.remove());
        modal.querySelector('.modal-close').addEventListener('click', () => modal.remove());
        modal.querySelector('#modalCloseBtn').addEventListener('click', () => modal.remove());
        modal.querySelector('#modalContactBtn').addEventListener('click', () => {
            alert('感谢您的关注！我们的销售团队会尽快与您联系。\n\n演示模式：实际项目中会跳转到联系页面。');
            modal.remove();
        });
    }
    
    // 执行分析
    const analyzeBtn = document.getElementById('executeAnalyze');
    if (analyzeBtn) {
        const newAnalyzeBtn = analyzeBtn.cloneNode(true);
        analyzeBtn.parentNode.replaceChild(newAnalyzeBtn, analyzeBtn);
        newAnalyzeBtn.addEventListener('click', async () => {
            const files = currentFileManager.getFiles();
            if (files.length === 0) {
                Utils.showError('analyzeError', '请先上传文档');
                return;
            }
            
            // 获取选中的分析类型
            const analyzeTypes = [];
            document.querySelectorAll('input[name="analyzeType"]:checked').forEach(cb => {
                analyzeTypes.push(cb.value);
            });
            
            if (analyzeTypes.length === 0) {
                Utils.showError('analyzeError', '请至少选择一种分析类型');
                return;
            }
            
            const resultArea = document.getElementById('analyzeResultArea');
            const progressArea = document.getElementById('analyzeProgressArea');
            const errorEl = document.getElementById('analyzeError');
            
            resultArea.innerHTML = '';
            errorEl.classList.add('hidden');
            
            const targetFile = files[0];
            
            // ========== 显示进度条 ==========
            progressArea.innerHTML = `
                <div class="progress-container-enhanced" id="analyzeProgress">
                    <div class="progress-header">
                        <div class="progress-title">
                            <i data-lucide="loader" class="progress-icon spin"></i>
                            <span id="analyzeProgressStatus">准备分析</span>
                        </div>
                        <div class="progress-percent-wrapper">
                            <span id="analyzeProgressPercent">0%</span>
                        </div>
                    </div>
                    <div class="progress-bar-wrapper">
                        <div class="progress-bar-bg">
                            <div class="progress-bar-fill" id="analyzeProgressFill" style="width: 0%"></div>
                        </div>
                    </div>
                    <div id="analyzeProgressDetail" class="progress-detail" style="display: none;"></div>
                    <div class="progress-steps">
                        <div class="progress-step-item" id="analyzeStep1">
                            <span class="step-icon">1</span>
                            <span class="step-label">读取文档</span>
                        </div>
                        <div class="progress-step-item" id="analyzeStep2">
                            <span class="step-icon">2</span>
                            <span class="step-label">AI 分析</span>
                        </div>
                        <div class="progress-step-item" id="analyzeStep3">
                            <span class="step-icon">3</span>
                            <span class="step-label">生成报告</span>
                        </div>
                    </div>
                </div>
            `;
            lucide.createIcons();
            
            // 进度条辅助函数
            function updateAnalyzeProgress(percent, status, detail = '', stepIndex = -1) {
                const fill = document.getElementById('analyzeProgressFill');
                const percentEl = document.getElementById('analyzeProgressPercent');
                const statusEl = document.getElementById('analyzeProgressStatus');
                const detailEl = document.getElementById('analyzeProgressDetail');
                
                if (fill) fill.style.width = percent + '%';
                if (percentEl) percentEl.textContent = percent + '%';
                if (statusEl) statusEl.textContent = status;
                if (detailEl) {
                    if (detail) {
                        detailEl.textContent = detail;
                        detailEl.style.display = 'block';
                    } else {
                        detailEl.style.display = 'none';
                    }
                }
                
                ['analyzeStep1', 'analyzeStep2', 'analyzeStep3'].forEach((id, idx) => {
                    const step = document.getElementById(id);
                    if (step) {
                        step.classList.remove('active', 'completed');
                        if (idx < stepIndex) step.classList.add('completed');
                        else if (idx === stepIndex) step.classList.add('active');
                    }
                });
            }
            
            function completeAnalyzeProgress(status, detail) {
                updateAnalyzeProgress(100, status, detail, 3);
                setTimeout(() => { if (progressArea) progressArea.innerHTML = ''; }, 2000);
            }
            
            // 开始处理
            updateAnalyzeProgress(20, '读取文档', `正在读取: ${targetFile.name}`, 0);
            
            try {
                // ========== 同时获取分析数据和提取数据 ==========
                updateAnalyzeProgress(40, 'AI 分析中', '正在分析文档内容和提取关键数据...', 1);
                
                // 并行请求：文档分析 + 智能提取
                const formData1 = new FormData();
                formData1.append('documents', targetFile);
                formData1.append('analyze_types', JSON.stringify(analyzeTypes));
                
                // 智能提取：尝试提取常见的结构化字段
                const formData2 = new FormData();
                formData2.append('documents', targetFile);
                formData2.append('command', '提取城市、GDP总量、人口、人均GDP、财政收入、财政支出');
                
                const [analyzeResponse, extractResponse] = await Promise.all([
                    fetch('/api/analyze', { method: 'POST', body: formData1 }),
                    fetch('/api/extract', { method: 'POST', body: formData2 })
                ]);
                
                updateAnalyzeProgress(75, '处理结果', '正在整理数据...', 1);
                
                let analyzeData = {};
                let extractData = { fields: [], data: [] };
                
                if (analyzeResponse.ok) {
                    analyzeData = await analyzeResponse.json();
                }
                
                if (extractResponse.ok) {
                    extractData = await extractResponse.json();
                }
                
                updateAnalyzeProgress(90, '生成报告', '正在渲染结果...', 2);
                
                // 合并数据：将提取的数据注入到分析结果中
                const combinedData = {
                    ...analyzeData,
                    extractedData: extractData.data || [],
                    extractedFields: extractData.fields || []
                };
                
                completeAnalyzeProgress('分析完成', '报告已生成');
                // 保存到全局变量以便调试
                window.lastAnalyzeData = combinedData;
                console.log('合并后的数据:', combinedData);
                console.log('提取的数据:', combinedData.extractedData);
                console.log('提取的字段:', combinedData.extractedFields);
                // 渲染结果，传入合并后的数据
                renderAnalyzeResult(combinedData, [targetFile], analyzeTypes, resultArea);
                
            } catch (error) {
                console.error('分析失败:', error);
                progressArea.innerHTML = '';
                Utils.showError('analyzeError', error.message);
            }
        });
}
}

// 渲染分析结果
function renderAnalyzeResult(data, files, analyzeTypes, container) {
    // 判断是否有数据可以可视化
    const hasData = data.data && data.data.length > 0;
    const hasKeywords = data.keywords && data.keywords.length > 0;
    const showVisualization = hasData || hasKeywords;
    
    // 构建标签页列表
    const allTabs = [...analyzeTypes];
    if (showVisualization) {
        allTabs.push('visualization');
    }
    
    let html = `
        <div class="analyze-result-container">
            <div class="analyze-result-header">
                <i data-lucide="check-circle" style="width: 20px; height: 20px;"></i>
                <span>分析完成！共分析 ${files.length} 个文档</span>
            </div>
            <div class="analyze-result-tabs">
                ${allTabs.map(type => `
                    <button class="analyze-tab ${type === allTabs[0] ? 'active' : ''}" data-tab="${type}">
                        ${getAnalyzeTabIcon(type)} ${getAnalyzeTabName(type)}
                    </button>
                `).join('')}
            </div>
            <div class="analyze-result-content">
                ${allTabs.map(type => `
                    <div id="analyze-tab-${type}" class="analyze-tab-content ${type === allTabs[0] ? 'active' : ''}">
                        ${type === 'visualization' ? renderVisualizationContent(data) : renderAnalyzeTabContent(data, type)}
                    </div>
                `).join('')}
            </div>
            <div class="analyze-result-footer">
                <button class="btn-secondary" id="exportAnalyzeResult">
                    <i data-lucide="download" style="width: 14px; height: 14px;"></i> 导出分析报告
                </button>
            </div>
        </div>
    `;
    
    container.innerHTML = html;
    lucide.createIcons();
    
    // 标签切换
    document.querySelectorAll('.analyze-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            const tabId = tab.dataset.tab;
            
            // 切换标签激活状态
            document.querySelectorAll('.analyze-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            
            // 切换内容显示
            document.querySelectorAll('.analyze-tab-content').forEach(content => {
                content.classList.remove('active');
            });
            document.getElementById(`analyze-tab-${tabId}`).classList.add('active');
            
            // ========== 关键：如果切换到可视化标签，渲染智能图表 ==========
            if (tabId === 'visualization') {
                setTimeout(() => {
                    if (typeof initVisualizationEvents === 'function') {
                        initVisualizationEvents(data);
                    } else {
                        console.warn('initVisualizationEvents 函数未定义');
                    }
                    lucide.createIcons();
                }, 100);
            }
        });
    });
    
    // 如果默认显示的就是可视化标签，立即渲染图表
    if (allTabs[0] === 'visualization') {
        setTimeout(() => {
            if (typeof initVisualizationEvents === 'function') {
                initVisualizationEvents(data);
            }
            lucide.createIcons();
        }, 100);
    }
    
    // 导出结果
    document.getElementById('exportAnalyzeResult')?.addEventListener('click', () => {
        const report = {
            files: files.map(f => f.name),
            analyze_time: new Date().toISOString(),
            results: data
        };
        const blob = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `document_analysis_${Date.now()}.json`;
        a.click();
        URL.revokeObjectURL(url);
    });
}

function getAnalyzeTabIcon(type) {
    const icons = {
        summary: '<i data-lucide="file-text"></i>',
        keywords: '<i data-lucide="hash"></i>',
        entities: '<i data-lucide="users"></i>',
        sentiment: '<i data-lucide="smile"></i>',
        topics: '<i data-lucide="layers"></i>',
        stats: '<i data-lucide="bar-chart-2"></i>'
    };
    return icons[type] || '<i data-lucide="circle"></i>';
}

function getAnalyzeTabName(type) {
    const names = {
        summary: '文档摘要',
        keywords: '关键词',
        entities: '实体识别',
        sentiment: '情感分析',
        topics: '主题分类',
        stats: '统计信息',
        visualization: '数据视图'
    };
    return names[type] || type;
}

function renderAnalyzeTabContent(data, type) {
    switch(type) {
        case 'summary':
            const summary = data.summary || '暂无摘要信息';
            // 清理可能的JSON格式
            const cleanSummary = String(summary).replace(/^"|"$/g, '').replace(/\\n/g, '\n');
            return `
                <div class="analyze-summary">
                    <div class="summary-item">
                        <div class="summary-label">核心摘要</div>
                        <div class="summary-text">${escapeHtml(cleanSummary)}</div>
                    </div>
                </div>
            `;
            
        case 'keywords':
            let keywords = data.keywords || [];
            // 如果返回的是字符串，转换为数组
            if (typeof keywords === 'string') {
                keywords = keywords.split(/[，,、\s]+/).filter(k => k.trim());
            }
            if (!Array.isArray(keywords)) {
                keywords = [];
            }
            return `
                <div class="analyze-keywords">
                    <div class="keyword-cloud">
                        ${keywords.map(k => `
                            <span class="keyword-tag">${escapeHtml(k)}</span>
                        `).join('')}
                        ${keywords.length === 0 ? '<div class="empty-result">未提取到关键词</div>' : ''}
                    </div>
                </div>
            `;
            
        case 'entities':
            let entities = data.entities || {};
            if (typeof entities === 'string') {
                entities = { "实体": entities.split(/[，,、\s]+/) };
            }
            return `
                <div class="analyze-entities">
                    ${Object.entries(entities).map(([type, items]) => {
                        let itemArray = Array.isArray(items) ? items : [items];
                        return `
                            <div class="entity-group">
                                <div class="entity-type">${escapeHtml(type)}</div>
                                <div class="entity-list">
                                    ${itemArray.map(item => `<span class="entity-item">${escapeHtml(item)}</span>`).join('')}
                                </div>
                            </div>
                        `;
                    }).join('')}
                    ${Object.keys(entities).length === 0 ? '<div class="empty-result">未识别到实体</div>' : ''}
                </div>
            `;
            
        case 'sentiment':
            const sentiment = data.sentiment || { label: '中性', score: 0 };
            const label = sentiment.label || '中性';
            const score = sentiment.score || 0;
            return `
                <div class="analyze-sentiment">
                    <div class="sentiment-score">
                        <div class="sentiment-label">情感倾向</div>
                        <div class="sentiment-value ${label === '积极' ? 'positive' : label === '消极' ? 'negative' : 'neutral'}">
                            ${label}
                        </div>
                        <div class="sentiment-bar">
                            <div class="sentiment-fill" style="width: ${(score + 1) * 50}%"></div>
                        </div>
                        <div class="sentiment-detail">置信度: ${Math.abs(score).toFixed(2)}</div>
                    </div>
                </div>
            `;
            
        case 'topics':
            let topics = data.topics || [];
            if (!Array.isArray(topics)) {
                topics = [];
            }
            return `
                <div class="analyze-topics">
                    ${topics.map(t => `
                        <div class="topic-item">
                            <div class="topic-name">${escapeHtml(t.name || '未知')}</div>
                            <div class="topic-confidence">${Math.round((t.confidence || 0.8) * 100)}%</div>
                            <div class="topic-bar">
                                <div class="topic-fill" style="width: ${(t.confidence || 0.8) * 100}%"></div>
                            </div>
                        </div>
                    `).join('')}
                    ${topics.length === 0 ? '<div class="empty-result">未识别到主题</div>' : ''}
                </div>
            `;
            
        case 'stats':
            const stats = data.stats || {};
            return `
                <div class="analyze-stats">
                    <div class="stat-grid">
                        <div class="stat-card">
                            <div class="stat-value">${stats.char_count || 0}</div>
                            <div class="stat-label">字符数</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value">${stats.word_count || 0}</div>
                            <div class="stat-label">词数</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value">${stats.paragraph_count || 0}</div>
                            <div class="stat-label">段落数</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value">${stats.reading_time || 0}</div>
                            <div class="stat-label">阅读时长(分钟)</div>
                        </div>
                    </div>
                </div>
            `;
            
        default:
            return '<div class="empty-result">暂无数据</div>';
    }
}

// HTML 转义函数
function escapeHtml(text) {
    if (!text) return '';
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

// ========== 智能可视化系统 ==========

// 数据分析器 - 兼容多种数据格式
class DataAnalyzer {
    constructor(data) {
        this.data = data;
        this.rawData = data;
        
        // 尝试多种方式获取结构化数据
        this.records = this.extractRecords(data);
        this.fields = this.extractFields(data);
        this.stats = data.stats || {};
        this.keywords = data.keywords || [];
    }
    
    // 从各种格式中提取记录
    extractRecords(data) {
        // ========== 优先使用智能提取的数据 ==========
        if (data.extractedData && Array.isArray(data.extractedData) && data.extractedData.length > 0) {
            console.log('使用智能提取的数据:', data.extractedData.length, '条');
            return data.extractedData;
        }
        
        // 格式1: 标准格式 { fields: [], data: [] }
        if (data.data && Array.isArray(data.data) && data.data.length > 0) {
            return data.data;
        }
        
        // 格式2: 直接从表格提取的数据（可能是数组）
        if (Array.isArray(data) && data.length > 0) {
            return data;
        }
        
        // 格式3: 从 stats 构建虚拟数据（降级方案）
        if (data.stats && Object.keys(data.stats).length > 0) {
            return [{
                name: '字符数',
                value: data.stats.char_count || 0
            }, {
                name: '词数',
                value: data.stats.word_count || 0
            }, {
                name: '段落数',
                value: data.stats.paragraph_count || 0
            }, {
                name: '阅读时长(分钟)',
                value: data.stats.reading_time || 0
            }];
        }
        
        // 格式4: 从关键词构建数据（降级方案）
        if (data.keywords && Array.isArray(data.keywords) && data.keywords.length > 0) {
            return data.keywords.slice(0, 10).map((keyword, index) => ({
                name: typeof keyword === 'string' ? keyword : keyword.name || keyword,
                value: Math.floor(100 - index * 5)
            }));
        }
        
        return [];
    }

    extractFields(data) {
        // 优先使用提取的字段
        if (data.extractedFields && Array.isArray(data.extractedFields)) {
            return data.extractedFields;
        }
        
        if (data.fields && Array.isArray(data.fields)) {
            return data.fields;
        }
        
        if (this.records.length > 0) {
            return Object.keys(this.records[0]);
        }
        
        return [];
    }

    hasNumericFields() {
        return this.getNumericFields().length > 0;
    }
    
    getNumericFields() {
        if (this.records.length === 0) return [];
        const sample = this.records[0];
        return Object.keys(sample).filter(key => {
            const val = sample[key];
            return typeof val === 'number' || (!isNaN(parseFloat(val)) && isFinite(val));
        });
    }
    
    getCategoryFields() {
        if (this.records.length === 0) return [];
        const sample = this.records[0];
        const numericFields = this.getNumericFields();
        return Object.keys(sample).filter(key => !numericFields.includes(key));
    }
    
    analyzeDataType() {
        const recordCount = this.records.length;
        const numericFields = this.getNumericFields();
        const categoryFields = this.getCategoryFields();
        
        console.log('数据分析结果:', { recordCount, numericFields, categoryFields });
        
        if (recordCount === 0) {
            return { 
                type: 'empty', 
                recommendation: '暂无结构化数据', 
                charts: [], 
                preferred: null 
            };
        }
        
        // 有分类字段和数值字段
        if (categoryFields.length > 0 && numericFields.length > 0) {
            return {
                type: 'comparison',
                recommendation: '数据对比',
                charts: ['bar', 'pie', 'table'],
                preferred: 'bar',
                categoryField: categoryFields[0],
                valueFields: numericFields
            };
        }
        
        // 只有数值字段
        if (numericFields.length > 0) {
            return {
                type: 'numeric',
                recommendation: '数值分析',
                charts: ['bar', 'pie', 'table'],
                preferred: 'bar',
                categoryField: 'name',
                valueFields: numericFields
            };
        }
        
        // 只有分类字段 - 可以统计频次
        if (categoryFields.length > 0) {
            // 转换为频次数据
            const freqMap = {};
            this.records.forEach(r => {
                const val = String(r[categoryFields[0]] || '未知');
                freqMap[val] = (freqMap[val] || 0) + 1;
            });
            
            // 保存频次数据供后续使用
            this.frequencyData = Object.entries(freqMap).map(([name, value]) => ({ name, value }));
            
            return {
                type: 'frequency',
                recommendation: '频次统计',
                charts: ['bar', 'pie', 'table'],
                preferred: 'bar',
                categoryField: categoryFields[0]
            };
        }
        
        return { 
            type: 'table', 
            recommendation: '数据表格', 
            charts: ['table'], 
            preferred: 'table' 
        };
    }
}
function renderVisualizationContent(data) {
    console.log('渲染可视化，原始数据:', data);
    
    const analyzer = new DataAnalyzer(data);
    const analysis = analyzer.analyzeDataType();
    const recordCount = analyzer.records.length;
    const numericFields = analyzer.getNumericFields();
    const categoryFields = analyzer.getCategoryFields();
    
    console.log('分析结果:', { analysis, recordCount, numericFields, categoryFields });
    
    if (analysis.type === 'empty' || recordCount === 0) {
        // 检查是否有统计信息可以展示
        if (data.stats && Object.keys(data.stats).length > 0) {
            return renderStatsOnlyView(data.stats);
        }
        
        // 检查是否有摘要
        if (data.summary) {
            return `
                <div class="visualization-message">
                    <i data-lucide="file-text" style="width: 40px; height: 40px; opacity: 0.5;"></i>
                    <p>文档摘要</p>
                    <div class="summary-box">${escapeHtml(String(data.summary))}</div>
                    <small style="margin-top: 16px; display: block; color: var(--text-placeholder);">
                        需要结构化数据才能生成图表。请使用"智能提取"功能先提取表格数据。
                    </small>
                </div>
            `;
        }
        
        return `
            <div class="visualization-empty">
                <i data-lucide="database" style="width: 48px; height: 48px; opacity: 0.4;"></i>
                <p>暂无可可视化的结构化数据</p>
                <small>请使用"智能提取"功能提取表格数据，或选择包含数值的文档进行分析</small>
            </div>
        `;
    }
    
    const availableCharts = analysis.charts || ['bar', 'table'];
    
    return `
        <div class="smart-visualization">
            <div class="data-summary">
                <div class="summary-item">
                    <span class="summary-label">推荐视图</span>
                    <span class="summary-value">${analysis.recommendation}</span>
                </div>
                <div class="summary-item">
                    <span class="summary-label">记录数</span>
                    <span class="summary-value">${recordCount} 条</span>
                </div>
                <div class="summary-item">
                    <span class="summary-label">数值字段</span>
                    <span class="summary-value">${numericFields.length > 0 ? numericFields.slice(0, 3).join('、') : '无'}</span>
                </div>
                <div class="summary-item">
                    <span class="summary-label">分类字段</span>
                    <span class="summary-value">${categoryFields.length > 0 ? categoryFields.slice(0, 3).join('、') : '无'}</span>
                </div>
            </div>
            
            <div class="chart-selector">
                <span class="selector-label">
                    <i data-lucide="chart-bar" style="width: 14px; height: 14px;"></i>
                    图表类型：
                </span>
                <div class="chart-type-buttons">
                    ${availableCharts.includes('bar') ? '<button class="chart-type-btn active" data-chart="bar"><i data-lucide="bar-chart-3"></i> 柱状图</button>' : ''}
                    ${availableCharts.includes('pie') ? '<button class="chart-type-btn" data-chart="pie"><i data-lucide="pie-chart"></i> 饼图</button>' : ''}
                    ${availableCharts.includes('line') ? '<button class="chart-type-btn" data-chart="line"><i data-lucide="trending-up"></i> 折线图</button>' : ''}
                    <button class="chart-type-btn ${!availableCharts.includes('bar') ? 'active' : ''}" data-chart="table"><i data-lucide="table"></i> 数据表</button>
                </div>
            </div>
            
            <div class="chart-main-container">
                <div id="mainChart" class="main-chart" style="height: 350px;"></div>
                <div id="chartTable" class="chart-table-container" style="display: none;"></div>
            </div>
            
            <div class="visualization-tip">
                <i data-lucide="info" style="width: 14px; height: 14px;"></i>
                <span>根据数据特征，推荐使用${analysis.recommendation}视图。</span>
            </div>
        </div>
    `;
}

// 只显示统计信息的视图
function renderStatsOnlyView(stats) {
    return `
        <div class="stats-only-view">
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-value">${stats.char_count || 0}</div>
                    <div class="stat-label">字符数</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">${stats.word_count || 0}</div>
                    <div class="stat-label">词数</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">${stats.paragraph_count || 0}</div>
                    <div class="stat-label">段落数</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">${stats.reading_time || 0}</div>
                    <div class="stat-label">阅读时长(分钟)</div>
                </div>
            </div>
            <div class="visualization-tip" style="margin-top: 20px;">
                <i data-lucide="info" style="width: 14px; height: 14px;"></i>
                <span>这是文档的基础统计信息。如需更多可视化，请使用"智能提取"功能提取表格数据。</span>
            </div>
        </div>
    `;
}
function renderSmartChart(data, chartType = 'bar') {
    console.log('渲染图表，类型:', chartType);
    console.log('传入数据:', data);
    
    const analyzer = new DataAnalyzer(data);
    console.log('分析器记录数:', analyzer.records.length);
    console.log('分析器字段:', analyzer.fields);
    console.log('数值字段:', analyzer.getNumericFields());
    console.log('分类字段:', analyzer.getCategoryFields());
    
    const chartDom = document.getElementById('mainChart');
    const tableDom = document.getElementById('chartTable');
    
    if (!chartDom) {
        console.error('找不到图表容器 #mainChart');
        return;
    }
    
    if (typeof echarts === 'undefined') {
        console.error('ECharts 未加载');
        chartDom.innerHTML = '<p style="text-align:center;padding:40px;">图表库加载失败</p>';
        return;
    }
    
    // 隐藏表格，显示图表
    if (tableDom) tableDom.style.display = 'none';
    chartDom.style.display = 'block';
    
    const chart = echarts.init(chartDom);
    let option = {};
    
    if (chartType === 'bar') {
        option = generateBarChartOption(analyzer);
    } else if (chartType === 'pie') {
        option = generatePieChartOption(analyzer);
    } else if (chartType === 'line') {
        option = generateLineChartOption(analyzer);
    } else if (chartType === 'table') {
        chartDom.style.display = 'none';
        if (tableDom) {
            tableDom.style.display = 'block';
            renderDataTable(analyzer, tableDom);
        }
        return;
    }
    
    console.log('图表配置:', option);
    chart.setOption(option);
    
    // 响应式
    window.addEventListener('resize', () => chart.resize());
}
// 生成柱状图配置
function generateBarChartOption(analyzer) {
    const records = analyzer.records;
    
    console.log('生成柱状图，记录数:', records.length);
    console.log('记录样例:', records.slice(0, 3));
    
    if (records.length === 0) {
        return {
            title: { text: '暂无数据', left: 'center', top: 'center' }
        };
    }
    
    // 智能选择字段
    let categoryField = analyzer.getCategoryFields()[0];
    let valueField = analyzer.getNumericFields()[0];
    
    // 如果没有分类字段，使用第一个字段
    if (!categoryField && Object.keys(records[0]).length > 0) {
        categoryField = Object.keys(records[0])[0];
    }
    
    // 如果没有数值字段，尝试将字符串转换为数值
    if (!valueField) {
        for (const key of Object.keys(records[0])) {
            const val = records[0][key];
            if (typeof val === 'string' && !isNaN(parseFloat(val))) {
                valueField = key;
                break;
            }
        }
    }
    
    // 如果还是没有，使用第二个字段
    if (!valueField && Object.keys(records[0]).length > 1) {
        valueField = Object.keys(records[0])[1];
    }
    
    console.log('选择的分类字段:', categoryField);
    console.log('选择的数值字段:', valueField);
    
    if (!categoryField || !valueField) {
        return {
            title: { text: '无法识别数据字段', left: 'center', top: 'center' }
        };
    }
    
    // 提取数据
    const categories = [];
    const values = [];
    
    records.slice(0, 15).forEach(r => {
        let catValue = r[categoryField] || '未知';
        // 截断过长的分类名
        if (catValue.length > 15) {
            catValue = catValue.substring(0, 12) + '...';
        }
        categories.push(String(catValue));
        
        let numValue = r[valueField];
        if (typeof numValue === 'string') {
            // 尝试提取数字（处理带单位的字符串如 "56708.71亿元"）
            const match = String(numValue).match(/(\d+\.?\d*)/);
            numValue = match ? parseFloat(match[1]) : parseFloat(numValue) || 0;
        }
        values.push(Number(numValue) || 0);
    });
    
    console.log('分类数据:', categories);
    console.log('数值数据:', values);
    
    // 过滤掉数值全为0的情况
    const hasData = values.some(v => v > 0);
    
    return {
        title: {
            text: `${valueField} 对比分析`,
            left: 'center',
            top: 5,
            textStyle: { fontSize: 14, fontWeight: 500, color: 'var(--text-primary)' }
        },
        tooltip: { 
            trigger: 'axis',
            formatter: function(params) {
                return params[0].name + '<br/>' + 
                       params[0].marker + ' ' + valueField + ': ' + 
                       params[0].value.toLocaleString();
            }
        },
        grid: { 
            left: '15%', 
            right: '5%', 
            bottom: hasData ? '20%' : '10%', 
            top: '18%' 
        },
        xAxis: {
            type: 'category',
            data: categories,
            axisLabel: { 
                rotate: categories.length > 5 ? 45 : 0, 
                color: 'var(--text-secondary)', 
                fontSize: 11,
                interval: 0
            },
            axisLine: { lineStyle: { color: 'var(--border-light)' } }
        },
        yAxis: {
            type: 'value',
            name: valueField,
            axisLabel: { 
                color: 'var(--text-secondary)',
                formatter: function(value) {
                    if (value >= 100000000) return (value / 100000000).toFixed(1) + '亿';
                    if (value >= 10000) return (value / 10000).toFixed(0) + '万';
                    return value;
                }
            },
            axisLine: { lineStyle: { color: 'var(--border-light)' } },
            splitLine: { lineStyle: { color: 'var(--border-light)', type: 'dashed' } }
        },
        series: [{
            name: valueField,
            type: 'bar',
            data: values,
            itemStyle: {
                // 使用设计系统的渐变色
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                    { offset: 0, color: '#2C3E66' },      // var(--accent-solid)
                    { offset: 0.5, color: '#4A6A8A' },    // var(--gradient-end)
                    { offset: 1, color: '#6B8CBF' }       // 更浅的蓝色
                ]),
                borderRadius: [4, 4, 0, 0]
            },
            label: {
                show: hasData && values.length <= 10,
                position: 'top',
                color: '#5B6E8C',  // var(--text-secondary)
                fontSize: 11,
                formatter: function(params) {
                    const val = params.value;
                    if (val >= 100000000) return (val / 100000000).toFixed(1) + '亿';
                    if (val >= 10000) return (val / 10000).toFixed(0) + '万';
                    return val.toLocaleString();
                }
            }
        }]
    };
}
function generatePieChartOption(analyzer) {
    const records = analyzer.records;
    
    if (records.length === 0) {
        return { title: { text: '暂无数据', left: 'center', top: 'center' } };
    }
    
    let categoryField = analyzer.getCategoryFields()[0];
    let valueField = analyzer.getNumericFields()[0];
    
    if (!categoryField && Object.keys(records[0]).length > 0) {
        categoryField = Object.keys(records[0])[0];
    }
    if (!valueField && Object.keys(records[0]).length > 1) {
        valueField = Object.keys(records[0])[1];
    }
    
    const pieData = records.slice(0, 10).map(r => {
        let name = String(r[categoryField] || '未知');
        if (name.length > 15) name = name.substring(0, 12) + '...';
        
        let numValue = r[valueField];
        if (typeof numValue === 'string') {
            const match = String(numValue).match(/(\d+\.?\d*)/);
            numValue = match ? parseFloat(match[1]) : parseFloat(numValue) || 0;
        }
        
        return {
            name: name,
            value: Number(numValue) || 0
        };
    }).filter(d => d.value > 0);
    
    console.log('饼图数据:', pieData);
    
    if (pieData.length === 0) {
        return { title: { text: '无有效数值数据', left: 'center', top: 'center' } };
    }
    
    return {
        // 标题颜色
        title: {
            textStyle: { 
                fontSize: 14, 
                fontWeight: 500, 
                color: '#1A2C3E'  // var(--text-primary)
            }
        },

        // X轴
        xAxis: {
            axisLabel: { 
                color: '#5B6E8C',  // var(--text-secondary)
                fontSize: 11
            },
            axisLine: { 
                lineStyle: { color: '#EDF2F7' }  // var(--border-light)
            },
            axisTick: {
                lineStyle: { color: '#EDF2F7' }
            }
        },

        // Y轴
        yAxis: {
            axisLabel: { 
                color: '#5B6E8C'  // var(--text-secondary)
            },
            axisLine: { 
                lineStyle: { color: '#EDF2F7' }  // var(--border-light)
            },
            splitLine: { 
                lineStyle: { 
                    color: '#EDF2F7',  // var(--border-light)
                    type: 'dashed' 
                } 
            }
        },

        // 提示框
        tooltip: {
            backgroundColor: '#FFFFFF',
            borderColor: '#EDF2F7',
            textStyle: { color: '#1A2C3E' }
        },

        // 图例
        legend: {
            textStyle: { 
                color: '#5B6E8C'  // var(--text-secondary)
            }
        },
        series: [{
            type: 'pie',
            radius: ['40%', '70%'],
            center: ['55%', '55%'],
            data: pieData,
            itemStyle: {
                borderRadius: 6,
                borderColor: 'var(--surface-card)',
                borderWidth: 2,
                // 使用设计系统配色
                color: function(params) {
                    const colors = [
                        '#2C3E66',  // 深蓝
                        '#4A6A8A',  // 中蓝
                        '#6B8CBF',  // 浅蓝
                        '#8BA3C8',  // 更浅蓝
                        '#A3B3CC',  // 灰蓝
                        '#5B6E8C',  // 文字灰蓝
                        '#1F2F4D',  // 深色
                        '#819BBB',  // 亮蓝
                    ];
                    return colors[params.dataIndex % colors.length];
                }
            },
            label: {
                color: '#5B6E8C',  // var(--text-secondary)
                fontSize: 11,
                formatter: '{b}: {d}%'
            },
            emphasis: {
                label: { 
                    show: true, 
                    fontWeight: 'bold',
                    color: '#2C3E66'  // var(--accent-solid)
                }
            }
        }]
    };
}
// 生成折线图配置
function generateLineChartOption(analyzer) {
    const records = analyzer.records;
    const dateField = analyzer.findDateField() || analyzer.getCategoryFields()[0] || 'date';
    const valueField = analyzer.getNumericFields()[0] || 'value';
    
    const sortedRecords = [...records].sort((a, b) => {
        return String(a[dateField] || '').localeCompare(String(b[dateField] || ''));
    });
    
    const categories = sortedRecords.map(r => String(r[dateField] || '未知'));
    const values = sortedRecords.map(r => {
        const val = r[valueField];
        return typeof val === 'number' ? val : parseFloat(val) || 0;
    });
    
    return {
       // 标题颜色
        title: {
            textStyle: { 
                fontSize: 14, 
                fontWeight: 500, 
                color: '#1A2C3E'  // var(--text-primary)
            }
        },

        // X轴
        xAxis: {
            axisLabel: { 
                color: '#5B6E8C',  // var(--text-secondary)
                fontSize: 11
            },
            axisLine: { 
                lineStyle: { color: '#EDF2F7' }  // var(--border-light)
            },
            axisTick: {
                lineStyle: { color: '#EDF2F7' }
            }
        },

        // Y轴
        yAxis: {
            axisLabel: { 
                color: '#5B6E8C'  // var(--text-secondary)
            },
            axisLine: { 
                lineStyle: { color: '#EDF2F7' }  // var(--border-light)
            },
            splitLine: { 
                lineStyle: { 
                    color: '#EDF2F7',  // var(--border-light)
                    type: 'dashed' 
                } 
            }
        },

        // 提示框
        tooltip: {
            backgroundColor: '#FFFFFF',
            borderColor: '#EDF2F7',
            textStyle: { color: '#1A2C3E' }
        },

        // 图例
        legend: {
            textStyle: { 
                color: '#5B6E8C'  // var(--text-secondary)
            }
        },
        series: [{
            name: valueField,
            type: 'line',
            data: values,
            smooth: true,
            lineStyle: { 
                color: '#2C3E66',  // var(--accent-solid)
                width: 3 
            },
            areaStyle: {
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                    { offset: 0, color: 'rgba(44, 62, 102, 0.25)' },  // #2C3E66
                    { offset: 1, color: 'rgba(44, 62, 102, 0.05)' }
                ])
            },
            symbol: 'circle',
            symbolSize: 8,
            itemStyle: {
                color: '#2C3E66',
                borderColor: '#FFFFFF',
                borderWidth: 2
            }
        }]
    };
}

// 渲染数据表格
function renderDataTable(analyzer, container) {
    const records = analyzer.records;
    const columns = Object.keys(records[0] || {});
    
    let html = '<table class="data-table"><thead><tr>';
    columns.forEach(col => { html += `<th>${col}</th>`; });
    html += '</tr></thead><tbody>';
    
    records.slice(0, 50).forEach(record => {
        html += '<tr>';
        columns.forEach(col => {
            const val = record[col];
            html += `<td>${val !== undefined && val !== null ? val : ''}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody></table>';
    
    if (records.length > 50) {
        html += `<div style="text-align: center; padding: 12px; color: var(--text-placeholder);">共 ${records.length} 条记录，仅显示前 50 条</div>`;
    }
    
    container.innerHTML = html;
}

// ========== 初始化可视化事件 ==========
function initVisualizationEvents(data) {
    // 图表类型切换
    document.querySelectorAll('.chart-type-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.chart-type-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            const chartType = btn.dataset.chart;
            renderSmartChart(data, chartType);
        });
    });
    
    // 默认渲染推荐图表
    const analyzer = new DataAnalyzer(data);
    const analysis = analyzer.analyzeDataType();
    renderSmartChart(data, analysis.preferred || 'bar');
}