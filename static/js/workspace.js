// 工作区主逻辑
let currentFileManager = null;
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
        graph: '知识图谱'
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
    }
}

function getTitleIcon(func) {
    const icons = {
        extract: 'file-search',
        fill: 'table',
        format: 'paintbrush',
        qa: 'message-circle',
        graph: 'network'
    };
    return icons[func] || 'file-text';
}

// 智能提取 UI
function renderExtractUI(container) {
    container.innerHTML = `
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
    
    document.getElementById('executeExtract').addEventListener('click', async () => {
        const command = document.getElementById('extractCommand').value.trim();
        if (!command) {
            Utils.showError('extractError', '请输入提取指令');
            return;
        }
        
        const area = document.getElementById('extractResultArea');
        area.innerHTML = '<div class="progress-container"><div class="progress-bar-bg"><div class="progress-bar-fill" style="width: 50%"></div></div><div><i data-lucide="loader" style="width: 16px; height: 16px; animation: spin 1s linear infinite;"></i> 处理中...</div></div>';
        lucide.createIcons();
        
        const files = currentFileManager.getFiles();
        const formData = new FormData();
        for (const file of files) {
            formData.append('documents', file);
        }
        formData.append('command', command);
        
        try {
            const response = await fetch('/api/extract', {
                method: 'POST',
                body: formData
            });
            
            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || '提取失败');
            }
            
            const data = await response.json();
            const fields = data.fields || [];
            const records = data.data || [];
            
            if (records.length === 0) {
                area.innerHTML = '<div style="text-align: center; padding: 40px; color: var(--warning);"><i data-lucide="inbox" style="width: 32px; height: 32px; margin-bottom: 12px;"></i><br>未提取到数据，请检查文档内容或指令</div>';
                lucide.createIcons();
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
            
            document.getElementById('exportTableBtn')?.addEventListener('click', () => {
                Utils.exportToCSV(records, fields, '提取结果');
            });
            
        } catch (error) {
            console.error('提取失败:', error);
            Utils.showError('extractError', error.message);
            area.innerHTML = `<div class="error-message"><i data-lucide="alert-circle" style="width: 16px; height: 16px;"></i> ${error.message}</div>`;
            lucide.createIcons();
        }
    });
}
// 表格填写 UI（支持 Word 和 Excel 模板，自动识别）
function renderFillUI(container) {
    
    container.innerHTML = `
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
            
            if (!command) { Utils.showError('fillError', '请输入字段指令'); return; }
            if (!currentTemplateFile) { Utils.showError('fillError', '请上传模板文件'); return; }
            
            const files = currentFileManager.getFiles();
            if (files.length === 0) { Utils.showError('fillError', '请先上传源文档'); return; }
            
            const area = document.getElementById('fillResultArea');
            area.innerHTML = '<div class="progress-container"><div class="progress-bar-bg"><div class="progress-bar-fill" style="width: 50%"></div></div><div><i data-lucide="loader" style="width: 16px; height: 16px; animation: spin 1s linear infinite;"></i> 处理中...</div></div>';
            lucide.createIcons();
            
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
                const response = await fetch('/api/fill', {
                    method: 'POST',
                    body: formData
                });
                
                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.error || '生成失败');
                }
                
                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                const extension = isWordTemplate ? '.docx' : '.xlsx';
                
                area.innerHTML = `
                    <div style="text-align: center; padding: 20px;">
                        <a href="${url}" class="btn-secondary" download="filled_template${extension}" style="display: inline-flex; align-items: center; gap: 8px; padding: 10px 24px; text-decoration: none; border-radius: 40px;">
                            <i data-lucide="download" style="width: 14px; height: 14px;"></i> 下载${isWordTemplate ? '文档' : '表格'}文件
                        </a>
                    </div>
                `;
                lucide.createIcons();
                
            } catch (error) {
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
            <input type="text" id="qaInput" class="input" placeholder="输入问题...">
            <button class="btn-primary" id="sendQa">
                发送
            </button>
        </div>
        <div style="display: flex; gap: 8px; margin-top: 12px; flex-wrap: wrap;">
            <button class="btn-secondary quick-q" style="padding: 6px 12px;"><i data-lucide="file-text" style="width: 12px; height: 12px;"></i> 总结文档内容</button>
            <button class="btn-secondary quick-q" style="padding: 6px 12px;"><i data-lucide="list" style="width: 12px; height: 12px;"></i> 有哪些关键实体？</button>
            <button class="btn-secondary quick-q" style="padding: 6px 12px;"><i data-lucide="lightbulb" style="width: 12px; height: 12px;"></i> 文档的主要观点是什么？</button>
        </div>
    `;
    lucide.createIcons();
    
    const chatMessages = document.getElementById('chatMessages');
    const addMessage = (role, content) => {
        const msg = document.createElement('div');
        msg.className = `message ${role}`;
        const icon = role === 'user' ? 'user' : 'bot';
        msg.innerHTML = `<div class="message-avatar"><i data-lucide="${icon}" style="width: 18px; height: 18px;"></i></div><div class="message-content">${content}</div>`;
        chatMessages.appendChild(msg);
        chatMessages.scrollTop = chatMessages.scrollHeight;
        lucide.createIcons();
    };
    
    const askQuestion = async (question) => {
        addMessage('user', question);
        
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
            addMessage('assistant', data.answer);
            
        } catch (error) {
            addMessage('assistant', `抱歉，服务暂时不可用：${error.message}`);
        }
    };
    
    document.getElementById('sendQa').addEventListener('click', () => {
        const q = document.getElementById('qaInput').value.trim();
        if (!q) return;
        askQuestion(q);
        document.getElementById('qaInput').value = '';
    });
    
    document.querySelectorAll('.quick-q').forEach(btn => {
        btn.addEventListener('click', () => {
            askQuestion(btn.textContent.trim());
        });
    });
}

// 知识图谱 UI（美化版）
function renderGraphUI(container) {
    container.innerHTML = `
        <div class="graph-container">
            <div class="graph-canvas" id="graphCanvas"></div>
            <div class="graph-detail" id="graphDetail">
                <div style="text-align: center; color: var(--text-placeholder); padding: 40px;">
                    <i data-lucide="loader" style="width: 32px; height: 32px; animation: spin 1s linear infinite;"></i><br>
                    正在生成知识图谱...
                </div>
            </div>
        </div>
        <div class="graph-toolbar" style="margin-top: 16px; display: flex; gap: 12px; justify-content: center; flex-wrap: wrap;">
            <button class="graph-tool-btn" id="graphResetView" title="重置视图">
                <i data-lucide="maximize" style="width: 16px; height: 16px;"></i> 重置视图
            </button>
            <button class="graph-tool-btn" id="graphFitView" title="适应屏幕">
                <i data-lucide="zoom-in" style="width: 16px; height: 16px;"></i> 适应屏幕
            </button>
            <button class="graph-tool-btn" id="graphExport" title="导出图片">
                <i data-lucide="camera" style="width: 16px; height: 16px;"></i> 导出图片
            </button>
        </div>
    `;
    
    // 添加加载动画样式
    if (!document.querySelector('#graphSpinStyle')) {
        const style = document.createElement('style');
        style.id = 'graphSpinStyle';
        style.textContent = '@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }';
        document.head.appendChild(style);
    }
    
    lucide.createIcons();
    
    // 绑定工具栏按钮事件
    setTimeout(() => {
        document.getElementById('graphResetView')?.addEventListener('click', () => {
            if (currentChart) currentChart.dispatchAction({ type: 'restore' });
        });
        document.getElementById('graphFitView')?.addEventListener('click', () => {
            if (currentChart) currentChart.dispatchAction({ type: 'graphRoam', roam: 'scale' });
        });
        document.getElementById('graphExport')?.addEventListener('click', () => {
            if (currentChart) {
                const url = currentChart.getDataURL({ type: 'png' });
                const link = document.createElement('a');
                link.download = 'knowledge-graph.png';
                link.href = url;
                link.click();
            }
        });
    }, 100);
    
    const loadGraph = async () => {
        const files = currentFileManager.getFiles();
        if (files.length === 0) {
            document.getElementById('graphDetail').innerHTML = `
                <div style="text-align: center; color: var(--warning); padding: 40px;">
                    <i data-lucide="alert-triangle" style="width: 32px; height: 32px;"></i><br>
                    请先上传文档
                </div>
            `;
            lucide.createIcons();
            return;
        }
        
        const formData = new FormData();
        for (const file of files) {
            formData.append('documents', file);
        }
        
        try {
            const response = await fetch('/api/graph', {
                method: 'POST',
                body: formData
            });
            
            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || '图谱生成失败');
            }
            
            const data = await response.json();
            
            if (!data.nodes || data.nodes.length === 0) {
                throw new Error('未能提取到有效的实体关系');
            }
            
            const canvas = document.getElementById('graphCanvas');
            if (canvas && typeof echarts !== 'undefined') {
                if (currentChart) currentChart.dispose();
                currentChart = echarts.init(canvas);
                
                const isDark = document.documentElement.classList.contains('dark');
                
                const categoryColors = [
                    { light: '#5470c6', dark: '#6E8BC0' },
                    { light: '#fac858', dark: '#F4C542' },
                    { light: '#ee6666', dark: '#E87676' },
                    { light: '#73c0de', dark: '#8FC8E8' },
                    { light: '#3ba272', dark: '#5DB98A' },
                    { light: '#fc8452', dark: '#FCA374' },
                    { light: '#9a60b4', dark: '#B885D0' },
                    { light: '#ea7ccc', dark: '#EE9AD4' }
                ];
                
                const nodes = data.nodes.map((node, idx) => {
                    const catIdx = node.category || 0;
                    const colors = categoryColors[catIdx % categoryColors.length];
                    return {
                        ...node,
                        symbolSize: node.symbolSize || 45,
                        itemStyle: {
                            color: isDark ? colors.dark : colors.light,
                            borderColor: isDark ? '#ffffff' : '#ffffff',
                            borderWidth: 2,
                            shadowBlur: 10,
                            shadowColor: 'rgba(0,0,0,0.3)'
                        },
                        label: {
                            show: true,
                            fontSize: 12,
                            fontWeight: '500',
                            color: isDark ? '#EFF3F8' : '#1A2C3E',
                            offset: [8, 0]
                        }
                    };
                });
                
                const links = data.links.map(link => ({
                    ...link,
                    lineStyle: {
                        color: isDark ? '#4A5A72' : '#CBD5E1',
                        width: 2,
                        curveness: 0.3,
                        type: 'solid',
                        opacity: 0.7
                    },
                    label: {
                        show: true,
                        fontSize: 10,
                        color: isDark ? '#9CA3AF' : '#6B7280',
                        backgroundColor: isDark ? 'rgba(17,22,31,0.8)' : 'rgba(255,255,255,0.8)',
                        padding: [2, 5, 2, 5],
                        borderRadius: 4,
                        offset: [0, -10]
                    },
                    emphasis: {
                        lineStyle: {
                            width: 3,
                            color: isDark ? '#819BBB' : '#2C3E66',
                            opacity: 1
                        }
                    }
                }));
                
                const categories = (data.categories || []).map((cat, idx) => {
                    const colors = categoryColors[idx % categoryColors.length];
                    return {
                        name: cat.name,
                        itemStyle: {
                            color: isDark ? colors.dark : colors.light
                        }
                    };
                });
                
                currentChart.setOption({
                    title: {
                        text: '文档实体关系图谱',
                        left: 'center',
                        top: 5,
                        textStyle: {
                            color: isDark ? '#EFF3F8' : '#1A2C3E',
                            fontSize: 14,
                            fontWeight: '600'
                        }
                    },
                    tooltip: { trigger: 'item' },
                    series: [{
                        type: 'graph',
                        layout: 'force',
                        force: { repulsion: 500, edgeLength: 150, gravity: 0.1, layoutAnimation: true },
                        roam: true,
                        draggable: true,
                        edgeSymbol: ['none', 'arrow'],
                        edgeSymbolSize: [0, 8],
                        label: {
                            show: true,
                            position: 'right',
                            fontSize: 11,
                            offset: [5, 0],
                            color: isDark ? '#EFF3F8' : '#1A2C3E',
                            fontWeight: '500'
                        },
                        edgeLabel: {
                            show: true,
                            fontSize: 10,
                            position: 'middle',
                            formatter: (params) => params.data.name || '',
                            color: isDark ? '#9CA3AF' : '#6B7280',
                            backgroundColor: isDark ? 'rgba(17,22,31,0.7)' : 'rgba(255,255,255,0.7)',
                            padding: [2, 6, 2, 6],
                            borderRadius: 12
                        },
                        data: nodes,
                        links: links,
                        categories: categories,
                        lineStyle: { color: isDark ? '#4A5A72' : '#CBD5E1', width: 2, curveness: 0.3, opacity: 0.6 },
                        emphasis: { scale: true, focus: 'adjacency', lineStyle: { width: 3, color: isDark ? '#819BBB' : '#2C3E66', opacity: 1 } },
                        animation: true,
                        animationDuration: 800
                    }]
                });
                
                currentChart.on('click', (params) => {
                    if (params.dataType === 'node') {
                        const node = params.data;
                        const categoryName = categories[node.category]?.name || '实体';
                        document.getElementById('graphDetail').innerHTML = `
                            <div style="display: flex; flex-direction: column; gap: 12px;">
                                <div style="border-bottom: 1px solid var(--border-light); padding-bottom: 10px;">
                                    <h3 style="margin: 0; font-size: 1rem;"><i data-lucide="star" style="width: 16px; height: 16px;"></i> ${node.name}</h3>
                                </div>
                                <div><strong>类型：</strong> ${categoryName}</div>
                                <div><strong>重要性：</strong> ${node.symbolSize || 40}</div>
                                <div><strong>来源文档：</strong> ${files.map(f => f.name).join(', ')}</div>
                            </div>
                        `;
                        lucide.createIcons();
                    }
                });
            }
            
        } catch (error) {
            console.error('图谱加载失败:', error);
            document.getElementById('graphDetail').innerHTML = `
                <div style="text-align: center; color: var(--danger); padding: 40px;">
                    <i data-lucide="x-circle" style="width: 32px; height: 32px;"></i><br>
                    ${error.message}<br>
                    <small style="color: var(--text-placeholder);">请检查文档内容或稍后重试</small>
                </div>
            `;
            lucide.createIcons();
        }
    };
    
    setTimeout(loadGraph, 200);
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