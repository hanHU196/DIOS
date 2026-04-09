// 工作区主逻辑
let currentFileManager = null;
let currentFunction = null;
let currentChart = null;

// 页面跳转
document.addEventListener('DOMContentLoaded', () => {
    // 返回首页 - 使用正确的路由
    const backHomeBtn = document.getElementById('backHomeBtn');
    if (backHomeBtn) {
        backHomeBtn.addEventListener('click', () => {
            window.location.href = '/';
        });
    }
    
    // 初始化文件管理器
    initFileManager();
    
    // 绑定功能卡片点击事件
    bindFunctionCards();
});

// 初始化文件管理器
function initFileManager() {
    currentFileManager = new FileManager();
    currentFileManager.init({
        fileInputId: 'globalFileInput',
        listContainerId: 'globalFileList',
        listBodyId: 'globalFileListBody',
        countSpanId: 'globalFileCount',
        onChange: (files) => {
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

// 表格填写 UI
function renderFillUI(container) {
    container.innerHTML = `
        <div class="form-group">
            <label class="form-label">Excel 模板</label>
            <input type="file" id="templateFile" accept=".xlsx,.xls" style="padding: 8px;">
        </div>
        <div class="form-group">
            <label class="form-label">字段指令</label>
            <input type="text" id="fillCommand" class="input" placeholder="甲方、乙方、金额">
        </div>
        <button class="btn-primary" id="executeFill">
            生成并下载
        </button>
        <div id="fillResultArea" style="margin-top: 20px;"></div>
        <div id="fillError" class="error-message hidden"></div>
    `;
    lucide.createIcons();
    
    document.getElementById('executeFill').addEventListener('click', async () => {
        const command = document.getElementById('fillCommand').value.trim();
        const template = document.getElementById('templateFile').files[0];
        
        if (!command) { Utils.showError('fillError', '请输入字段指令'); return; }
        if (!template) { Utils.showError('fillError', '请上传模板文件'); return; }
        
        const files = currentFileManager.getFiles();
        if (files.length === 0) { Utils.showError('fillError', '请先上传源文档'); return; }
        
        const area = document.getElementById('fillResultArea');
        area.innerHTML = '<div class="progress-container"><div class="progress-bar-bg"><div class="progress-bar-fill" style="width: 50%"></div></div><div><i data-lucide="loader" style="width: 16px; height: 16px; animation: spin 1s linear infinite;"></i> 处理中...</div></div>';
        lucide.createIcons();
        
        const formData = new FormData();
        for (const file of files) {
            formData.append('documents', file);
        }
        formData.append('template', template);
        formData.append('command', command);
        
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
            
            area.innerHTML = `
                <div style="text-align: center; padding: 20px;">
                    <a href="${url}" class="btn-secondary" download="filled_template.xlsx" style="display: inline-block; padding: 10px 24px; text-decoration: none;">
                        <i data-lucide="download" style="width: 14px; height: 14px;"></i> 下载表格文件
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

// 格式调整 UI
function renderFormatUI(container) {
    container.innerHTML = `
        <div class="form-group">
            <label class="form-label">格式指令</label>
            <input type="text" id="formatCommand" class="input" placeholder="将所有段落居中、标题加粗">
            <div style="font-size: 12px; color: var(--text-placeholder); margin-top: 6px;">
                <i data-lucide="lightbulb" style="width: 12px; height: 12px; display: inline;"></i> 支持：对齐、加粗、字体、字号等
            </div>
        </div>
        <button class="btn-primary" id="executeFormat">
            应用格式
        </button>
        <div id="formatResultArea" style="margin-top: 20px;"></div>
        <div id="formatError" class="error-message hidden"></div>
    `;
    lucide.createIcons();
    
    document.getElementById('executeFormat').addEventListener('click', async () => {
        const command = document.getElementById('formatCommand').value.trim();
        if (!command) { Utils.showError('formatError', '请输入格式指令'); return; }
        
        const files = currentFileManager.getFiles();
        if (files.length === 0) { Utils.showError('formatError', '请先上传文档'); return; }
        
        const area = document.getElementById('formatResultArea');
        area.innerHTML = '<div class="progress-container"><div class="progress-bar-bg"><div class="progress-bar-fill" style="width: 50%"></div></div><div><i data-lucide="loader" style="width: 16px; height: 16px; animation: spin 1s linear infinite;"></i> 处理中...</div></div>';
        lucide.createIcons();
        
        const formData = new FormData();
        formData.append('document', files[0]);
        formData.append('command', command);
        
        try {
            const response = await fetch('/api/format', {
                method: 'POST',
                body: formData
            });
            
            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || '格式调整失败');
            }
            
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            
            area.innerHTML = `
                <div style="text-align: center; padding: 20px;">
                    <a href="${url}" class="btn-secondary" download="formatted_document.docx" style="display: inline-block; padding: 10px 24px; text-decoration: none;">
                        <i data-lucide="download" style="width: 14px; height: 14px;"></i> 下载调整后文档
                    </a>
                </div>
            `;
            lucide.createIcons();
            
        } catch (error) {
            Utils.showError('formatError', error.message);
            area.innerHTML = `<div class="error-message"><i data-lucide="alert-circle" style="width: 16px; height: 16px;"></i> ${error.message}</div>`;
            lucide.createIcons();
        }
    });
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