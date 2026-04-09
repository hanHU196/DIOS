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
                        ⚠️ 请先上传文档
                    </div>
                `;
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
                <button class="file-remove" data-index="${idx}">✕</button>
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
        extract: '📄 智能文档提取',
        fill: '📊 表格智能填写',
        format: '🎨 文档格式调整',
        qa: '💬 智能问答',
        graph: '🕸️ 知识图谱'
    };
    resultTitle.innerHTML = titles[func] || '执行结果';
    
    const files = currentFileManager.getFiles();
    if (files.length === 0) {
        resultContent.innerHTML = '<div style="text-align: center; padding: 60px; color: var(--warning);">⚠️ 请先上传文档</div>';
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

// 智能提取 UI
function renderExtractUI(container) {
    container.innerHTML = `
        <div class="form-group">
            <label class="form-label">提取指令</label>
            <input type="text" id="extractCommand" class="input" placeholder="例如：甲方、乙方、合同金额">
            <div style="font-size: 12px; color: var(--text-placeholder); margin-top: 6px;">💡 支持格式：提取甲方、乙方、金额</div>
        </div>
        <button class="btn-primary" id="executeExtract">开始提取</button>
        <div id="extractResultArea" style="margin-top: 20px;"></div>
        <div id="extractError" class="error-message hidden"></div>
    `;
    
    document.getElementById('executeExtract').addEventListener('click', async () => {
        const command = document.getElementById('extractCommand').value.trim();
        if (!command) {
            Utils.showError('extractError', '请输入提取指令');
            return;
        }
        
        const area = document.getElementById('extractResultArea');
        area.innerHTML = '<div class="progress-container"><div class="progress-bar-bg"><div class="progress-bar-fill" style="width: 50%"></div></div><div>处理中...</div></div>';
        
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
                area.innerHTML = '<div style="text-align: center; padding: 40px; color: var(--warning);">未提取到数据，请检查文档内容或指令</div>';
                return;
            }
            
            let html = `<table class="data-table"><thead><tr>${fields.map(f => `<th>${f}</th>`).join('')}</tr></thead><tbody>`;
            records.forEach(r => {
                html += '<tr>' + fields.map(f => `<td contenteditable="true">${r[f] || '<span style="color: var(--text-placeholder);">无数据</span>'}</td>`).join('') + '</tr>';
            });
            html += `</tbody></table>
                <div style="margin-top: 16px;">
                    <button class="btn-secondary" id="exportTableBtn">📥 导出 CSV</button>
                </div>`;
            area.innerHTML = html;
            
            document.getElementById('exportTableBtn')?.addEventListener('click', () => {
                Utils.exportToCSV(records, fields, '提取结果');
            });
            
        } catch (error) {
            console.error('提取失败:', error);
            Utils.showError('extractError', error.message);
            area.innerHTML = `<div class="error-message">${error.message}</div>`;
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
        <button class="btn-primary" id="executeFill">生成并下载</button>
        <div id="fillResultArea" style="margin-top: 20px;"></div>
        <div id="fillError" class="error-message hidden"></div>
    `;
    
    document.getElementById('executeFill').addEventListener('click', async () => {
        const command = document.getElementById('fillCommand').value.trim();
        const template = document.getElementById('templateFile').files[0];
        
        if (!command) { Utils.showError('fillError', '请输入字段指令'); return; }
        if (!template) { Utils.showError('fillError', '请上传模板文件'); return; }
        
        const files = currentFileManager.getFiles();
        if (files.length === 0) { Utils.showError('fillError', '请先上传源文档'); return; }
        
        const area = document.getElementById('fillResultArea');
        area.innerHTML = '<div class="progress-container"><div class="progress-bar-bg"><div class="progress-bar-fill" style="width: 50%"></div></div><div>处理中...</div></div>';
        
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
                        📥 下载表格文件
                    </a>
                </div>
            `;
            
        } catch (error) {
            Utils.showError('fillError', error.message);
            area.innerHTML = `<div class="error-message">${error.message}</div>`;
        }
    });
}

// 格式调整 UI
function renderFormatUI(container) {
    container.innerHTML = `
        <div class="form-group">
            <label class="form-label">格式指令</label>
            <input type="text" id="formatCommand" class="input" placeholder="将所有段落居中、标题加粗">
            <div style="font-size: 12px; color: var(--text-placeholder); margin-top: 6px;">💡 支持：对齐、加粗、字体、字号等</div>
        </div>
        <button class="btn-primary" id="executeFormat">应用格式</button>
        <div id="formatResultArea" style="margin-top: 20px;"></div>
        <div id="formatError" class="error-message hidden"></div>
    `;
    
    document.getElementById('executeFormat').addEventListener('click', async () => {
        const command = document.getElementById('formatCommand').value.trim();
        if (!command) { Utils.showError('formatError', '请输入格式指令'); return; }
        
        const files = currentFileManager.getFiles();
        if (files.length === 0) { Utils.showError('formatError', '请先上传文档'); return; }
        
        const area = document.getElementById('formatResultArea');
        area.innerHTML = '<div class="progress-container"><div class="progress-bar-bg"><div class="progress-bar-fill" style="width: 50%"></div></div><div>处理中...</div></div>';
        
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
                        📥 下载调整后文档
                    </a>
                </div>
            `;
            
        } catch (error) {
            Utils.showError('formatError', error.message);
            area.innerHTML = `<div class="error-message">${error.message}</div>`;
        }
    });
}

// 智能问答 UI
function renderQaUI(container) {
    container.innerHTML = `
        <div class="chat-messages" id="chatMessages">
            <div class="message assistant">
                <div class="message-avatar">🤖</div>
                <div class="message-content">您好！已加载 ${currentFileManager.getFiles().length} 个文档，请问有什么可以帮助您？</div>
            </div>
        </div>
        <div style="display: flex; gap: 12px;">
            <input type="text" id="qaInput" class="input" placeholder="输入问题...">
            <button class="btn-primary" id="sendQa">发送</button>
        </div>
        <div style="display: flex; gap: 8px; margin-top: 12px; flex-wrap: wrap;">
            <button class="btn-secondary quick-q" style="padding: 6px 12px;">总结文档内容</button>
            <button class="btn-secondary quick-q" style="padding: 6px 12px;">有哪些关键实体？</button>
            <button class="btn-secondary quick-q" style="padding: 6px 12px;">文档的主要观点是什么？</button>
        </div>
    `;
    
    const chatMessages = document.getElementById('chatMessages');
    const addMessage = (role, content) => {
        const msg = document.createElement('div');
        msg.className = `message ${role}`;
        msg.innerHTML = `<div class="message-avatar">${role === 'user' ? '👤' : '🤖'}</div><div class="message-content">${content}</div>`;
        chatMessages.appendChild(msg);
        chatMessages.scrollTop = chatMessages.scrollHeight;
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
            askQuestion(btn.textContent);
        });
    });
}

// 知识图谱 UI
function renderGraphUI(container) {
    container.innerHTML = `
        <div class="graph-container">
            <div class="graph-canvas" id="graphCanvas"></div>
            <div class="graph-detail" id="graphDetail">点击节点查看详情</div>
        </div>
    `;
    
    const loadGraph = async () => {
        const files = currentFileManager.getFiles();
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
                throw new Error('图谱生成失败');
            }
            
            const data = await response.json();
            
            const canvas = document.getElementById('graphCanvas');
            if (canvas && typeof echarts !== 'undefined') {
                if (currentChart) currentChart.dispose();
                currentChart = echarts.init(canvas);
                currentChart.setOption({
                    series: [{
                        type: 'graph',
                        layout: 'force',
                        symbolSize: 50,
                        roam: true,
                        label: { show: true, position: 'right', fontSize: 12, color: 'var(--text-primary)' },
                        data: data.nodes || [],
                        links: data.links || [],
                        categories: data.categories || [],
                        lineStyle: { color: 'var(--text-placeholder)', curveness: 0.3 },
                        emphasis: { focus: 'adjacency' }
                    }]
                });
                
                currentChart.on('click', (params) => {
                    if (params.dataType === 'node') {
                        document.getElementById('graphDetail').innerHTML = `
                            <strong>${params.name}</strong><br>
                            类型：${params.data.category !== undefined ? (data.categories?.[params.data.category]?.name || '实体') : '实体'}<br>
                            来源文档：${files.map(f => f.name).join(', ')}
                        `;
                    }
                });
            }
            
        } catch (error) {
            console.error('图谱加载失败:', error);
            document.getElementById('graphDetail').innerHTML = `
                <div style="text-align: center; color: var(--danger);">
                    图谱生成失败：${error.message}
                </div>
            `;
        }
    };
    
    setTimeout(loadGraph, 100);
}