// 文件管理模块
class FileManager {
    constructor() {
        this.files = [];
        this.fileInput = null;
        this.fileListContainer = null;
        this.fileListBody = null;
        this.fileCountSpan = null;
        this.onChangeCallback = null;
    }
    
    init(config) {
        this.fileInput = document.getElementById(config.fileInputId);
        this.fileListContainer = document.getElementById(config.listContainerId);
        this.fileListBody = document.getElementById(config.listBodyId);
        this.fileCountSpan = document.getElementById(config.countSpanId);
        this.onChangeCallback = config.onChange || null;
        
        this.bindEvents();
        this.render();
    }
    
    bindEvents() {
        if (!this.fileInput) return;
        
        this.fileInput.addEventListener('change', (e) => {
            const newFiles = Array.from(e.target.files);
            this.addFiles(newFiles);
        });
        
        const clearBtn = document.getElementById('clearAllFiles');
        if (clearBtn) {
            clearBtn.addEventListener('click', () => this.clearFiles());
        }
    }
    
    addFiles(newFiles) {
        for (const file of newFiles) {
            const exists = this.files.some(f => 
                f.name === file.name && f.size === file.size
            );
            if (!exists) {
                this.files.push(file);
            }
        }
        this.updateFileInput();
        this.render();
        if (this.onChangeCallback) this.onChangeCallback(this.files);
    }
    
    removeFile(index) {
        this.files.splice(index, 1);
        this.updateFileInput();
        this.render();
        if (this.onChangeCallback) this.onChangeCallback(this.files);
    }
    
    clearFiles() {
        this.files = [];
        this.updateFileInput();
        this.render();
        if (this.onChangeCallback) this.onChangeCallback(this.files);
    }
    
    updateFileInput() {
        if (!this.fileInput) return;
        const dt = new DataTransfer();
        this.files.forEach(f => dt.items.add(f));
        this.fileInput.files = dt.files;
    }
    
    getFiles() {
        return [...this.files];
    }
    
    formatFileSize(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
    
    // ========== 添加这个方法来防止 XSS 攻击 ==========
    escapeHtml(str) {
        if (!str) return '';
        return str.replace(/[&<>]/g, function(m) {
            if (m === '&') return '&amp;';
            if (m === '<') return '&lt;';
            if (m === '>') return '&gt;';
            return m;
        });
    }
    // ================================================
    
    render() {
        if (!this.fileListContainer || !this.fileListBody) return;
        
        if (this.files.length === 0) {
            this.fileListContainer.classList.add('hidden');
            if (this.fileCountSpan) {
                this.fileCountSpan.textContent = '0';
            }
            return;
        }
        
        this.fileListContainer.classList.remove('hidden');
        if (this.fileCountSpan) {
            this.fileCountSpan.textContent = this.files.length;
        }
        
        this.fileListBody.innerHTML = '';
        this.files.forEach((file, idx) => {
            const item = document.createElement('div');
            item.className = 'file-item';
            const iconHtml = typeof Utils !== 'undefined' && Utils.getFileIcon 
                ? Utils.getFileIcon(file.name) 
                : `<span style="font-size: 20px;">📄</span>`;
            
            item.innerHTML = `
                <div class="file-info">
                    ${iconHtml}
                    <div>
                        <div style="font-size: 0.85rem; font-weight: 500;">${this.escapeHtml(file.name)}</div>
                        <div style="font-size: 0.7rem; color: var(--text-placeholder);">${this.formatFileSize(file.size)}</div>
                    </div>
                </div>
                <button class="file-remove" data-index="${idx}">✕</button>
            `;
            this.fileListBody.appendChild(item);
        });
        
        // 初始化 Lucide 图标
        if (typeof lucide !== 'undefined') {
            lucide.createIcons();
        }
        
        this.fileListBody.querySelectorAll('.file-remove').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = parseInt(btn.dataset.index);
                this.removeFile(idx);
            });
        });
    }
    
    setupDragDrop(dropAreaId) {
        const dropArea = document.getElementById(dropAreaId);
        if (!dropArea) return;
        
        dropArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropArea.style.borderColor = 'var(--accent-solid)';
            dropArea.style.backgroundColor = 'var(--accent-soft)';
        });
        
        dropArea.addEventListener('dragleave', () => {
            dropArea.style.borderColor = '';
            dropArea.style.backgroundColor = '';
        });
        
        dropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            dropArea.style.borderColor = '';
            dropArea.style.backgroundColor = '';
            const files = Array.from(e.dataTransfer.files);
            this.addFiles(files);
        });
        
        dropArea.addEventListener('click', () => {
            if (this.fileInput) this.fileInput.click();
        });
    }
}