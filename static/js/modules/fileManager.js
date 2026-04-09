// 全局文件管理
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
    
    render() {
        if (!this.fileListContainer || !this.fileListBody) return;
        
        if (this.files.length === 0) {
            this.fileListContainer.classList.add('hidden');
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
            item.innerHTML = `
                <div class="file-info">
                    <span style="font-size: 20px;">${Utils.getFileIcon(file.name)}</span>
                    <div>
                        <div style="font-size: 0.85rem;">${file.name}</div>
                        <div style="font-size: 0.7rem; color: var(--text-placeholder);">${Utils.formatFileSize(file.size)}</div>
                    </div>
                </div>
                <button class="file-remove" data-index="${idx}">✕</button>
            `;
            this.fileListBody.appendChild(item);
        });
        
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
        });
        
        dropArea.addEventListener('dragleave', () => {
            dropArea.style.borderColor = '';
        });
        
        dropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            dropArea.style.borderColor = '';
            const files = Array.from(e.dataTransfer.files);
            this.addFiles(files);
        });
        
        dropArea.addEventListener('click', () => {
            if (this.fileInput) this.fileInput.click();
        });
    }
}