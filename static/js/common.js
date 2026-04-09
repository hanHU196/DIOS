// ========== 简化版主题管理 ==========
(function() {
    // 初始化主题
    const savedTheme = localStorage.getItem('theme');
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    
    if (savedTheme === 'dark' || (!savedTheme && prefersDark)) {
        document.documentElement.classList.add('dark');
    } else {
        document.documentElement.classList.remove('dark');
    }
    
    // 更新图标
    function updateIcon() {
        const isDark = document.documentElement.classList.contains('dark');
        const icons = document.querySelectorAll('#themeIcon, #themeIconNav');
        icons.forEach(icon => {
            if (icon) {
                icon.setAttribute('data-lucide', isDark ? 'sun' : 'moon');
            }
        });
        if (typeof lucide !== 'undefined') {
            lucide.createIcons();
        }
    }
    
    // 切换主题
    function toggleTheme() {
        const isDark = document.documentElement.classList.contains('dark');
        if (isDark) {
            document.documentElement.classList.remove('dark');
            localStorage.setItem('theme', 'light');
        } else {
            document.documentElement.classList.add('dark');
            localStorage.setItem('theme', 'dark');
        }
        updateIcon();
    }
    
    // 绑定按钮
    function bindButtons() {
        const buttons = document.querySelectorAll('.theme-toggle, .theme-toggle-nav, #themeToggle, #themeToggleNav');
        buttons.forEach(btn => {
            btn.onclick = toggleTheme;
        });
    }
    
    // 页面加载完成后初始化
    document.addEventListener('DOMContentLoaded', () => {
        updateIcon();
        bindButtons();
    });
})();

// ========== 工具函数 ==========
const Utils = {
    formatFileSize(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    },
    
    // 获取文件图标（返回 Lucide 线性图标）
    getFileIcon(fileName) {
        const ext = fileName.split('.').pop().toLowerCase();
        
        // 返回 Lucide 图标名称和对应的样式
        const iconMap = {
            txt: { name: 'file-text', color: '#6B7280' },
            docx: { name: 'file-text', color: '#2B5797' },
            md: { name: 'file-text', color: '#083344' },
            xlsx: { name: 'table', color: '#217346' },
            xls: { name: 'table', color: '#217346' },
            pdf: { name: 'file', color: '#E74C3C' },
            pptx: { name: 'presentation', color: '#D24726' },
            jpg: { name: 'image', color: '#8B5CF6' },
            png: { name: 'image', color: '#8B5CF6' },
            zip: { name: 'archive', color: '#F59E0B' }
        };
        
        const defaultIcon = { name: 'file', color: '#9AA9BF' };
        const icon = iconMap[ext] || defaultIcon;
        
        // 返回 Lucide 图标 HTML
        return `<i data-lucide="${icon.name}" style="width: 20px; height: 20px; stroke: ${icon.color}; stroke-width: 1.5;"></i>`;
    },
    
    // 渲染文件列表后需要重新初始化 Lucide 图标
    renderFileListWithIcons() {
        if (typeof lucide !== 'undefined') {
            lucide.createIcons();
        }
    },
    
    showError(elementId, message, timeout = 3000) {
        const el = document.getElementById(elementId);
        if (el) {
            el.textContent = message;
            el.classList.remove('hidden');
            setTimeout(() => el.classList.add('hidden'), timeout);
        }
    },
    
    exportToCSV(records, fields, filename = 'export') {
        let csv = fields.join(',') + '\n';
        records.forEach(record => {
            const row = fields.map(f => {
                let val = record[f] || '';
                val = String(val).replace(/"/g, '""');
                return `"${val}"`;
            }).join(',');
            csv += row + '\n';
        });
        
        const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${filename}_${Date.now()}.csv`;
        a.click();
        URL.revokeObjectURL(url);
    }
};