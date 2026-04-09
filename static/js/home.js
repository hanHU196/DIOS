// 首页交互逻辑
document.addEventListener('DOMContentLoaded', () => {
    // 跳转到工作区 - 使用 /workspace 路由
    const goToWorkspace = () => {
        window.location.href = '/workspace';
    };
    
    // 绑定所有跳转按钮
    const getStartedBtn = document.getElementById('getStartedBtn');
    const heroStartBtn = document.getElementById('heroStartBtn');
    const ctaStartBtn = document.getElementById('ctaStartBtn');
    
    if (getStartedBtn) getStartedBtn.addEventListener('click', goToWorkspace);
    if (heroStartBtn) heroStartBtn.addEventListener('click', goToWorkspace);
    if (ctaStartBtn) ctaStartBtn.addEventListener('click', goToWorkspace);
    
    // 登录按钮
    const loginBtn = document.getElementById('loginBtn');
    if (loginBtn) {
        loginBtn.addEventListener('click', () => {
            alert('演示模式：登录功能即将开放');
        });
    }
    
    // 演示按钮
    const demoBtn = document.getElementById('demoBtn');
    if (demoBtn) {
        demoBtn.addEventListener('click', () => {
            alert('演示视频即将上线');
        });
    }
});