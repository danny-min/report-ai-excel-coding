const { contextBridge, ipcRenderer } = require('electron');

// 暴露安全的 API 给渲染进程
contextBridge.exposeInMainWorld('electronAPI', {
    // API 请求
    apiRequest: (method, endpoint, data, isFormData = false) => {
        return ipcRenderer.invoke('api-request', { method, endpoint, data, isFormData });
    },
    
    // 选择文件
    selectFile: () => ipcRenderer.invoke('select-file'),
    
    // 保存文件
    saveFile: (defaultName) => ipcRenderer.invoke('save-file', defaultName),
    
    // 窗口控制
    windowMinimize: () => ipcRenderer.invoke('window-minimize'),
    windowMaximize: () => ipcRenderer.invoke('window-maximize'),
    windowClose: () => ipcRenderer.invoke('window-close'),
    windowIsMaximized: () => ipcRenderer.invoke('window-is-maximized'),
    
    // 关闭对话框结果
    closeDialogResult: (choice) => ipcRenderer.invoke('close-dialog-result', choice),
    
    // 监听显示关闭对话框事件
    onShowCloseDialog: (callback) => {
        ipcRenderer.on('show-close-dialog', callback);
    },
    
    // 设置后端端口
    setBackendPort: (port) => ipcRenderer.invoke('set-backend-port', port),
    
    // 获取后端端口
    getBackendPort: () => ipcRenderer.invoke('get-backend-port'),
    
    // 重启应用
    restartApp: () => ipcRenderer.invoke('restart-app'),
    
    // 强制清理后台进程
    forceCleanup: () => ipcRenderer.invoke('force-cleanup'),
    
    // 诊断端口问题
    diagnosePort: () => ipcRenderer.invoke('diagnose-port'),
});









































