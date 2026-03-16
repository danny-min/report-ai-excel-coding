/**
 * Electron 主进程
 * AI 报表生成器
 */

'use strict';

// 首先安装 EPIPE 防护（必须在最前面）
const safeConsole = require('./safe-console');
safeConsole.initialize();

const { app, BrowserWindow, ipcMain, dialog, Tray, Menu, nativeImage } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');

// 全局变量
let mainWindow = null;
let pythonProcess = null;
let pythonPort = 5000;  // 默认端口，如果被占用会自动切换
let tray = null;
let isQuitting = false;  // 是否真正退出（而不是最小化到托盘）

/**
 * 检查端口是否可用
 */
function isPortAvailable(port) {
    return new Promise((resolve) => {
        const net = require('net');
        const server = net.createServer();
        
        server.once('error', (err) => {
            if (err.code === 'EADDRINUSE') {
                resolve(false);
            } else {
                resolve(false);
            }
        });
        
        server.once('listening', () => {
            server.close();
            resolve(true);
        });
        
        server.listen(port, '127.0.0.1');
    });
}

/**
 * 查找可用端口
 */
async function findAvailablePort(startPort = 5000, maxAttempts = 10) {
    for (let i = 0; i < maxAttempts; i++) {
        const port = startPort + i;
        const available = await isPortAvailable(port);
        if (available) {
            return port;
        }
        console.log(`[Main] 端口 ${port} 已被占用，尝试下一个...`);
    }
    return startPort; // 如果都不可用，返回默认端口让后端报错
}

// 是否为开发模式
const isDev = process.argv.includes('--dev');

/**
 * 获取 Python 后端路径
 */
function getPythonBackendPath() {
    if (isDev) {
        // 开发模式：使用源码目录
        return path.join(__dirname, '..', '..', 'python-backend');
    } else {
        // 生产模式：使用打包后的目录
        return path.join(process.resourcesPath, 'python-backend');
    }
}

/**
 * 清理残留的后端进程
 */
function cleanupOldBackendProcesses() {
    return new Promise((resolve) => {
        if (process.platform === 'win32') {
            // Windows: 先杀死所有可能残留的 report-backend.exe 进程
            const cleanup = spawn('taskkill', ['/F', '/IM', 'report-backend.exe'], {
                stdio: 'ignore',
                windowsHide: true
            });
            cleanup.on('close', () => {
                console.log('[Main] 已清理残留的后端进程');
                // 等待一小段时间确保端口释放
                setTimeout(resolve, 500);
            });
            cleanup.on('error', () => {
                resolve(); // 忽略错误，可能没有残留进程
            });
        } else {
            // 非 Windows 平台
            const cleanup = spawn('pkill', ['-f', 'report-backend'], {
                stdio: 'ignore'
            });
            cleanup.on('close', () => {
                setTimeout(resolve, 500);
            });
            cleanup.on('error', () => {
                resolve();
            });
        }
    });
}

/**
 * 启动 Python 后端
 */
function startPythonBackend() {
    return new Promise(async (resolve, reject) => {
        // 先清理可能残留的旧进程
        await cleanupOldBackendProcesses();
        
        const backendPath = getPythonBackendPath();
        
        let pythonExecutable;
        let args;

        if (isDev) {
            // 开发模式：使用 python 命令
            pythonExecutable = 'python';
            args = [path.join(backendPath, 'app.py')];
        } else {
            // 生产模式：使用打包的 exe
            if (process.platform === 'win32') {
                pythonExecutable = path.join(backendPath, 'report-backend.exe');
            } else {
                pythonExecutable = path.join(backendPath, 'report-backend');
            }
            args = [];
        }

        console.log(`[Main] 启动 Python 后端: ${pythonExecutable}`);
        console.log(`[Main] 后端目录: ${backendPath}`);
        console.log(`[Main] resourcesPath: ${process.resourcesPath}`);
        console.log(`[Main] isDev: ${isDev}`);
        
        // 列出后端目录内容（用于调试）
        try {
            if (fs.existsSync(backendPath)) {
                const files = fs.readdirSync(backendPath);
                console.log(`[Main] 后端目录内容: ${files.join(', ')}`);
            } else {
                console.log(`[Main] 后端目录不存在: ${backendPath}`);
            }
        } catch (e) {
            console.log(`[Main] 无法读取后端目录: ${e.message}`);
        }

        // 检查文件是否存在
        const execPath = isDev ? path.join(backendPath, 'app.py') : pythonExecutable;
        if (!fs.existsSync(execPath)) {
            console.warn(`[Main] Python 后端文件不存在: ${execPath}`);
            // 不阻塞启动，继续运行
            resolve();
            return;
        }

        pythonProcess = spawn(pythonExecutable, args, {
            cwd: backendPath,
            env: { 
                ...process.env, 
                PYTHONIOENCODING: 'utf-8',
                BACKEND_PORT: pythonPort.toString()  // 传递端口配置给后端
            },
            stdio: ['pipe', 'pipe', 'pipe'],
        });

        // 使用安全的子进程包装
        safeConsole.createSafeChildProcess(pythonProcess);

        pythonProcess.stdout.on('data', (data) => {
            const output = data.toString();
            console.log(`[Python] ${output}`);
            
            // 检测服务启动成功
            if (output.includes('Running on') || output.includes('Uvicorn running')) {
                console.log('[Main] Python 后端启动成功');
                resolve();
            }
        });

        pythonProcess.stderr.on('data', (data) => {
            console.warn(`[Python Error] ${data.toString()}`);
        });

        pythonProcess.on('error', (err) => {
            console.error(`[Main] Python 进程错误: ${err.message}`);
            resolve(); // 不阻塞，继续运行
        });

        pythonProcess.on('close', (code) => {
            console.log(`[Main] Python 进程退出，代码: ${code}`);
            pythonProcess = null;
        });

        // 3 秒超时
        setTimeout(() => {
            resolve();
        }, 3000);
    });
}

/**
 * 停止 Python 后端
 */
function stopPythonBackend() {
    console.log('[Main] 停止 Python 后端');
    
    if (process.platform === 'win32') {
        // Windows 下使用 taskkill 杀死所有 report-backend.exe 进程
        // 这样可以确保清理所有残留进程，不仅仅是当前启动的
        spawn('taskkill', ['/F', '/IM', 'report-backend.exe'], {
            stdio: 'ignore',
            windowsHide: true
        });
        
        // 如果有记录的进程 PID，也尝试通过 PID 杀死
        if (pythonProcess && pythonProcess.pid) {
            spawn('taskkill', ['/pid', pythonProcess.pid.toString(), '/f', '/t'], {
                stdio: 'ignore',
                windowsHide: true
            });
        }
    } else {
        // 非 Windows 平台
        spawn('pkill', ['-f', 'report-backend'], { stdio: 'ignore' });
        if (pythonProcess) {
            pythonProcess.kill('SIGTERM');
        }
    }
    
    pythonProcess = null;
}

/**
 * 创建主窗口
 */
function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1400,
        height: 900,
        minWidth: 1000,
        minHeight: 700,
        frame: false,           // 无边框窗口
        titleBarStyle: 'hidden', // 隐藏标题栏
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js'),
        },
        icon: path.join(__dirname, '..', 'renderer', 'assets', 'icon.ico'),
        show: false,
    });

    // 加载页面
    const indexPath = path.join(__dirname, '..', 'renderer', 'index.html');
    mainWindow.loadFile(indexPath);

    // 窗口准备好后最大化并显示
    mainWindow.once('ready-to-show', () => {
        mainWindow.maximize();  // 默认最大化
        mainWindow.show();
    });

    // 开发模式打开开发者工具
    if (isDev) {
        mainWindow.webContents.openDevTools();
    }

    // 点击关闭按钮时最小化到托盘（除非是真正退出）
    mainWindow.on('close', (event) => {
        if (!isQuitting) {
            event.preventDefault();
            mainWindow.hide();
            
            // 显示托盘气泡提示（仅首次）
            if (tray && !app.isPackaged) {
                tray.displayBalloon({
                    title: 'AI 报表生成器',
                    content: '应用已最小化到系统托盘，双击图标可恢复窗口',
                });
            }
        }
    });

    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

/**
 * 创建系统托盘
 */
function createTray() {
    // 创建托盘图标
    const iconPath = path.join(__dirname, '..', 'renderer', 'assets', 'icon.ico');
    let trayIcon;
    
    if (fs.existsSync(iconPath)) {
        trayIcon = nativeImage.createFromPath(iconPath);
    } else {
        // 如果没有图标文件，创建一个简单的图标
        trayIcon = nativeImage.createEmpty();
    }
    
    tray = new Tray(trayIcon);
    tray.setToolTip('AI 报表生成器');
    
    // 托盘右键菜单
    const contextMenu = Menu.buildFromTemplate([
        {
            label: '📊 显示窗口',
            click: () => {
                if (mainWindow) {
                    mainWindow.show();
                    mainWindow.focus();
                }
            }
        },
        { type: 'separator' },
        {
            label: '❌ 退出',
            click: () => {
                isQuitting = true;
                app.quit();
            }
        }
    ]);
    
    tray.setContextMenu(contextMenu);
    
    // 双击托盘图标显示窗口
    tray.on('double-click', () => {
        if (mainWindow) {
            mainWindow.show();
            mainWindow.focus();
        }
    });
}

/**
 * 检查 Python 后端是否就绪
 */
async function waitForBackend(maxRetries = 30, delayMs = 1000) {
    for (let i = 0; i < maxRetries; i++) {
        try {
            const response = await axios.get(`http://127.0.0.1:${pythonPort}/api/health`, {
                timeout: 2000,
                proxy: false,
            });
            if (response.status === 200) {
                console.log('[Main] Python 后端健康检查通过');
                return true;
            }
        } catch (e) {
            console.log(`[Main] 等待后端启动... (${i + 1}/${maxRetries})`);
        }
        await new Promise(resolve => setTimeout(resolve, delayMs));
    }
    console.error('[Main] Python 后端启动超时');
    return false;
}

/**
 * 尝试启动或重启 Python 后端
 */
async function ensureBackendRunning() {
    // 先检查是否已经在运行
    try {
        const response = await axios.get(`http://127.0.0.1:${pythonPort}/api/health`, {
            timeout: 2000,
            proxy: false,
        });
        if (response.status === 200) {
            return true;
        }
    } catch (e) {
        // 后端未运行，尝试启动
    }

    // 尝试重新启动后端
    console.log('[Main] 尝试重新启动 Python 后端...');
    stopPythonBackend();
    await new Promise(resolve => setTimeout(resolve, 1000));
    await startPythonBackend();
    return await waitForBackend(10, 500);
}

/**
 * IPC 处理：API 请求代理
 */
ipcMain.handle('api-request', async (event, { method, endpoint, data, isFormData }) => {
    const url = `http://127.0.0.1:${pythonPort}${endpoint}`;
    
    // API 请求函数
    const makeRequest = async () => {
        let response;
        const config = {
            timeout: 300000, // 5 分钟超时（LLM 生成可能很慢）
            proxy: false,    // 禁用代理，直接连接本地后端
        };

        if (method.toUpperCase() === 'GET') {
            response = await axios.get(url, { params: data, ...config });
        } else if (method.toUpperCase() === 'POST') {
            if (isFormData && data) {
                // 处理 FormData（文件上传）
                const formData = new FormData();
                for (const key in data) {
                    if (key === 'file' && data[key].path) {
                        formData.append(key, fs.createReadStream(data[key].path), data[key].name);
                    } else {
                        formData.append(key, data[key]);
                    }
                }
                response = await axios.post(url, formData, {
                    ...config,
                    headers: formData.getHeaders(),
                });
            } else {
                response = await axios.post(url, data, config);
            }
        } else if (method.toUpperCase() === 'DELETE') {
            response = await axios.delete(url, config);
        }
        return response;
    };

    // 重试逻辑
    const maxRetries = 3;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            const response = await makeRequest();
            return { success: true, data: response?.data || {} };
        } catch (error) {
            const isConnectionError = error.code === 'ECONNREFUSED' || 
                                       error.code === 'ENOTFOUND' ||
                                       error.code === 'UNKNOWN' ||
                                       error.message.includes('UNKNOWN') ||
                                       error.message.includes('connect');
            
            console.error(`[API Error] ${endpoint} (尝试 ${attempt}/${maxRetries}):`, error.message);
            
            if (isConnectionError && attempt < maxRetries) {
                console.log('[Main] 检测到连接错误，尝试重启后端...');
                const backendReady = await ensureBackendRunning();
                if (!backendReady) {
                    return {
                        success: false,
                        error: 'Python 后端启动失败，请检查是否有其他程序占用 5000 端口，或尝试重启应用'
                    };
                }
                await new Promise(resolve => setTimeout(resolve, 500));
                continue;
            }
            
            return { 
                success: false, 
                error: error.response?.data?.error || error.message 
            };
        }
    }
    
    return { success: false, error: '请求失败，已达最大重试次数' };
});

/**
 * IPC 处理：选择文件
 */
ipcMain.handle('select-file', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile'],
        filters: [
            { name: 'Excel/CSV 文件', extensions: ['xlsx', 'xls', 'csv'] },
            { name: '所有文件', extensions: ['*'] },
        ],
    });

    if (result.canceled || result.filePaths.length === 0) {
        return null;
    }

    const filePath = result.filePaths[0];
    return {
        path: filePath,
        name: path.basename(filePath),
    };
});

/**
 * IPC 处理：保存文件
 */
ipcMain.handle('save-file', async (event, defaultName) => {
    const result = await dialog.showSaveDialog(mainWindow, {
        defaultPath: defaultName || 'result.xlsx',
        filters: [
            { name: 'Excel 文件', extensions: ['xlsx'] },
            { name: 'CSV 文件', extensions: ['csv'] },
        ],
    });

    if (result.canceled) {
        return null;
    }

    return result.filePath;
});

/**
 * IPC 处理：窗口控制
 */
ipcMain.handle('window-minimize', () => {
    if (mainWindow) mainWindow.minimize();
});

ipcMain.handle('window-maximize', () => {
    if (mainWindow) {
        if (mainWindow.isMaximized()) {
            mainWindow.unmaximize();
        } else {
            mainWindow.maximize();
        }
    }
});

ipcMain.handle('window-close', async () => {
    // 不再直接关闭，而是通知渲染进程显示自定义对话框
    if (mainWindow) {
        mainWindow.webContents.send('show-close-dialog');
    }
});

// 处理关闭对话框的选择结果
ipcMain.handle('close-dialog-result', async (event, choice) => {
    if (choice === 'minimize') {
        // 最小化到托盘
        if (mainWindow) mainWindow.hide();
    } else if (choice === 'quit') {
        // 退出程序
        isQuitting = true;
        app.quit();
    }
    // choice === 'cancel': 什么都不做
});

ipcMain.handle('window-is-maximized', () => {
    return mainWindow ? mainWindow.isMaximized() : false;
});

/**
 * 获取端口配置文件路径
 */
function getPortConfigPath() {
    return path.join(app.getPath('userData'), 'port-config.json');
}

/**
 * 读取端口配置
 */
function loadPortConfig() {
    try {
        const configPath = getPortConfigPath();
        if (fs.existsSync(configPath)) {
            const config = JSON.parse(fs.readFileSync(configPath, 'utf-8'));
            if (config.port && config.port >= 1024 && config.port <= 65535) {
                return config.port;
            }
        }
    } catch (e) {
        console.error('[Main] 读取端口配置失败:', e.message);
    }
    return 5000; // 默认端口
}

/**
 * 保存端口配置
 */
function savePortConfig(port) {
    try {
        const configPath = getPortConfigPath();
        fs.writeFileSync(configPath, JSON.stringify({ port }, null, 2));
        console.log(`[Main] 端口配置已保存: ${port}`);
        return true;
    } catch (e) {
        console.error('[Main] 保存端口配置失败:', e.message);
        return false;
    }
}

/**
 * IPC 处理：设置后端端口
 */
ipcMain.handle('set-backend-port', (event, port) => {
    if (port >= 1024 && port <= 65535) {
        savePortConfig(port);
        return { success: true, message: `端口已设置为 ${port}，重启后生效` };
    }
    return { success: false, message: '端口号无效' };
});

/**
 * IPC 处理：获取后端端口
 */
ipcMain.handle('get-backend-port', () => {
    return pythonPort;
});

/**
 * IPC 处理：重启应用
 */
ipcMain.handle('restart-app', () => {
    console.log('[Main] 正在重启应用...');
    isQuitting = true;
    stopPythonBackend();
    app.relaunch();
    app.exit(0);
});

/**
 * IPC 处理：诊断端口问题
 */
ipcMain.handle('diagnose-port', async () => {
    const { exec } = require('child_process');
    const net = require('net');
    
    const configPath = getPortConfigPath();
    const configuredPort = loadPortConfig();
    
    // 检查端口是否被占用的函数
    const checkPort = (port) => {
        return new Promise((resolve) => {
            const server = net.createServer();
            server.once('error', () => resolve({ port, inUse: true }));
            server.once('listening', () => {
                server.close();
                resolve({ port, inUse: false });
            });
            server.listen(port, '127.0.0.1');
        });
    };
    
    // 检查多个端口
    const portsToCheck = [5000, 5001, 5002, 8000, 8080, configuredPort];
    const uniquePorts = [...new Set(portsToCheck)];
    const portChecks = await Promise.all(uniquePorts.map(checkPort));
    
    // 获取占用端口的进程信息
    const getPortProcess = (port) => {
        return new Promise((resolve) => {
            exec(`netstat -ano | findstr :${port}`, { windowsHide: true }, (err, stdout) => {
                if (stdout) {
                    const lines = stdout.trim().split('\n');
                    const listening = lines.find(l => l.includes('LISTENING'));
                    if (listening) {
                        const parts = listening.trim().split(/\s+/);
                        const pid = parts[parts.length - 1];
                        exec(`tasklist /FI "PID eq ${pid}" /NH`, { windowsHide: true }, (err2, stdout2) => {
                            if (stdout2 && !stdout2.includes('没有')) {
                                const procName = stdout2.trim().split(/\s+/)[0];
                                resolve(procName);
                            } else {
                                resolve(`PID: ${pid}`);
                            }
                        });
                    } else {
                        resolve(null);
                    }
                } else {
                    resolve(null);
                }
            });
        });
    };
    
    // 为占用的端口获取进程信息
    for (const check of portChecks) {
        if (check.inUse) {
            check.process = await getPortProcess(check.port);
        }
    }
    
    // 获取所有 report-backend 进程
    const getBackendProcesses = () => {
        return new Promise((resolve) => {
            exec('tasklist /FI "IMAGENAME eq report-backend.exe"', { windowsHide: true }, (err, stdout) => {
                if (stdout && !stdout.includes('没有运行')) {
                    const lines = stdout.trim().split('\n').slice(3); // 跳过标题行
                    resolve(lines.filter(l => l.includes('report-backend')));
                } else {
                    resolve([]);
                }
            });
        });
    };
    
    const backendProcesses = await getBackendProcesses();
    
    // 生成建议
    let suggestion = '';
    const currentPortCheck = portChecks.find(c => c.port === pythonPort);
    
    if (currentPortCheck && currentPortCheck.inUse) {
        if (currentPortCheck.process && currentPortCheck.process.includes('report-backend')) {
            suggestion = '端口被旧的后端进程占用，请点击"强制清理后台进程"后重启';
        } else {
            suggestion = `端口 ${pythonPort} 被其他程序 (${currentPortCheck.process}) 占用，请换一个端口`;
        }
    } else if (backendProcesses.length > 1) {
        suggestion = '检测到多个后端进程在运行，请点击"强制清理后台进程"';
    } else {
        const availablePort = portChecks.find(c => !c.inUse);
        if (availablePort) {
            suggestion = `端口 ${availablePort.port} 可用，建议使用此端口`;
        } else {
            suggestion = '所有常用端口都被占用，请尝试其他端口号';
        }
    }
    
    return {
        configPath,
        configuredPort,
        currentPort: pythonPort,
        portChecks,
        backendProcesses,
        suggestion
    };
});

/**
 * IPC 处理：强制清理后台进程
 */
ipcMain.handle('force-cleanup', async () => {
    console.log('[Main] 强制清理所有后台进程...');
    
    return new Promise((resolve) => {
        if (process.platform === 'win32') {
            // Windows: 使用 wmic 和 taskkill 双重清理
            const { exec } = require('child_process');
            
            // 先尝试正常的 taskkill
            exec('taskkill /F /IM report-backend.exe /T', { windowsHide: true }, (err1) => {
                // 再用 wmic 确保清理干净
                exec('wmic process where "name=\'report-backend.exe\'" delete', { windowsHide: true }, (err2) => {
                    // 检查是否还有残留进程
                    exec('tasklist /FI "IMAGENAME eq report-backend.exe" /NH', { windowsHide: true }, (err3, stdout) => {
                        const hasProcess = stdout && stdout.toLowerCase().includes('report-backend');
                        if (hasProcess) {
                            // 最后尝试通过 PID 强制杀死
                            exec('for /f "tokens=2" %i in (\'tasklist /FI "IMAGENAME eq report-backend.exe" /NH\') do taskkill /F /PID %i', 
                                { windowsHide: true, shell: true }, 
                                () => {
                                    resolve({ 
                                        success: true, 
                                        message: '已尝试多种方式清理进程' 
                                    });
                                }
                            );
                        } else {
                            resolve({ 
                                success: true, 
                                message: '所有后台进程已清理完毕' 
                            });
                        }
                    });
                });
            });
        } else {
            // 非 Windows
            const { exec } = require('child_process');
            exec('pkill -9 -f report-backend', (err) => {
                resolve({ 
                    success: true, 
                    message: err ? '未找到进程或已清理' : '进程已清理' 
                });
            });
        }
    });
});

/**
 * 应用启动
 */
app.whenReady().then(async () => {
    console.log('[Main] Electron 应用启动');
    
    // 读取端口配置
    pythonPort = loadPortConfig();
    console.log(`[Main] 使用端口: ${pythonPort}`);
    
    // 启动 Python 后端
    await startPythonBackend();
    
    // 等待后端就绪
    const backendReady = await waitForBackend(15, 1000);
    if (!backendReady) {
        console.warn('[Main] Python 后端未能在预期时间内启动，应用将继续运行');
    }
    
    // 创建窗口
    createWindow();
    
    // 创建系统托盘
    createTray();

    app.on('activate', () => {
        if (mainWindow) {
            mainWindow.show();
        } else {
            createWindow();
        }
    });
});

/**
 * 所有窗口关闭时退出
 */
app.on('window-all-closed', () => {
    // 只有真正退出时才停止后端和退出应用
    if (isQuitting) {
        stopPythonBackend();
        
        if (process.platform !== 'darwin') {
            app.quit();
        }
    }
    // 如果只是隐藏窗口（最小化到托盘），不退出
});

/**
 * 应用退出前清理
 */
app.on('before-quit', () => {
    stopPythonBackend();
});

/**
 * 应用退出时清理
 */
app.on('quit', () => {
    stopPythonBackend();
});
