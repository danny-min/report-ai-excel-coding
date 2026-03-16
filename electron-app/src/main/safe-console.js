/**
 * 安全控制台模块
 * 防止 EPIPE / ERR_STREAM_WRITE_AFTER_END 导致 Electron 主进程崩溃
 * 兼容 Windows + 麒麟 Linux 系统
 */

'use strict';

// 保存原始的 console 方法
const originalConsole = {
    log: console.log.bind(console),
    warn: console.warn.bind(console),
    error: console.error.bind(console),
    info: console.info.bind(console),
    debug: console.debug.bind(console),
};

// 保存原始的 stdout/stderr write 方法
const originalStdoutWrite = process.stdout.write.bind(process.stdout);
const originalStderrWrite = process.stderr.write.bind(process.stderr);

/**
 * 检查流是否可写
 */
function isStreamWritable(stream) {
    return stream && 
           !stream.destroyed && 
           stream.writable !== false &&
           typeof stream.write === 'function';
}

/**
 * 安全写入流
 */
function safeWrite(stream, originalWrite, chunk, encoding, callback) {
    // 检查流状态
    if (!isStreamWritable(stream)) {
        // 流已关闭，静默忽略
        if (typeof callback === 'function') {
            setImmediate(callback);
        }
        return true;
    }

    try {
        return originalWrite(chunk, encoding, callback);
    } catch (err) {
        // 捕获同步异常
        if (isEpipeError(err)) {
            // EPIPE 错误，静默处理
            return true;
        }
        // 其他错误向上抛出
        throw err;
    }
}

/**
 * 判断是否为 EPIPE 相关错误
 */
function isEpipeError(err) {
    if (!err) return false;
    
    const errorCodes = [
        'EPIPE',
        'ERR_STREAM_WRITE_AFTER_END',
        'ERR_STREAM_DESTROYED',
        'ECONNRESET',
        'ENOTCONN',
    ];
    
    return errorCodes.includes(err.code) || 
           (err.message && err.message.includes('broken pipe'));
}

/**
 * 创建安全的 console 方法
 */
function createSafeConsoleMethod(originalMethod, methodName) {
    return function(...args) {
        if (!isStreamWritable(process.stdout) && !isStreamWritable(process.stderr)) {
            // 两个流都不可写，静默返回
            return;
        }

        try {
            originalMethod.apply(console, args);
        } catch (err) {
            if (!isEpipeError(err)) {
                // 非 EPIPE 错误，尝试写入日志文件作为备用
                // 这里不再向控制台输出，避免循环
            }
        }
    };
}

/**
 * 安装安全控制台
 */
function installSafeConsole() {
    // 重写 process.stdout.write
    process.stdout.write = function(chunk, encoding, callback) {
        return safeWrite(process.stdout, originalStdoutWrite, chunk, encoding, callback);
    };

    // 重写 process.stderr.write
    process.stderr.write = function(chunk, encoding, callback) {
        return safeWrite(process.stderr, originalStderrWrite, chunk, encoding, callback);
    };

    // 重写 console 方法
    console.log = createSafeConsoleMethod(originalConsole.log, 'log');
    console.warn = createSafeConsoleMethod(originalConsole.warn, 'warn');
    console.error = createSafeConsoleMethod(originalConsole.error, 'error');
    console.info = createSafeConsoleMethod(originalConsole.info, 'info');
    console.debug = createSafeConsoleMethod(originalConsole.debug, 'debug');

    // 处理 stdout/stderr 的 error 事件
    process.stdout.on('error', (err) => {
        if (!isEpipeError(err)) {
            // 非 EPIPE 错误，记录到备用日志
            logToFile('stdout error: ' + err.message);
        }
    });

    process.stderr.on('error', (err) => {
        if (!isEpipeError(err)) {
            logToFile('stderr error: ' + err.message);
        }
    });
}

/**
 * 安装全局异常处理
 */
function installGlobalErrorHandlers() {
    // 未捕获的异常
    process.on('uncaughtException', (err, origin) => {
        if (isEpipeError(err)) {
            // EPIPE 错误，静默处理，不崩溃
            return;
        }
        
        // 其他未捕获的异常，记录但不崩溃
        logToFile(`Uncaught Exception [${origin}]: ${err.stack || err.message}`);
        
        // 可选：显示错误对话框
        // 但不要调用 process.exit()，让程序继续运行
    });

    // 未处理的 Promise 拒绝
    process.on('unhandledRejection', (reason, promise) => {
        if (reason && isEpipeError(reason)) {
            return;
        }
        
        logToFile(`Unhandled Rejection: ${reason}`);
    });

    // 警告事件
    process.on('warning', (warning) => {
        // 静默处理或记录
    });
}

/**
 * 备用日志写入（写入文件）
 */
const fs = require('fs');
const path = require('path');

let logFilePath = null;

function getLogFilePath() {
    if (logFilePath) return logFilePath;
    
    try {
        const { app } = require('electron');
        const userDataPath = app.getPath('userData');
        logFilePath = path.join(userDataPath, 'error.log');
    } catch {
        // 如果还没初始化 app，使用临时目录
        logFilePath = path.join(process.env.TEMP || '/tmp', 'electron-app-error.log');
    }
    
    return logFilePath;
}

function logToFile(message) {
    try {
        const logPath = getLogFilePath();
        const timestamp = new Date().toISOString();
        const logMessage = `[${timestamp}] ${message}\n`;
        
        fs.appendFileSync(logPath, logMessage, 'utf8');
    } catch {
        // 写入文件也失败，放弃
    }
}

/**
 * 创建安全的子进程通信
 */
function createSafeChildProcess(child) {
    if (!child) return child;

    // 包装 stdin.write
    if (child.stdin && typeof child.stdin.write === 'function') {
        const originalWrite = child.stdin.write.bind(child.stdin);
        
        child.stdin.write = function(chunk, encoding, callback) {
            if (!isStreamWritable(child.stdin)) {
                if (typeof callback === 'function') {
                    setImmediate(callback);
                }
                return true;
            }
            
            try {
                return originalWrite(chunk, encoding, callback);
            } catch (err) {
                if (isEpipeError(err)) {
                    return true;
                }
                throw err;
            }
        };

        child.stdin.on('error', (err) => {
            if (!isEpipeError(err)) {
                logToFile('child stdin error: ' + err.message);
            }
        });
    }

    // 监听子进程的 stdout/stderr 错误
    if (child.stdout) {
        child.stdout.on('error', (err) => {
            if (!isEpipeError(err)) {
                logToFile('child stdout error: ' + err.message);
            }
        });
    }

    if (child.stderr) {
        child.stderr.on('error', (err) => {
            if (!isEpipeError(err)) {
                logToFile('child stderr error: ' + err.message);
            }
        });
    }

    return child;
}

/**
 * 初始化所有安全措施
 */
function initialize() {
    installSafeConsole();
    installGlobalErrorHandlers();
    
    console.log('[SafeConsole] EPIPE 防护已启用');
}

// 导出模块
module.exports = {
    initialize,
    installSafeConsole,
    installGlobalErrorHandlers,
    createSafeChildProcess,
    isEpipeError,
    isStreamWritable,
    logToFile,
    
    // 导出原始方法以便需要时使用
    originalConsole,
    originalStdoutWrite,
    originalStderrWrite,
};















