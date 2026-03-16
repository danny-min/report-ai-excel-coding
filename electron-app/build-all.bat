@echo off
chcp 65001 >nul
echo ========================================
echo AI 报表生成器 - 完整构建流程
echo ========================================

cd /d "%~dp0"

echo.
echo [Step 1] 安装 Node.js 依赖...
call npm install

echo.
echo [Step 2] 打包 Python 后端...
cd python-backend
pip install -r requirements.txt
pyinstaller --clean --noconfirm build.spec
cd ..

echo.
echo [Step 3] 检查 Python 后端...
if not exist "python-backend\dist\report-backend\report-backend.exe" (
    echo ❌ Python 后端打包失败！
    pause
    exit /b 1
)
echo ✅ Python 后端打包成功

echo.
echo [Step 4] 打包 Electron 应用...
call npm run dist:win

echo.
echo ========================================
echo 构建完成！
echo 输出目录: dist\
echo ========================================
pause











































