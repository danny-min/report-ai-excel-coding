@echo off
chcp 65001 >nul
echo ========================================
echo 打包 Python 后端
echo ========================================

cd /d "%~dp0python-backend"

echo.
echo [1/3] 安装依赖...
pip install -r requirements.txt

echo.
echo [2/3] 运行 PyInstaller...
pyinstaller --clean --noconfirm build.spec

echo.
echo [3/3] 检查输出...
if exist "dist\report-backend\report-backend.exe" (
    echo ✅ 打包成功！
    echo 输出目录: %cd%\dist\report-backend
) else (
    echo ❌ 打包失败，请检查错误信息
)

echo.
pause











































