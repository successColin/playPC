@echo off
chcp 65001 >nul
echo ========================================
echo   1688 商家采集工具 - 打包脚本
echo ========================================
echo.

REM 检查 PyInstaller 是否已安装
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [1/3] 正在安装 PyInstaller ...
    pip install pyinstaller
) else (
    echo [1/3] PyInstaller 已安装，跳过
)

echo.
echo [2/3] 正在清理旧的构建文件 ...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "alibaba.spec" del /f /q "alibaba.spec"

echo.
echo [3/3] 正在打包 alibaba.py → exe ...
echo        (首次打包可能需要几分钟，请耐心等待)
echo.

pyinstaller ^
    --name "1688商家采集工具" ^
    --onedir ^
    --console ^
    --noconfirm ^
    --clean ^
    --hidden-import=undetected_chromedriver ^
    --hidden-import=selenium ^
    --hidden-import=openpyxl ^
    --hidden-import=requests ^
    --hidden-import=urllib3 ^
    --hidden-import=certifi ^
    --hidden-import=charset_normalizer ^
    --collect-all undetected_chromedriver ^
    alibaba.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo   打包失败！请检查上方错误信息
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo   打包成功！
echo   输出目录: dist\1688商家采集工具\
echo   可执行文件: dist\1688商家采集工具\1688商家采集工具.exe
echo ========================================
echo.
echo 使用说明:
echo   1. 将 dist\1688商家采集工具 整个文件夹复制到目标电脑
echo   2. 目标电脑需要安装 Chrome 浏览器
echo   3. 双击 1688商家采集工具.exe 即可运行
echo   4. 也支持命令行参数，如:
echo      1688商家采集工具.exe --max-shops 100
echo.
pause
