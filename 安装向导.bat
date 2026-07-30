@echo off
chcp 65001 >nul
echo ============================================
echo   红外测温报告工具 - 安装向导
echo ============================================
echo.
echo 第一步：安装 Tesseract OCR 引擎
echo --------------------------------------------
echo 即将运行 Tesseract 安装程序。
echo 安装时请务必勾选"Additional language data"中的中文简体！
echo.
pause
start "" "%~dp0tesseract-ocr-w64-setup.exe"
echo.
echo 等待 Tesseract 安装完成后，按任意键继续...
pause >nul
echo.
echo ============================================
echo   安装完成！
echo ============================================
echo.
echo 使用方法：
echo   1. 打开 91号文 或 43号文 文件夹
echo   2. 放入红外图片到"图谱"文件夹
echo   3. 填写"必填信息.xlsx"
echo   4. 双击 exe 运行，按菜单选择操作
echo.
echo 按任意键关闭...
pause >nul
