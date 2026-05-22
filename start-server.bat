@echo off
chcp 65001 >nul
cd /d "C:\Users\57823\Desktop\AIxitong"
echo.
echo ╔══════════════════════════════════════╗
echo ║   AI 项目监理系统 - 本地服务器       ║
echo ╚══════════════════════════════════════╝
echo.
echo 正在启动本地服务器...
echo 浏览器将自动打开工作台页面
echo 按 Ctrl+C 关闭服务器
echo.
start http://localhost:8765/workbench.html
python -m http.server 8765
pause
