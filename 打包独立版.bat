@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" call "安装工具.bat"
".venv\Scripts\python.exe" -m pip install pyinstaller==6.10.0
".venv\Scripts\pyinstaller.exe" --noconfirm --clean "每周提现工具.spec"
if errorlevel 1 goto :failed
if not exist "发布包" mkdir "发布包"
if not exist "发布包\每周钱包流水" mkdir "发布包\每周钱包流水"
copy /y "dist\每周提现工具.exe" "发布包\每周提现工具.exe" >nul
copy /y "飞书配置.json" "发布包\飞书配置.json" >nul
copy /y "Windows使用说明.md" "发布包\Windows使用说明.md" >nul
echo 请把 BigSeller 钱包流水放在这里。>"发布包\每周钱包流水\请把流水放在这里.txt"
echo 独立程序已生成在“发布包”文件夹，不需要安装 Python。
pause
exit /b 0
:failed
echo 打包失败，请查看上方错误信息。
pause
exit /b 1
