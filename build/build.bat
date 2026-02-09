@echo off
setlocal
python -m PyInstaller build\pyinstaller.spec --noconfirm
endlocal
