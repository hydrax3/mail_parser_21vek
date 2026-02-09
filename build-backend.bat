@echo off
echo Building Python Backend...
python -m PyInstaller --onefile --noconsole --name api python/api.py

if not exist "electron-app\resources" mkdir "electron-app\resources"
copy "dist\api.exe" "electron-app\resources\api.exe"
echo Done.
