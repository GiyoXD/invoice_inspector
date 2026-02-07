@echo off
echo Building Python CLI Adapter...
if not exist dist mkdir dist
pyinstaller --noconfirm --onefile --name inspector_cli cli.py
echo Build Complete. Checking dist...
if exist dist\inspector_cli.exe (
    echo [SUCCESS] inspector_cli.exe created.
) else (
    echo [ERROR] Build failed.
)
pause
