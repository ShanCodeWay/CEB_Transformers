@echo off
echo Running Transformer Data Application...
python main.py

if %errorlevel% neq 0 (
    echo An error occurred while running the application.
    pause
) else (
    echo Application finished successfully.
    pause
)
