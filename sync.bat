@echo off
echo ===================================================
echo  Pushing updates to GitHub (excel-ai-agent)...
echo ===================================================
git add -A
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value 2^>nul') do set datetime=%%I
if "%datetime%"=="" (
    set commit_msg=Auto-update: %date% %time%
) else (
    set commit_msg=Auto-update: %datetime:~0,4%-%datetime:~4,2%-%datetime:~6,2% %datetime:~8,2%:%datetime:~10,2%
)
git commit -m "%commit_msg%"
git push origin main
echo.
echo [DONE] GitHub repository updated successfully!
echo.
pause