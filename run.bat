@echo off
chcp 65001 > nul

echo start [%date% %time%]
echo.

if exist convert.exe (
    convert.exe
) else (
    echo [오류] convert.exe 파일을 찾을 수 없습니다.
    pause
)

echo.
echo 작업이 완료되었습니다. 2초 뒤 창이 닫힙니다.
timeout /t 2 > nul
exit