@echo off
chcp 65001 >nul
setlocal
REM CG 리스트(.xlsx) 1개 + 원본 영상(또는 영상 폴더)을 이 파일 위에 끌어다 놓으면 실행됩니다.
REM 결과는 엑셀 파일 옆 cgcut_out 폴더에 생깁니다.

if "%~1"=="" (
  echo CG 리스트 엑셀과 원본 영상을 함께 이 파일 위에 끌어다 놓으세요.
  pause
  exit /b 1
)

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo [오류] 파이썬이 없습니다. https://www.python.org/downloads/ 에서 설치하세요. 설치할 때 "Add python.exe to PATH" 체크.
  pause
  exit /b 1
)

where ffmpeg >nul 2>nul
if errorlevel 1 (
  echo [오류] ffmpeg가 없습니다. 명령 프롬프트에서  winget install Gyan.FFmpeg  실행 후 다시 시도하세요.
  pause
  exit /b 1
)

%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo openpyxl 설치 중...
  %PY% -m pip install --quiet openpyxl
)

%PY% "%~dp0cgcut.py" --drop %*
echo.
pause
