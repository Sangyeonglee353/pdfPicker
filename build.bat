@echo off
echo PDF PICKER 빌드 시작...

REM 기존 가상환경 삭제
rmdir /s /q venv

REM Python 3.13 경로 확인
set PYTHON_PATH=C:\Users\Ian\AppData\Local\Programs\Python\Python313\python.exe
if not exist "%PYTHON_PATH%" (
    echo Python 3.13이 설치되어 있지 않습니다.
    pause
    exit /b 1
)

REM 가상환경 생성 및 활성화
"%PYTHON_PATH%" -m venv venv
call venv\Scripts\activate

REM 필요한 패키지 설치
pip install --upgrade wheel setuptools
pip install --upgrade -r requirements.txt

REM PyInstaller로 실행 파일 생성
python -m PyInstaller build.spec --clean

echo 빌드 완료!
pause