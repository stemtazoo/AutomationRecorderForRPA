@echo off

REM 仮想環境がなければ作成し、requirements.txtをインストール
if not exist .venv (
    python -m venv .venv
    call .venv\Scripts\activate
    pip install --upgrade pip
    pip install -r requirements.txt
    deactivate
    echo.
    echo 仮想環境と必要なライブラリのインストールが完了しました。
    echo もう一度、この start.bat をダブルクリックしてアプリを起動してください。
    pause
    exit /b
)

REM 仮想環境を有効化してmain.pyを起動
call .venv\Scripts\activate
python main.py
REM （必要ならここでdeactivateしてもよいが、ウィンドウ自体が閉じるなら不要）
pause
