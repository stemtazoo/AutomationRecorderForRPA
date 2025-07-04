@echo off

REM 仮想環境がなければ作成し、requirements.txtをインストール
if not exist .venv (
    python -m venv .venv
    call .venv\Scripts\activate
    pip install --upgrade pip
    pip install -r requirements.txt
    deactivate
)

REM 仮想環境を有効化してmain.pyを起動
call .venv\Scripts\activate
python main.py
REM （必要ならここでdeactivateしてもよいが、ウィンドウ自体が閉じるなら不要）
pause
