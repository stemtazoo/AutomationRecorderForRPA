"""Utility functions for inspector output."""

import win32gui


def format_inspector_output(uia_info, win32_info):
    """Return a nicely formatted string for UIA and Win32 element info."""
    uia = f'''[UIA]
=== 要素情報 ===
タイトル: {uia_info.get("name", "")}
クラス名: {uia_info.get("class_name", "")}
コントロールタイプ: {uia_info.get("control_type", "")}
オートメーションID: {uia_info.get("automation_id", "")}
矩形: {uia_info.get("rectangle", "")}

階層パス例:
dlg.child_window(title="{uia_info.get("name", "")}", control_type="{uia_info.get("control_type", "")}", automation_id="{uia_info.get("automation_id", "")}")

パターン例:
.click_input()   # クリック
.set_focus()     # フォーカス移動
.get_value()     # 値取得
.window_text()   # テキスト取得

コード例:
{uia_info.get("code_example", "")}
'''

    win32 = f'''[Win32]
=== 要素情報 ===
ウィンドウテキスト: {win32_info.get("window_text", "")}
クラス名: {win32_info.get("class_name", "")}
ハンドル: {win32_info.get("handle", "")}
矩形: {win32_info.get("rectangle", "")}

階層パス例:
dlg.child_window(title="{win32_info.get("window_text", "")}", class_name="{win32_info.get("class_name", "")}", handle={win32_info.get("handle", "")})

パターン例:
.click()         # クリック
.set_focus()     # フォーカス移動
.window_text()   # テキスト取得
.is_visible()    # 表示状態の判定

コード例:
{win32_info.get("code_example", "")}
'''

    hint = 'ヒント:\n・階層パス例は要素の一意な指定に便利です。\n・パターン例は自動化でよく使うメソッドです。\n・コード例はそのままスクリプトにコピペして使えます。'
    return uia + "\n\n" + win32 + "\n\n" + hint


def get_window_title_with_parent(hwnd):
    """Return window title or walk parent windows if empty."""
    title = win32gui.GetWindowText(hwnd)
    if title:
        return title
    parent = win32gui.GetParent(hwnd)
    if parent:
        return get_window_title_with_parent(parent)
    return ""
