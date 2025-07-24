import tkinter as tk
from tkinter import ttk
import logging
import pyautogui
import win32gui
from pynput import keyboard
from pywinauto.controls.hwndwrapper import HwndWrapper
from pywinauto import Desktop
from pywinauto.findwindows import ElementNotFoundError
import comtypes.client
from ...utils.inspector_utils import format_inspector_output, get_window_title_with_parent


class UIInspectorTab:
    """Tab for inspecting UI elements under the mouse cursor."""

    def __init__(self, app):
        """UI要素インスペクタタブを初期化します。"""

        self.app = app
        self.frame = ttk.Frame(app.notebook)
        app.notebook.add(self.frame, text='UI要素インスペクタ')

        label = tk.Label(self.frame, text="Ctrl+Shift+Xでマウス下のUI要素情報を取得します。", font=("Arial", 12))
        label.pack(pady=5)

        self.text_widget = tk.Text(self.frame, wrap=tk.WORD, font=("Arial", 12), height=15)
        self.text_widget.pack(padx=10, pady=10, fill="both", expand=True)
        self.text_widget.insert("end", "Ctrl+Shift+Xを押すと、ここにUI要素情報が表示されます。")
        self.text_widget.config(state="disabled")

        self.start_hotkey_listener()

    def start_hotkey_listener(self):
        """Ctrl+Shift+X のホットキーを登録します。"""
        try:
            self.hotkey_listener = keyboard.GlobalHotKeys({
                '<ctrl>+<shift>+x': self.inspect_element_under_cursor
            })
            self.hotkey_listener.daemon = True
            self.hotkey_listener.start()
        except Exception:
            logging.error('start_hotkey_listener error', exc_info=True)

    def generate_code_example(self, elem):
        """要素情報から簡単なクリックコードを生成します。"""
        props = []
        title = elem.window_text()
        ctrl_type = elem.element_info.control_type
        auto_id = elem.element_info.automation_id
        if title:
            props.append(f'title="{title}"')
        if ctrl_type:
            props.append(f'control_type="{ctrl_type}"')
        if auto_id:
            props.append(f'automation_id="{auto_id}"')
        if props:
            return f'dlg.child_window({", ".join(props)}).click_input()'
        return "# 要素を特定する情報が不足しています"

    def find_deepest_element_at_point(self, x, y, backend='uia'):
        """指定された座標で最も深い（具体的な）UI要素を見つけます。"""
        try:
            # まず基本的な方法で要素を取得
            desktop = Desktop(backend=backend)
            root_elem = desktop.from_point(x, y)
            
            if not root_elem:
                return None
            
            # より深い要素を探索
            current_elem = root_elem
            max_depth = 10  # 無限ループを防ぐための最大深度
            depth = 0
            
            while depth < max_depth:
                try:
                    # 子要素を取得
                    children = current_elem.children()
                    if not children:
                        break
                    
                    # 指定座標を含む子要素を探す
                    target_child = None
                    for child in children:
                        try:
                            rect = child.rectangle()
                            if (rect.left <= x <= rect.right and 
                                rect.top <= y <= rect.bottom):
                                target_child = child
                                break
                        except:
                            continue
                    
                    if target_child is None:
                        break
                    
                    # より具体的な要素が見つかった場合、それを使用
                    if (hasattr(target_child, 'element_info') and 
                        target_child.element_info.control_type and
                        target_child.element_info.control_type not in ['Window', 'Pane']):
                        current_elem = target_child
                        depth += 1
                    else:
                        break
                        
                except:
                    break
            
            return current_elem
            
        except Exception as e:
            logging.error(f"find_deepest_element_at_point error: {e}")
            return None

    def get_element_with_uiautomation(self, x, y):
        """UIAutomationを直接使用してより詳細な要素を取得します。"""
        try:
            # UIAutomationオブジェクトを作成
            uia = comtypes.client.CreateObject("UIAutomation.CUIAutomation")
            
            # 指定座標から要素を取得
            point = comtypes.pointer(comtypes.Structure._fields_[0][1](x, y))
            element = uia.ElementFromPoint(point)
            
            if element:
                # 要素の詳細情報を取得
                element_info = {
                    'name': element.CurrentName if hasattr(element, 'CurrentName') else 'N/A',
                    'control_type': element.CurrentControlType if hasattr(element, 'CurrentControlType') else 'N/A',
                    'automation_id': element.CurrentAutomationId if hasattr(element, 'CurrentAutomationId') else 'N/A',
                    'class_name': element.CurrentClassName if hasattr(element, 'CurrentClassName') else 'N/A',
                    'help_text': element.CurrentHelpText if hasattr(element, 'CurrentHelpText') else 'N/A',
                    'bounding_rect': element.CurrentBoundingRectangle if hasattr(element, 'CurrentBoundingRectangle') else 'N/A'
                }
                return element, element_info
            return None, None
            
        except Exception as e:
            logging.error(f"get_element_with_uiautomation error: {e}")
            return None, None

    def get_tkinter_specific_elements(self, x, y):
        """Tkinter専用の詳細な要素探索を行います。"""
        try:
            # Tkinterウィンドウのすべての子要素を詳細に探索
            root_hwnd = win32gui.WindowFromPoint((x, y))
            
            # 親ウィンドウを取得
            parent_hwnd = root_hwnd
            while True:
                temp_parent = win32gui.GetParent(parent_hwnd)
                if temp_parent:
                    parent_hwnd = temp_parent
                else:
                    break
            
            # すべての子ウィンドウを収集
            child_windows = []
            
            def enum_callback(hwnd, results):
                try:
                    rect = win32gui.GetWindowRect(hwnd)
                    class_name = win32gui.GetClassName(hwnd)
                    window_text = win32gui.GetWindowText(hwnd)
                    
                    # 座標が範囲内にある要素のみ収集
                    if (rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3] and
                        rect[2] - rect[0] > 0 and rect[3] - rect[1] > 0):
                        results.append({
                            'hwnd': hwnd,
                            'class_name': class_name,
                            'window_text': window_text,
                            'rect': rect,
                            'area': (rect[2] - rect[0]) * (rect[3] - rect[1])
                        })
                except:
                    pass
                return True
            
            # 親ウィンドウから再帰的に子要素を探索
            win32gui.EnumChildWindows(parent_hwnd, enum_callback, child_windows)
            
            # 面積が最小の要素（最も具体的な要素）を見つける
            if child_windows:
                # 面積でソート（小さい順）
                child_windows.sort(key=lambda x: x['area'])
                
                # 最も小さい要素を返す（ただし、最小サイズ制限を設ける）
                for element in child_windows:
                    if element['area'] > 100:  # 10x10ピクセル以上
                        return element
                
                # それでも見つからない場合は最初の要素を返す
                return child_windows[0] if child_windows else None
            
            return None
            
        except Exception as e:
            logging.error(f"get_tkinter_specific_elements error: {e}")
            return None

    def get_detailed_element_at_coordinate(self, x, y, backend='uia'):
        """座標における詳細な要素情報を段階的に取得します。"""
        try:
            desktop = Desktop(backend=backend)
            
            # レベル1: 基本的な要素取得
            try:
                root_element = desktop.from_point(x, y)
                if not root_element:
                    return None
            except:
                return None
            
            # レベル2: より詳細な子要素の探索
            candidates = [root_element]
            
            # 3レベルまで子要素を探索
            for level in range(3):
                new_candidates = []
                for candidate in candidates:
                    try:
                        children = candidate.children()
                        for child in children:
                            try:
                                rect = child.rectangle()
                                # 座標が子要素の範囲内にある場合
                                if (rect.left <= x <= rect.right and 
                                    rect.top <= y <= rect.bottom):
                                    new_candidates.append(child)
                            except:
                                continue
                    except:
                        continue
                
                if new_candidates:
                    candidates = new_candidates
                else:
                    break
            
            # 最も具体的な要素を選択（面積が最小のもの）
            if candidates:
                best_candidate = None
                min_area = float('inf')
                
                for candidate in candidates:
                    try:
                        rect = candidate.rectangle()
                        area = (rect.right - rect.left) * (rect.bottom - rect.top)
                        
                        # 要素に有用な情報があるかチェック
                        has_useful_info = (
                            candidate.window_text() or
                            (hasattr(candidate, 'element_info') and 
                             candidate.element_info.automation_id) or
                            (hasattr(candidate, 'element_info') and 
                             candidate.element_info.control_type not in ['Window', 'Pane', ''])
                        )
                        
                        if area < min_area and (has_useful_info or area < 10000):
                            min_area = area
                            best_candidate = candidate
                    except:
                        continue
                
                return best_candidate if best_candidate else candidates[0]
            
            return root_element
            
        except Exception as e:
            logging.error(f"get_detailed_element_at_coordinate error: {e}")
            return None

    def get_chrome_specific_element(self, x, y):
        """Chrome専用の要素取得を試行します。"""
        try:
            # アクセシビリティAPIを使用してChrome要素を取得
            import win32api
            import win32con

            # より精密な座標での要素検索
            hwnd = win32gui.WindowFromPoint((x, y))

            # Chromeの場合、複数のレベルで子ウィンドウを探す
            chrome_hwnds = []

            def enum_child_windows(hwnd, results):
                def callback(child_hwnd, _):
                    class_name = win32gui.GetClassName(child_hwnd)
                    if class_name:
                        results.append((child_hwnd, class_name))
                    return True
                win32gui.EnumChildWindows(hwnd, callback, None)

            # 親ウィンドウから子ウィンドウを列挙
            parent_hwnd = win32gui.GetParent(hwnd)
            if parent_hwnd:
                enum_child_windows(parent_hwnd, chrome_hwnds)

            # 座標に最も近い要素を探す
            closest_element = None
            min_distance = float('inf')

            for child_hwnd, class_name in chrome_hwnds:
                try:
                    rect = win32gui.GetWindowRect(child_hwnd)
                    if rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]:
                        # 要素の中心からの距離を計算
                        center_x = (rect[0] + rect[2]) // 2
                        center_y = (rect[1] + rect[3]) // 2
                        distance = ((x - center_x) ** 2 + (y - center_y) ** 2) ** 0.5

                        if distance < min_distance:
                            min_distance = distance
                            closest_element = (child_hwnd, class_name, rect)
                except Exception:
                    continue

            return closest_element

        except Exception as e:
            logging.error(f"get_chrome_specific_element error: {e}")
            return None

    def get_accessibility_info(self, x, y):
        """アクセシビリティ情報を取得します。"""
        try:
            # IAccessibleインターフェースを使用
            import comtypes.client
            import comtypes.gen.Accessibility as Accessibility
            
            hwnd = win32gui.WindowFromPoint((x, y))
            
            # アクセシビリティオブジェクトを取得
            try:
                import oleacc
                pacc, child_id = oleacc.AccessibleObjectFromWindow(
                    hwnd, oleacc.OBJID_CLIENT, oleacc.IAccessible
                )
                
                if pacc:
                    # 座標からアクセシブル要素を取得
                    acc_element = pacc.accHitTest(x, y)
                    if acc_element:
                        acc_info = {
                            'name': pacc.get_accName(acc_element) if hasattr(pacc, 'get_accName') else 'N/A',
                            'description': pacc.get_accDescription(acc_element) if hasattr(pacc, 'get_accDescription') else 'N/A',
                            'role': pacc.get_accRole(acc_element) if hasattr(pacc, 'get_accRole') else 'N/A',
                            'state': pacc.get_accState(acc_element) if hasattr(pacc, 'get_accState') else 'N/A',
                            'value': pacc.get_accValue(acc_element) if hasattr(pacc, 'get_accValue') else 'N/A'
                        }
                        return acc_info
            except ImportError:
                # oleaccが利用できない場合はスキップ
                pass
                
        except Exception as e:
            logging.error(f"get_accessibility_info error: {e}")
        
        return None

    def get_element_under_mouse(self):
        """現在のマウス位置にある要素を取得します。"""
        try:
            x, y = pyautogui.position()
            backend = self.app.backend_var.get()
            
            # ウィンドウクラスを確認
            hwnd = win32gui.WindowFromPoint((x, y))
            window_class = win32gui.GetClassName(hwnd)
            
            # Tkinter専用処理
            if 'Tk' in window_class:
                tk_element = self.get_tkinter_specific_elements(x, y)
                if tk_element:
                    return {'type': 'tkinter_specific', 'element': tk_element, 'info': None}
            
            # Chrome等のブラウザの場合は特別な処理
            if 'Chrome' in window_class or 'Browser' in window_class:
                # Chrome専用の要素取得を試行
                chrome_element = self.get_chrome_specific_element(x, y)
                if chrome_element:
                    return {'type': 'chrome_specific', 'element': chrome_element, 'info': None}
                
                # アクセシビリティ情報を取得
                acc_info = self.get_accessibility_info(x, y)
                if acc_info:
                    return {'type': 'accessibility', 'element': None, 'info': acc_info}
            
            # 詳細な座標ベース探索を試行
            detailed_elem = self.get_detailed_element_at_coordinate(x, y, backend)
            if detailed_elem:
                return {'type': 'detailed_coordinate', 'element': detailed_elem, 'info': None}
            
            # UIAutomationを直接使用してみる
            uia_element, uia_info = self.get_element_with_uiautomation(x, y)
            if uia_element and uia_info:
                return {'type': 'uiautomation', 'element': uia_element, 'info': uia_info}
            
            # 改良されたメソッドを試す
            elem = self.find_deepest_element_at_point(x, y, backend)
            if elem:
                return {'type': 'pywinauto', 'element': elem, 'info': None}
            
            # それでも見つからない場合は従来の方法を使用
            elem = Desktop(backend=backend).from_point(x, y)
            if elem:
                return {'type': 'pywinauto', 'element': elem, 'info': None}
            
            return None
            
        except ElementNotFoundError:
            return None

    def get_alternative_element_info(self, x, y):
        """Win32 APIを使った代替の要素取得方法"""
        try:
            # より詳細なWin32情報を取得
            point = (x, y)
            hwnd = win32gui.WindowFromPoint(point)
            
            # 子ウィンドウを探す
            child_hwnd = win32gui.ChildWindowFromPoint(hwnd, point)
            if child_hwnd and child_hwnd != hwnd:
                hwnd = child_hwnd
            
            # さらに深い子ウィンドウを探す
            while True:
                deeper_child = win32gui.ChildWindowFromPoint(hwnd, 
                    (x - win32gui.GetWindowRect(hwnd)[0], 
                     y - win32gui.GetWindowRect(hwnd)[1]))
                if deeper_child and deeper_child != hwnd:
                    hwnd = deeper_child
                else:
                    break
            
            return hwnd
            
        except Exception as e:
            logging.error(f"get_alternative_element_info error: {e}")
            return win32gui.WindowFromPoint((x, y))

    def inspect_element_under_cursor(self):
        """マウス下の要素情報を取得して表示します。"""
        try:
            x, y = pyautogui.position()
            elem_data = self.get_element_under_mouse()
            
            if not elem_data:
                result = "要素が見つかりませんでした。"
            else:
                # より詳細なHWND取得
                hwnd = self.get_alternative_element_info(x, y)
                window_title = get_window_title_with_parent(hwnd)
                backend = self.app.backend_var.get()
                
                dlg_code = f"""【dlg設定サンプル】
from pywinauto.application import Application
# backend は 'uia' または 'win32' から選べます
app = Application(backend=\"{backend}\").connect(title=\"{window_title}\")
dlg = app.window(title=\"{window_title}\")
# ↓このdlg変数を使って下のコード例をそのまま利用できます！
"""
                
                # 座標情報
                coord_info = f"\n【マウス座標】\nX: {x}, Y: {y}\n"
                
                # 要素の種類に応じて情報を取得
                if elem_data['type'] == 'tkinter_specific':
                    # Tkinter専用取得の場合
                    tk_elem = elem_data['element']
                    tk_result = f"""
【Tkinter専用取得結果】
ウィンドウハンドル: {tk_elem['hwnd']}
クラス名: {tk_elem['class_name']}
ウィンドウテキスト: {tk_elem['window_text']}
座標: {tk_elem['rect']}
面積: {tk_elem['area']}

【推奨操作コード】
# ハンドルを使用した操作
dlg.child_window(handle={tk_elem['hwnd']}).click()

# クラス名とテキストを組み合わせた操作
dlg.child_window(class_name="{tk_elem['class_name']}", title="{tk_elem['window_text']}").click()

# 座標ベースの直接操作（最も確実）
import pyautogui
pyautogui.click({x}, {y})
"""
                    result = f"{dlg_code}\n画面名: {window_title}\n{coord_info}\n{tk_result}"
                
                elif elem_data['type'] == 'detailed_coordinate':
                    # 詳細座標探索の場合
                    elem = elem_data['element']
                    detailed_info = {
                        "name": elem.window_text() if elem.window_text() else 'N/A',
                        "class_name": elem.element_info.class_name if hasattr(elem, 'element_info') and elem.element_info.class_name else 'N/A',
                        "control_type": elem.element_info.control_type if hasattr(elem, 'element_info') and elem.element_info.control_type else 'N/A',
                        "automation_id": elem.element_info.automation_id if hasattr(elem, 'element_info') and elem.element_info.automation_id else 'N/A',
                        "rectangle": str(elem.rectangle()),
                        "code_example": self.generate_code_example(elem),
                    }
                    
                    detailed_result = f"""
【詳細座標探索結果】
名前: {detailed_info['name']}
クラス名: {detailed_info['class_name']}
コントロールタイプ: {detailed_info['control_type']}
オートメーションID: {detailed_info['automation_id']}
矩形: {detailed_info['rectangle']}

【推奨コード】
{detailed_info['code_example']}

【代替コード】
# より具体的な特定方法
dlg.child_window(class_name="{detailed_info['class_name']}", title="{detailed_info['name']}").click_input()
"""
                    
                    # Win32情報も併せて表示
                    win32_wrap = HwndWrapper(hwnd)
                    win32_info = {
                        "window_text": win32_wrap.window_text(),
                        "class_name": win32_wrap.friendly_class_name(),
                        "handle": win32_wrap.handle,
                        "rectangle": str(win32_wrap.rectangle()),
                    }
                    
                    result = f"{dlg_code}\n画面名: {window_title}\n{coord_info}\n{detailed_result}\n{format_inspector_output(detailed_info, win32_info)}"
                
                elif elem_data['type'] == 'chrome_specific':
                    # Chrome専用取得の場合
                    hwnd, class_name, rect = elem_data['element']
                    chrome_result = f"""
【Chrome専用取得結果】
ウィンドウハンドル: {hwnd}
クラス名: {class_name}
座標: {rect}

【注意】
Chromeの内部要素は通常のUI自動化では取得困難です。
以下の代替手段を検討してください：

1. Chrome拡張機能の使用
2. Seleniumによるブラウザ自動化
3. 座標ベースのクリック操作
4. Chrome DevTools Protocolの使用

【座標ベースの操作例】
import pyautogui
pyautogui.click({x}, {y})  # 直接座標をクリック
"""
                    result = f"{dlg_code}\n画面名: {window_title}\n{coord_info}\n{chrome_result}"
                
                elif elem_data['type'] == 'accessibility':
                    # アクセシビリティ取得の場合
                    acc_info = elem_data['info']
                    acc_result = f"""
【アクセシビリティ情報】
名前: {acc_info['name']}
説明: {acc_info['description']}
ロール: {acc_info['role']}
状態: {acc_info['state']}
値: {acc_info['value']}

【推奨操作方法】
# 座標ベースでの操作を推奨
import pyautogui
pyautogui.click({x}, {y})
"""
                    result = f"{dlg_code}\n画面名: {window_title}\n{coord_info}\n{acc_result}"
                
                elif elem_data['type'] == 'uiautomation':
                    # UIAutomation直接取得の場合
                    uia_info = elem_data['info']
                    uia_result = f"""
【UIAutomation直接取得結果】
名前: {uia_info['name']}
コントロールタイプ: {uia_info['control_type']}
オートメーションID: {uia_info['automation_id']}
クラス名: {uia_info['class_name']}
ヘルプテキスト: {uia_info['help_text']}
境界矩形: {uia_info['bounding_rect']}

【推奨コード例】
# UIAutomationIDが利用可能な場合
dlg.child_window(auto_id="{uia_info['automation_id']}").click_input()
# または名前で特定
dlg.child_window(title="{uia_info['name']}").click_input()
"""
                    
                    # Win32情報も併せて取得
                    win32_wrap = HwndWrapper(hwnd)
                    win32_info = {
                        "window_text": win32_wrap.window_text(),
                        "class_name": win32_wrap.friendly_class_name(),
                        "handle": win32_wrap.handle,
                        "rectangle": str(win32_wrap.rectangle()),
                        "code_example": f'dlg.child_window(title=\"{win32_wrap.window_text()}\", class_name=\"{win32_wrap.friendly_class_name()}\", handle={win32_wrap.handle}).click()',
                    }
                    
                    result = f"{dlg_code}\n画面名: {window_title}\n{coord_info}\n{uia_result}\n{format_inspector_output({}, win32_info)}"
                    
                else:
                    # pywinauto取得の場合（従来の処理）
                    elem = elem_data['element']
                    uia_info = {
                        "name": elem.window_text(),
                        "class_name": elem.element_info.class_name if hasattr(elem, 'element_info') else 'N/A',
                        "control_type": elem.element_info.control_type if hasattr(elem, 'element_info') else 'N/A',
                        "automation_id": elem.element_info.automation_id if hasattr(elem, 'element_info') else 'N/A',
                        "rectangle": str(elem.rectangle()),
                        "code_example": self.generate_code_example(elem),
                    }
                    
                    win32_wrap = HwndWrapper(hwnd)
                    win32_info = {
                        "window_text": win32_wrap.window_text(),
                        "class_name": win32_wrap.friendly_class_name(),
                        "handle": win32_wrap.handle,
                        "rectangle": str(win32_wrap.rectangle()),
                        "code_example": f'dlg.child_window(title=\"{win32_wrap.window_text()}\", class_name=\"{win32_wrap.friendly_class_name()}\", handle={win32_wrap.handle}).click()',
                    }
                    
                    result = f"{dlg_code}\n画面名: {window_title}\n{coord_info}\n{format_inspector_output(uia_info, win32_info)}"

            self.text_widget.config(state="normal")
            self.text_widget.delete("1.0", "end")
            self.text_widget.insert("end", result)
            self.text_widget.config(state="disabled")
            
        except Exception as e:
            logging.error(f"inspect_element_under_cursor error: {e}", exc_info=True)
            self.text_widget.config(state="normal")
            self.text_widget.delete("1.0", "end")
            self.text_widget.insert("end", f"エラーが発生しました: {str(e)}")
            self.text_widget.config(state="disabled")