import asyncio
import threading
import time
import logging
import toml
import pystray
from PIL import Image, ImageDraw
import pygetwindow as gw
import psutil
import os
import ctypes
from ctypes import wintypes
import re
from screeninfo import get_monitors
from pyvda import AppView, VirtualDesktop, get_virtual_desktops
import win32gui
import win32con
import queue

# --- 定数 ---
SETTINGS_FILE = "settings.toml"
LOG_FILE = "log.txt"
ANCHOR_POINTS = {
    "TopLeft": (0.0, 0.0), "TopCenter": (0.5, 0.0), "TopRight": (1.0, 0.0),
    "MiddleLeft": (0.0, 0.5), "MiddleCenter": (0.5, 0.5), "MiddleRight": (1.0, 0.5),
    "BottomLeft": (0.0, 1.0), "BottomCenter": (0.5, 1.0), "BottomRight": (1.0, 1.0)
}
# WinEventHookで無視するプロセスのリスト
IGNORED_PROCESSES = {"SystemSettings.exe", "TextInputHost.exe", "SearchHost.exe", "SearchApp.exe", "ShellExperienceHost.exe", "StartMenuExperienceHost.exe", "Widgets.exe", "LockApp.exe"}

# --- ログレベル変換 ---
def get_log_level_from_string(level_str: str) -> int:
    """ログレベルの文字列をloggingの定数に変換する"""
    return getattr(logging, level_str.upper(), logging.INFO)

# --- ロギング設定 ---
def setup_logging(level: int = logging.INFO):
    """ロギングの基本設定を行う"""
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(threadName)s: %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding='utf-8', mode='w'),
            logging.StreamHandler()
        ]
    )
    logging.getLogger("pyvda").setLevel(logging.INFO)
    logging.getLogger("PIL").setLevel(logging.INFO)

# --- 設定管理 ---
class Settings:
    def __init__(self, filepath):
        self.filepath = filepath
        self.globals = {}
        self.rules = []
        self.load()

    def load(self):
        """設定ファイルを読み込む"""
        try:
            with open(self.filepath, "r", encoding="utf-8") as f:
                data = toml.load(f)
            self.globals = data.get("global", {})
            self.rules = data.get("rules", [])
            logging.info(f"設定ファイルを読み込みました。Global: {len(self.globals)}項目, Rules: {len(self.rules)}個")
        except FileNotFoundError:
            logging.error(f"設定ファイル '{self.filepath}' が見つかりません。")
        except toml.TomlDecodeError as e:
            logging.error(f"設定ファイル '{self.filepath}' の形式が正しくありません (行: {e.lineno}, 列: {e.colno}): {e}")
        except Exception as e:
            logging.error(f"設定ファイルの読み込み中に予期せぬエラーが発生しました: {e}", exc_info=True)

# --- 座標計算 ---
class Calculator:
    def __init__(self, monitors, global_settings):
        self.monitors = monitors
        self.globals = global_settings
        logging.debug(f"Calculatorを初期化しました。モニター数: {len(monitors)}")

    def _parse_value(self, value, base_pixels):
        """サイズや座標の値をピクセル単位に変換する"""
        if isinstance(value, (int, float)):
            return int(value)
        if isinstance(value, str):
            try:
                value = value.strip()
                if value.endswith('%'):
                    return int(base_pixels * float(value.strip('%')) / 100)
                if value.endswith('px'):
                    return int(value.strip('px'))
                return int(value)
            except (ValueError, TypeError) as e:
                logging.warning(f'値 "{value}" の解析に失敗しました: {e}。Noneを返します。')
                return None
        return None

    def get_target_rect(self, rule_action, window):
        """ルールとウィンドウ情報に基づき、最終的な座標とサイズ (x, y, w, h) を計算する"""
        try:
            monitor = None
            target_monitor_num = rule_action.get("target_monitor")
            if target_monitor_num is not None:
                if isinstance(target_monitor_num, int) and 1 <= target_monitor_num <= len(self.monitors):
                    monitor = self.monitors[target_monitor_num - 1]
                    logging.debug(f"ルール指定により、ターゲットモニター {target_monitor_num} を使用します。")
                else:
                    logging.warning(f"指定されたターゲットモニター番号 '{target_monitor_num}' は無効です（有効範囲: 1～{len(self.monitors)}）。フォールバックしてモニターを自動検出します。")
            
            if monitor is None:
                monitor = self.get_window_monitor(window)

            is_absolute_move = isinstance(rule_action.get("move_to"), dict)
            monitor_idx = self.monitors.index(monitor)
            monitor_id = f"monitor_{monitor_idx + 1}"
            logging.debug(f"対象モニター: {monitor_id} ({monitor.width}x{monitor.height} at ({monitor.x},{monitor.y}))")

            offsets = self.globals.get("monitor_offsets", {})
            g_offset = {}
            if not is_absolute_move:
                g_offset = offsets.get(monitor_id, offsets.get("default", {}))
                logging.debug(f"グローバルオフセットを適用します: {g_offset}")

            offset_top = g_offset.get("top", 0)
            offset_bottom = g_offset.get("bottom", 0)
            offset_left = g_offset.get("left", 0)
            offset_right = g_offset.get("right", 0)

            work_area_x = monitor.x + offset_left
            work_area_y = monitor.y + offset_top
            work_area_width = monitor.width - offset_left - offset_right
            work_area_height = monitor.height - offset_top - offset_bottom
            logging.debug(f"作業領域: {work_area_width}x{work_area_height} at ({work_area_x},{work_area_y})")

            resize_to = rule_action.get("resize_to", {})
            w_val = resize_to.get("w", resize_to.get("width"))
            h_val = resize_to.get("h", resize_to.get("height"))
            width = self._parse_value(w_val, work_area_width) if w_val is not None else window.width
            height = self._parse_value(h_val, work_area_height) if h_val is not None else window.height
            if width is None or height is None:
                logging.warning("サイズ指定の解析に失敗したため、現在のサイズを維持します。")
                width = window.width if width is None else width
                height = window.height if height is None else height

            move_to = rule_action.get("move_to")
            anchor_name = rule_action.get("anchor", "TopLeft")
            anchor_x_ratio, anchor_y_ratio = ANCHOR_POINTS.get(anchor_name, (0.0, 0.0))
            logging.debug(f"ウィンドウのアンカー: {anchor_name} ({anchor_x_ratio}, {anchor_y_ratio})")

            base_x, base_y = None, None
            if isinstance(move_to, str):
                target_anchor_name = move_to
                target_x_ratio, target_y_ratio = ANCHOR_POINTS.get(target_anchor_name, (0.0, 0.0))
                base_x = work_area_x + int(work_area_width * target_x_ratio)
                base_y = work_area_y + int(work_area_height * target_y_ratio)
                logging.debug(f"移動先アンカー '{target_anchor_name}' -> ベース座標 ({base_x}, {base_y})")
            elif isinstance(move_to, dict):
                abs_x_val = move_to.get("x")
                abs_y_val = move_to.get("y")
                abs_x = self._parse_value(abs_x_val, monitor.width)
                abs_y = self._parse_value(abs_y_val, monitor.height)
                if abs_x is not None: base_x = monitor.x + abs_x
                if abs_y is not None: base_y = monitor.y + abs_y
                logging.debug(f"絶対/相対座標 {move_to} -> ベース座標 ({base_x}, {base_y})")

            if base_x is None and base_y is None:
                logging.debug("移動指定がないため、現在の位置を基準とします。")
                return window.left, window.top, width, height
            
            final_x = base_x if base_x is not None else window.left
            final_y = base_y if base_y is not None else window.top

            final_x -= int(width * anchor_x_ratio)
            final_y -= int(height * anchor_y_ratio)
            logging.debug(f"ウィンドウアンカー適用後 -> ({final_x}, {final_y})")

            rule_offset = rule_action.get("offset", {})
            offset_x = rule_offset.get("x", 0)
            offset_y = rule_offset.get("y", 0)
            if offset_x != 0 or offset_y != 0:
                final_x += offset_x
                final_y += offset_y
                logging.debug(f"ルールオフセット適用後 -> ({final_x}, {final_y})")

            return final_x, final_y, width, height
        except Exception as e:
            logging.error(f"座標計算中に予期せぬエラーが発生しました: {e}", exc_info=True)
            return window.left, window.top, window.width, window.height

    def get_window_monitor(self, window):
        """ウィンドウの中心が含まれるモニターを返す"""
        try:
            win_center_x = window.left + window.width / 2
            win_center_y = window.top + window.height / 2
            for i, m in enumerate(self.monitors):
                if m.x <= win_center_x < m.x + m.width and m.y <= win_center_y < m.y + m.height:
                    logging.debug(f"ウィンドウ '{window.title}' はモニター {i+1} にあります。")
                    return m
        except Exception as e:
            logging.warning(f"ウィンドウ '{window.title}' のモニター特定中にエラー: {e}。プライマリモニターを返します。")

        logging.debug(f"ウィンドウ '{window.title}' がどのモニターにも見つからないため、プライマリモニターを返します。")
        return self.monitors[0]

# --- ウィンドウ処理 ---
class WindowManager:
    def __init__(self, settings, loop):
        self.settings = settings
        self.loop = loop
        try:
            monitors = get_monitors()
        except Exception as e:
            logging.critical(f"モニター情報の取得に失敗しました。アプリケーションを続行できません: {e}")
            raise
        self.calculator = Calculator(monitors, self.settings.globals)
        self.processed_windows = set()
        self.is_paused = False
        self.lock = threading.Lock()

    def _get_process_name(self, hwnd):
        """ウィンドウハンドルからプロセス名を取得する"""
        try:
            pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value == 0:
                return None
            return psutil.Process(pid.value).name()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            return None
        except Exception as e:
            logging.error(f"プロセス名取得中に予期せぬエラー (PID: {pid.value}): {e}", exc_info=True)
            return None

    def _check_single_condition(self, window, process_name, class_name, condition):
        """単一の条件をチェックする"""
        title_pattern = condition.get("title")
        process_pattern = condition.get("process")
        class_pattern = condition.get("class")

        if title_pattern:
            try:
                if title_pattern.startswith("regex:"):
                    pattern = title_pattern.replace("regex:", "", 1)
                    if re.search(pattern, window.title, re.IGNORECASE if not condition.get("case_sensitive") else 0):
                        return True
                elif title_pattern.lower() in window.title.lower():
                    return True
            except re.error as e:
                logging.warning(f'タイトル条件の正規表現パターン "{pattern}" が不正です: {e}')
                return False

        if process_pattern and process_name:
            try:
                if process_pattern.startswith("regex:"):
                    pattern = process_pattern.replace("regex:", "", 1)
                    if re.fullmatch(pattern, process_name, re.IGNORECASE if not condition.get("case_sensitive") else 0):
                        return True
                elif process_pattern.lower() == process_name.lower():
                    return True
            except re.error as e:
                logging.warning(f'プロセス条件の正規表現パターン "{pattern}" が不正です: {e}')
                return False
        
        if class_pattern and class_name:
            try:
                if class_pattern.startswith("regex:"):
                    pattern = class_pattern.replace("regex:", "", 1)
                    if re.search(pattern, class_name, re.IGNORECASE if not condition.get("case_sensitive") else 0):
                        return True
                elif class_pattern.lower() in class_name.lower():
                    return True
            except re.error as e:
                logging.warning(f'クラス条件の正規表現パターン "{pattern}" が不正です: {e}')
                return False

        return False

    def _check_rule_conditions(self, window, process_name, class_name, rule_condition):
        """ルールの条件全体（AND/OR）をチェックする"""
        conditions = rule_condition.get("conditions")
        if not conditions:
            return self._check_single_condition(window, process_name, class_name, rule_condition)
        
        logic = rule_condition.get("logic", "AND").upper()
        try:
            if logic == "OR":
                return any(self._check_single_condition(window, process_name, class_name, c) for c in conditions)
            else: # AND
                return all(self._check_single_condition(window, process_name, class_name, c) for c in conditions)
        except Exception as e:
            logging.error(f"ルール条件の評価中にエラーが発生しました: {e}", exc_info=True)
            return False

    def handle_window_event(self, hwnd):
        """WinEventHookからのコールバック。ウィンドウを処理キューに追加する"""
        with self.lock:
            if self.is_paused:
                return
            # 既に処理済みのウィンドウは無視
            if hwnd in self.processed_windows:
                return
        
        try:
            # イベント発生直後はウィンドウ情報が不完全なことがあるため、少し待つ
            time.sleep(0.1)
            
            window = gw.Win32Window(hwnd)
            if not (window.visible and not window.isMinimized and window.title):
                return

            try:
                class_name = win32gui.GetClassName(hwnd)
            except win32gui.error:
                class_name = None # ウィンドウが既に存在しない場合など

            process_name = self._get_process_name(hwnd)
            logging.debug(f"イベント受信: タイトル='{window.title}', プロセス='{process_name}', クラス='{class_name}'")
            if process_name in IGNORED_PROCESSES:
                logging.info(f"プロセス '{process_name}' (ウィンドウタイトル: '{window.title}') は無視リストに含まれているため、処理をスキップします。")
                return

            for rule in self.settings.rules:
                if self._check_rule_conditions(window, process_name, class_name, rule.get("condition", {})):
                    with self.lock:
                        if hwnd in self.processed_windows: # ダブルチェック
                            continue
                        self.processed_windows.add(hwnd)
                    
                    # 非同期タスクをスレッドセーフに呼び出す
                    asyncio.run_coroutine_threadsafe(self._apply_rule_async(rule, window), self.loop)
                    break
        except gw.PyGetWindowException:
            # ウィンドウが既に閉じられている場合など
            pass
        except Exception as e:
            logging.error(f"ウィンドウイベント処理中にエラーが発生しました (HWND: {hwnd}): {e}", exc_info=True)

    def process_existing_windows(self):
        """起動時に既存のウィンドウをすべて処理する"""
        if not self.settings.globals.get("apply_on_startup", True):
            logging.info("起動時のルール適用はスキップします。")
            return
        
        logging.info("既存のウィンドウにルールを適用します...")
        try:
            for window in gw.getAllWindows():
                if window.visible and not window.isMinimized and window.title:
                    self.handle_window_event(window._hWnd)
        except Exception as e:
            logging.error(f"既存ウィンドウの処理中にエラー: {e}", exc_info=True)

    async def _apply_rule_async(self, rule, window):
        """非同期で単一のルールをウィンドウに適用する"""
        rule_name = rule.get("name", "無名ルール")
        action = rule.get("action", {})
        logging.info(f'非同期タスク開始: ルール "{rule_name}" を "{window.title}" に適用します。')

        try:
            delay_ms = action.get("execution_delay")
            if isinstance(delay_ms, int) and delay_ms > 0:
                logging.info(f" -> {delay_ms}ms の遅延を開始します。")
                await asyncio.sleep(delay_ms / 1000)
                logging.info(f" -> {delay_ms}ms の遅延が完了しました。")

            if not window.visible or window.isMinimized:
                logging.warning(f'遅延後、ウィンドウ "{window.title}" が非表示または最小化されたため、処理を中断します。')
                self._discard_window(window._hWnd)
                return

            with self.lock:
                if not window.visible or window.isMinimized:
                    self._discard_window(window._hWnd)
                    return

                target_workspace = action.get("target_workspace")
                if isinstance(target_workspace, int):
                    try:
                        num_desktops = len(get_virtual_desktops())
                        if 1 <= target_workspace <= num_desktops:
                            logging.info(f" -> 仮想デスクトップ {target_workspace} に移動します。")
                            AppView(hwnd=window._hWnd).move(VirtualDesktop(number=target_workspace))
                        else:
                            logging.warning(f"指定された仮想デスクトップ {target_workspace} は存在しません (利用可能なデスクトップ数: {num_desktops})。")
                    except Exception as e:
                        logging.error(f"仮想デスクトップの移動中にエラーが発生しました: {e}", exc_info=True)

                if action.get("maximize", "").upper() == "ON":
                    window.maximize()
                    logging.info(" -> ウィンドウを最大化しました。")
                elif action.get("minimize", "").upper() == "ON":
                    window.minimize()
                    logging.info(" -> ウィンドウを最小化しました。")
                else:
                    x, y, w, h = self.calculator.get_target_rect(action, window)
                    
                    current_left, current_top, current_width, current_height = window.left, window.top, window.width, window.height
                    
                    if w != current_width or h != current_height:
                        logging.info(f" -> サイズを {w}x{h} に変更します。")
                        window.resizeTo(w, h)
                    if x != current_left or y != current_top:
                        logging.info(f" -> 位置を ({x}, {y}) に移動します。")
                        window.moveTo(x, y)

        except gw.PyGetWindowException as e:
            logging.warning(f'ウィンドウ "{window.title}" の操作に失敗しました: {e}')
            self._discard_window(window._hWnd)
        except Exception as e:
            logging.error(f'ルール "{rule_name}" の適用中に予期せぬエラーが発生しました: {e}', exc_info=True)
            self._discard_window(window._hWnd)

    def _discard_window(self, hwnd):
        """処理済みセットからウィンドウを安全に削除する"""
        with self.lock:
            self.processed_windows.discard(hwnd)

# --- Win32 イベントフック (ctypes) ---
WINEVENTPROC = ctypes.WINFUNCTYPE(
    None, wintypes.HANDLE, wintypes.DWORD, wintypes.HWND, wintypes.LONG,
    wintypes.LONG, wintypes.DWORD, wintypes.DWORD)

class WinEventHook(threading.Thread):
    def __init__(self, callback):
        super().__init__(name="WinEventHookThread", daemon=True)
        self.callback = callback
        self.hook = None
        self.running = False
        self.user32 = ctypes.windll.user32
        self.event_proc_obj = WINEVENTPROC(self.event_proc)

    def run(self):
        """イベントフックを開始し、メッセージループに入る"""
        self.running = True
        try:
            self.hook = self.user32.SetWinEventHook(
                win32con.EVENT_OBJECT_CREATE,
                win32con.EVENT_OBJECT_SHOW,
                0, self.event_proc_obj, 0, 0, win32con.WINEVENT_OUTOFCONTEXT
            )
            if self.hook:
                logging.info("Win32 イベントフックを開始しました。")
                msg = wintypes.MSG()
                while self.user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
                    self.user32.TranslateMessage(ctypes.byref(msg))
                    self.user32.DispatchMessageW(ctypes.byref(msg))
            else:
                logging.error("Win32 イベントフックの開始に失敗しました。")
        except Exception as e:
            logging.critical(f"WinEventフックスレッドで致命的なエラー: {e}", exc_info=True)
        finally:
            logging.info("WinEventフックスレッドが終了しました。")

    def stop(self):
        """イベントフックを解除し、メッセージループを終了する"""
        if self.hook:
            self.user32.UnhookWinEvent(self.hook)
            self.hook = None
            logging.info("Win32 イベントフックを解除しました。")
        if self.running:
            self.user32.PostThreadMessageW(self.ident, win32con.WM_QUIT, 0, 0)
        self.running = False

    def event_proc(self, hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
        """イベントコールバック関数"""
        if dwEventThread == self.ident:
            return
        if idObject != win32con.OBJID_WINDOW or idChild != 0:
            return
        if not hwnd or not win32gui.IsWindow(hwnd) or win32gui.GetParent(hwnd) != 0:
            return
        try:
            self.callback(hwnd)
        except Exception as e:
            logging.error(f"イベント処理コールバックでエラー (HWND: {hwnd}): {e}", exc_info=True)

# --- システムトレイ ---
class Tray:
    def __init__(self, window_manager, win_event_hook):
        self.window_manager = window_manager
        self.win_event_hook = win_event_hook

        try:
            self.icon_running = Image.open("window_mover.ico")
            self.icon_paused = Image.open("window_mover_pause.ico")
        except FileNotFoundError as e:
            logging.warning(f"アイコンファイルが見つかりません ({e.filename})。デフォルトアイコンを生成します。")
            self.icon_running = self._create_default_image(64, 64, "black", "white")
            self.icon_paused = self._create_default_image(64, 64, "gray", "white")
        except Exception as e:
            logging.error(f"アイコンの読み込み中に予期せぬエラーが発生しました: {e}", exc_info=True)
            self.icon_running = self._create_default_image(64, 64, "black", "white")
            self.icon_paused = self._create_default_image(64, 64, "gray", "white")

        menu = pystray.Menu(
            pystray.MenuItem(
                lambda text: "再開" if self.window_manager.is_paused else "一時停止",
                self._toggle_pause_action
            ),
            pystray.MenuItem("設定を再読み込み", self._reload_settings_action),
            pystray.MenuItem("ログをクリア", self._clear_log_action),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("終了", self._exit_action)
        )
        self.icon = pystray.Icon("window_mover", self.icon_running, "Window Mover", menu)

    def _clear_log_action(self, icon, item):
        """ログファイルをクリアする"""
        with self.window_manager.lock:
            try:
                log_level_str = self.window_manager.settings.globals.get("log_level", "INFO")
                log_level = get_log_level_from_string(log_level_str)
                setup_logging(level=log_level)
                logging.info("ログファイルをクリアしました。")
            except Exception as e:
                print(f"ログファイルのクリア中にエラーが発生しました: {e}")
                logging.error(f"ログファイルのクリア中にエラーが発生しました: {e}", exc_info=True)

    def _toggle_pause_action(self, icon, item):
        """一時停止と再開を切り替える"""
        with self.window_manager.lock:
            self.window_manager.is_paused = not self.window_manager.is_paused
            if self.window_manager.is_paused:
                self.icon.icon = self.icon_paused
                logging.info("処理を一時停止しました。")
            else:
                self.icon.icon = self.icon_running
                logging.info("処理を再開しました。")
                if self.window_manager.settings.globals.get("apply_on_resume", True):
                    self.window_manager.processed_windows.clear()
                    logging.info("すべてのウィンドウにルールを再適用します。")
                    self.window_manager.process_existing_windows()
                else:
                    logging.info("新規ウィンドウのみルールを適用します（既存ウィンドウは対象外）。")

    def _create_default_image(self, width, height, color1, color2):
        """デフォルトのアイコン画像を生成する"""
        image = Image.new("RGB", (width, height), color1)
        dc = ImageDraw.Draw(image)
        dc.rectangle((width // 2, 0, width, height // 2), fill=color2)
        dc.rectangle((0, height // 2, width // 2, height), fill=color2)
        return image

    def _reload_settings_action(self, icon, item):
        """設定ファイルを再読み込みする"""
        logging.info("設定の再読み込みを開始します。")
        with self.window_manager.lock:
            try:
                self.window_manager.settings.load()
                try:
                    monitors = get_monitors()
                except Exception as e:
                    logging.error(f"モニター情報の取得に失敗しました: {e}")
                    return

                self.window_manager.calculator = Calculator(monitors, self.window_manager.settings.globals)
                logging.info(f"{len(monitors)}個のモニター情報を更新しました。")

                log_level_str = self.window_manager.settings.globals.get("log_level", "INFO")
                log_level = get_log_level_from_string(log_level_str)
                setup_logging(level=log_level)
                logging.info(f"ログレベルを「{log_level_str}」に設定しました。")

                if self.window_manager.settings.globals.get("apply_on_reload", True):
                    self.window_manager.processed_windows.clear()
                    logging.info("すべてのウィンドウにルールを再適用します。")
                    self.window_manager.process_existing_windows()
                else:
                    logging.info("新規ウィンドウのみルールを適用します（既存ウィンドウは対象外）。")
                
                logging.info("設定の再読み込みが完了しました。")

            except Exception as e:
                logging.error(f"設定の再読み込み中に予期せぬエラーが発生しました: {e}", exc_info=True)

    def _exit_action(self, icon, item):
        logging.info("アプリケーションを終了します。")
        self.win_event_hook.stop()
        icon.stop()

    def run(self):
        self.icon.run()

# --- 非同期処理スレッド ---
class AsyncWorker(threading.Thread):
    def __init__(self):
        super().__init__(name="AsyncWorkerThread", daemon=True)
        self.loop = asyncio.new_event_loop()

    def run(self):
        asyncio.set_event_loop(self.loop)
        try:
            self.loop.run_forever()
        finally:
            self.loop.close()
            logging.info("非同期ワーカースレッドが終了しました。")

    def stop(self):
        self.loop.call_soon_threadsafe(self.loop.stop)

# --- メイン処理 ---
def main():
    setup_logging()
    
    async_worker = None
    win_event_hook = None
    
    try:
        settings = Settings(SETTINGS_FILE)
        log_level_str = settings.globals.get("log_level", "INFO")
        log_level = get_log_level_from_string(log_level_str)
        setup_logging(level=log_level)
        logging.info(f"ログレベルを「{log_level_str}」に設定しました。")

        logging.info("アプリケーションを開始します。")

        # 非同期処理用のワーカースレッドを開始
        async_worker = AsyncWorker()
        async_worker.start()

        window_manager = WindowManager(settings, async_worker.loop)
        
        # Win32イベントフックを開始
        win_event_hook = WinEventHook(window_manager.handle_window_event)
        win_event_hook.start()
        
        # 起動時のウィンドウ処理
        # 少し待ってから実行しないと、イベントフックと競合する可能性がある
        async_worker.loop.call_later(1, window_manager.process_existing_windows)

        tray = Tray(window_manager, win_event_hook)
        tray.run() # これはブロッキング呼び出し

    except Exception as e:
        logging.critical(f"アプリケーションの起動中に致命的なエラーが発生しました: {e}", exc_info=True)
    finally:
        logging.info("アプリケーションのシャットダウン処理を開始します。")
        if win_event_hook and win_event_hook.is_alive():
            win_event_hook.stop()
        if async_worker and async_worker.is_alive():
            async_worker.stop()
        logging.info("アプリケーションが終了しました。")

if __name__ == "__main__":
    main()