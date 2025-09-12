
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
import re
from screeninfo import get_monitors
from pyvda import AppView, VirtualDesktop, get_virtual_desktops

# --- 定数 ---
SETTINGS_FILE = "settings.toml"
LOG_FILE = "log.txt"
ANCHOR_POINTS = {
    "TopLeft": (0.0, 0.0), "TopCenter": (0.5, 0.0), "TopRight": (1.0, 0.0),
    "MiddleLeft": (0.0, 0.5), "MiddleCenter": (0.5, 0.5), "MiddleRight": (1.0, 0.5),
    "BottomLeft": (0.0, 1.0), "BottomCenter": (0.5, 1.0), "BottomRight": (1.0, 1.0)
}

# --- ログレベル変換 ---
def get_log_level_from_string(level_str: str) -> int:
    """ログレベルの文字列をloggingの定数に変換する"""
    return getattr(logging, level_str.upper(), logging.INFO)

# --- ロギング設定 ---
def setup_logging(level: int = logging.INFO):
    """ロギングの基本設定を行う"""
    # 既存のハンドラをすべて削除し、設定の重複を防ぐ
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    
    # 新しい設定で再構成
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding='utf-8', mode='w'),
            logging.StreamHandler()
        ]
    )
    # ライブラリのデバッグログを抑制
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
            # 1. ターゲットモニターを決定
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

            # 2. モニターの作業領域を計算
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

            # 3. ウィンドウサイズを計算
            resize_to = rule_action.get("resize_to", {})
            w_val = resize_to.get("w", resize_to.get("width"))
            h_val = resize_to.get("h", resize_to.get("height"))
            width = self._parse_value(w_val, work_area_width) if w_val is not None else window.width
            height = self._parse_value(h_val, work_area_height) if h_val is not None else window.height
            if width is None or height is None:
                logging.warning("サイズ指定の解析に失敗したため、現在のサイズを維持します。")
                width = window.width if width is None else width
                height = window.height if height is None else height

            # 4. ターゲット座標を計算
            move_to = rule_action.get("move_to")
            anchor_name = rule_action.get("anchor", "TopLeft")
            anchor_x_ratio, anchor_y_ratio = ANCHOR_POINTS.get(anchor_name, (0.0, 0.0))
            logging.debug(f"ウィンドウのアンカー: {anchor_name} ({anchor_x_ratio}, {anchor_y_ratio})")

            base_x, base_y = None, None
            if isinstance(move_to, str): # アンカー指定
                target_anchor_name = move_to
                target_x_ratio, target_y_ratio = ANCHOR_POINTS.get(target_anchor_name, (0.0, 0.0))
                base_x = work_area_x + int(work_area_width * target_x_ratio)
                base_y = work_area_y + int(work_area_height * target_y_ratio)
                logging.debug(f"移動先アンカー '{target_anchor_name}' -> ベース座標 ({base_x}, {base_y})")
            elif isinstance(move_to, dict): # 絶対/相対座標指定
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

            # 5. ウィンドウ自体のアンカーを適用
            final_x -= int(width * anchor_x_ratio)
            final_y -= int(height * anchor_y_ratio)
            logging.debug(f"ウィンドウアンカー適用後 -> ({final_x}, {final_y})")

            # 6. ルール個別のオフセットを適用
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
    def __init__(self, settings):
        self.settings = settings
        try:
            monitors = get_monitors()
        except Exception as e:
            logging.critical(f"モニター情報の取得に失敗しました。アプリケーションを続行できません: {e}")
            raise
        self.calculator = Calculator(monitors, self.settings.globals)
        self.processed_windows = set()
        self.running_event = threading.Event()
        self.running_event.set() # 初期状態は実行中
        self.lock = threading.Lock()

        if not self.settings.globals.get("apply_on_startup", True):
            logging.info("起動時のルール適用はスキップします（既存のウィンドウは対象外）。")
            try:
                # 既存の可視ウィンドウをすべて処理済みにする
                self.processed_windows.update(w._hWnd for w in gw.getAllWindows() if w.visible)
            except Exception as e:
                logging.error(f"起動時の既存ウィンドウ取得中にエラーが発生しました: {e}", exc_info=True)

    def _get_process_name(self, hwnd):
        """ウィンドウハンドルからプロセス名を取得する"""
        try:
            pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value == 0:
                logging.debug(f"PIDが0のため、プロセス名を取得できませんでした (HWND: {hwnd})。")
                return None
            return psutil.Process(pid.value).name()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
            logging.warning(f"プロセス名の取得に失敗しました (PID: {pid.value}): {e}")
            return None
        except Exception as e:
            logging.error(f"プロセス名取得中に予期せぬエラー (PID: {pid.value}): {e}", exc_info=True)
            return None

    def _check_single_condition(self, window, process_name, condition):
        """単一の条件をチェックする"""
        title_pattern = condition.get("title")
        process_pattern = condition.get("process")
        
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
        return False

    def _check_rule_conditions(self, window, process_name, rule_condition):
        """ルールの条件全体（AND/OR）をチェックする"""
        conditions = rule_condition.get("conditions")
        if not conditions:
            return self._check_single_condition(window, process_name, rule_condition)
        
        logic = rule_condition.get("logic", "AND").upper()
        try:
            if logic == "OR":
                return any(self._check_single_condition(window, process_name, c) for c in conditions)
            else: # AND
                return all(self._check_single_condition(window, process_name, c) for c in conditions)
        except Exception as e:
            logging.error(f"ルール条件の評価中にエラーが発生しました: {e}", exc_info=True)
            return False

    def run(self):
        """
        スレッドのエントリポイント。
        このメソッド内で非同期イベントループを開始し、ウィンドウ監視のメインループを実行する。
        """
        try:
            # スレッド内で新しいイベントループを作成して実行
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            loop.run_until_complete(self._monitor_windows_async())
        except Exception as e:
            logging.critical(f"非同期ウィンドウ監視ループで致命的なエラーが発生しました: {e}", exc_info=True)

    async def _apply_rule_async(self, rule, window):
        """非同期で単一のルールをウィンドウに適用する"""
        rule_name = rule.get("name", "無名ルール")
        action = rule.get("action", {})
        logging.info(f'非同期タスク開始: ルール "{rule_name}" を "{window.title}" に適用します。')

        try:
            # --- 遅延処理 ---
            delay_ms = action.get("execution_delay")
            if isinstance(delay_ms, int) and delay_ms > 0:
                logging.info(f" -> {delay_ms}ms の遅延を開始します。")
                await asyncio.sleep(delay_ms / 1000)
                logging.info(f" -> {delay_ms}ms の遅延が完了しました。")

            # 遅延後にウィンドウがまだ有効か確認
            if not window.visible or window.isMinimized:
                logging.warning(f'遅延後、ウィンドウ "{window.title}" が非表示または最小化されたため、処理を中断します。')
                return

            # ウィンドウ操作はロックで保護する
            with self.lock:
                # 再度ウィンドウの状態をチェック
                if not window.visible or window.isMinimized:
                    return

                # --- 仮想デスクトップ移動 ---
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

                # --- ウィンドウ操作 ---
                if action.get("maximize", "").upper() == "ON":
                    window.maximize()
                    logging.info(" -> ウィンドウを最大化しました。")
                elif action.get("minimize", "").upper() == "ON":
                    window.minimize()
                    logging.info(" -> ウィンドウを最小化しました。")
                else:
                    x, y, w, h = self.calculator.get_target_rect(action, window)
                    
                    # PyGetWindowのプロパティアクセスは遅い可能性があるので、一度だけアクセスする
                    current_left, current_top, current_width, current_height = window.left, window.top, window.width, window.height
                    
                    if w != current_width or h != current_height:
                        logging.info(f" -> サイズを {w}x{h} に変更します。")
                        window.resizeTo(w, h)
                    if x != current_left or y != current_top:
                        logging.info(f" -> 位置を ({x}, {y}) に移動します。")
                        window.moveTo(x, y)

        except gw.PyGetWindowException as e:
            logging.warning(f'ウィンドウ "{window.title}" の操作に失敗しました: {e}')
            with self.lock:
                self.processed_windows.discard(window._hWnd)
        except Exception as e:
            logging.error(f'ルール "{rule_name}" の適用中に予期せぬエラーが発生しました: {e}', exc_info=True)
            with self.lock:
                self.processed_windows.discard(window._hWnd)

    async def _monitor_windows_async(self):
        """ウィンドウを定期的に監視し、ルールを適用する非同期メインループ"""
        while True:
            if not self.running_event.is_set():
                await asyncio.sleep(0.5) # 一時停止中は少し長めに待つ
                continue

            polling_interval_ms = self.settings.globals.get("polling_interval", 1000)
            
            try:
                all_windows = gw.getAllWindows()
                
                with self.lock:
                    current_handles = {w._hWnd for w in all_windows}

                    for window in all_windows:
                        if not (window.visible and not window.isMinimized and window.title and window._hWnd not in self.processed_windows):
                            continue

                        process_name = self._get_process_name(window._hWnd)
                        
                        for rule in self.settings.rules:
                            if self._check_rule_conditions(window, process_name, rule.get("condition", {})):
                                self.processed_windows.add(window._hWnd)
                                asyncio.create_task(self._apply_rule_async(rule, window))
                                break 
                    
                    self.processed_windows.intersection_update(current_handles)

            except Exception as e:
                logging.error(f"ウィンドウ処理のメインループで予期せぬエラーが発生しました: {e}", exc_info=True)

            await asyncio.sleep(max(0.1, polling_interval_ms / 1000))

# --- システムトレイ ---
class Tray:
    def __init__(self, window_manager):
        self.window_manager = window_manager
        self.is_paused = False

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
                lambda text: "再開" if self.is_paused else "一時停止",
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
                # 現在のログレベルを維持してロガーを再セットアップ（これによりファイルが 'w' モードで開かれる）
                log_level_str = self.window_manager.settings.globals.get("log_level", "INFO")
                log_level = get_log_level_from_string(log_level_str)
                setup_logging(level=log_level)
                logging.info("ログファイルをクリアしました。")
            except Exception as e:
                # ロギングが機能しない可能性を考慮し、printでも出力
                print(f"ログファイルのクリア中にエラーが発生しました: {e}")
                logging.error(f"ログファイルのクリア中にエラーが発生しました: {e}", exc_info=True)

    def _toggle_pause_action(self, icon, item):
        """一時停止と再開を切り替える"""
        with self.window_manager.lock:
            self.is_paused = not self.is_paused
            if self.is_paused:
                self.window_manager.running_event.clear()
                self.icon.icon = self.icon_paused
                logging.info("処理を一時停止しました。")
            else:
                self.window_manager.running_event.set()
                self.icon.icon = self.icon_running
                logging.info("処理を再開しました。")
                if self.window_manager.settings.globals.get("apply_on_resume", True):
                    self.window_manager.processed_windows.clear()
                    logging.info("すべてのウィンドウにルールを再適用します。")
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
                # 1. 設定ファイルをリロード
                self.window_manager.settings.load()

                # 2. モニター情報を再取得
                try:
                    monitors = get_monitors()
                except Exception as e:
                    logging.error(f"モニター情報の取得に失敗しました: {e}")
                    return # モニター情報がなければ続行不可

                # 3. 新しい設定とモニター情報でCalculatorを再生成
                self.window_manager.calculator = Calculator(monitors, self.window_manager.settings.globals)
                logging.info(f"{len(monitors)}個のモニター情報を更新しました。")

                # 4. 新しいログレベルを適用
                log_level_str = self.window_manager.settings.globals.get("log_level", "INFO")
                log_level = get_log_level_from_string(log_level_str)
                setup_logging(level=log_level)
                logging.info(f"ログレベルを「{log_level_str}」に設定しました。")

                # 5. ルール再適用の設定
                if self.window_manager.settings.globals.get("apply_on_reload", True):
                    self.window_manager.processed_windows.clear()
                    logging.info("すべてのウィンドウにルールを再適用します。")
                else:
                    logging.info("新規ウィンドウのみルールを適用します（既存ウィンドウは対象外）。")
                
                logging.info("設定の再読み込みが完了しました。")

            except Exception as e:
                logging.error(f"設定の再読み込み中に予期せぬエラーが発生しました: {e}", exc_info=True)

    def _exit_action(self, icon, item):
        logging.info("アプリケーションを終了します。")
        icon.stop()

    def run(self):
        self.icon.run()

# --- メイン処理 ---
if __name__ == "__main__":
    # 起動時にまずデフォルトレベルでロガーを初期化
    # これにより、設定ファイル読み込み前のエラーも記録できる
    setup_logging()

    try:
        settings = Settings(SETTINGS_FILE)
        
        # 設定ファイルで指定されたレベルでロガーを再セットアップ
        log_level_str = settings.globals.get("log_level", "INFO")
        log_level = get_log_level_from_string(log_level_str)
        setup_logging(level=log_level)
        logging.info(f"ログレベルを「{log_level_str}」に設定しました。")


        logging.info("アプリケーションを開始します。")

        window_manager = WindowManager(settings)
        
        monitor_thread = threading.Thread(target=window_manager.run, daemon=True)
        monitor_thread.start()
        logging.info("ウィンドウの監視を開始しました。")

        tray = Tray(window_manager)
        tray.run()

    except Exception as e:
        logging.critical(f"アプリケーションの起動中に致命的なエラーが発生しました: {e}", exc_info=True)
    finally:
        logging.info("アプリケーションが終了しました。")
