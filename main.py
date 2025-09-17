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
import sys
import subprocess
import ctypes
from ctypes import wintypes
import re
from screeninfo import get_monitors
from pyvda import AppView, VirtualDesktop, get_virtual_desktops
import win32gui
import win32con
from pydantic import ValidationError

from settings_model import SettingsModel, MoveTo, ResizeTo

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
        self.model: SettingsModel = SettingsModel()
        self.load()

    def load(self):
        """設定ファイルを読み込み、Pydanticモデルで検証する"""
        try:
            with open(self.filepath, "r", encoding="utf-8") as f:
                data = toml.load(f)
            self.model = SettingsModel.model_validate(data)
            logging.info(f"設定ファイルを読み込み、検証しました。Global: {len(self.model.globals.model_dump())}項目, Ignores: {len(self.model.ignores)}個, Rules: {len(self.model.rules)}個")
        except FileNotFoundError:
            logging.warning(f"設定ファイル '{self.filepath}' が見つかりません。デフォルト設定で新しいファイルを生成します。")
            self._create_default_settings_file()
            self.model = SettingsModel() # 生成後はデフォルト設定で動作
        except toml.TomlDecodeError as e:
            logging.error(f"設定ファイル '{self.filepath}' の形式が正しくありません (行: {e.lineno}, 列: {e.colno}): {e}")
            self.model = SettingsModel()
        except ValidationError as e:
            logging.error(f"設定ファイル '{self.filepath}' のバリデーションに失敗しました:")
            for error in e.errors():
                logging.error(f"  -場所: {' -> '.join(map(str, error['loc']))}, エラー: {error['msg']}")
            self.model = SettingsModel()
        except Exception as e:
            logging.error(f"設定ファイルの読み込み中に予期せぬエラーが発生しました: {e}", exc_info=True)
            self.model = SettingsModel()

    def _create_default_settings_file(self):
        """デフォルトの設定ファイルを作成する"""
        default_content = """
#==============================================================================
# Window Mover 設定ファイル
#
# このファイルでは、ウィンドウをどのように自動で移動・リサイズするかのルールを
# 定義します。設定は大きく分けて以下の3つのセクションから構成されます。
#
# 1. [global]: アプリケーション全体の動作設定
# 2. [[ignores]]: ルール適用の対象外とするウィンドウの指定
# 3. [[rules]]: ウィンドウを操作するための具体的なルールのリスト
#
# ルールは上から下に順番に評価され、最初に条件が一致したルールが適用されます。
#==============================================================================

#==============================================================================
# 1. [global] - 全体設定
# アプリケーション全体の挙動をここで定義します。
#==============================================================================
[global]
  # --- ログ設定 ---
  # ログファイル(log.txt)に出力する情報の詳細度を設定します。
  # "DEBUG": 最も詳細なログ。ルールの動作を細かく確認したい場合に便利です。
  # "INFO": 標準的なログ。どのウィンドウにどのルールが適用されたかなどが記録されます。
  # "WARNING": 警告メッセージのみ。
  # "ERROR": エラーメッセージのみ。
  log_level = "INFO"

  # --- ルール適用のタイミング ---
  # trueに設定すると、以下のタイミングで既存の全ウィンドウに対してルールを再適用します。
  apply_on_startup = true   # アプリケーション起動時
  apply_on_reload = true    # 設定ファイル再読み込み時
  apply_on_resume = false   # 一時停止から再開した時

  # --- タイトル変更時の再評価 ---
  # trueに設定すると、ウィンドウのタイトルが変更された際にルールを再評価します。
  # これにより、ブラウザのタブを切り替えた際などに、タイトルに応じた別のルールを適用できます。
  recheck_on_title_change = false

  # --- クリーンアップ設定 ---
  # 閉じたウィンドウの情報を内部リストから削除する頻度を秒単位で指定します。
  cleanup_interval_seconds = 300

  # --- モニターオフセット設定 ---
  # タスクバーなど、ウィンドウを配置したくない領域をモニターごとに予約します。
  # ここで設定した値は、`move_to`でアンカーポイント("MiddleCenter"など)を指定した際の
  # 「作業領域」の計算に使われます。
  [global.monitor_offsets]
    # `monitor_N`の指定がないモニターに適用されるデフォルト値
    default = { top = 0, bottom = 0, left = 0, right = 0 }

    # モニター1番に対する設定 (例: 下部に48pxのタスクバーがある場合)
    monitor_1 = { top = 0, bottom = 48, left = 0, right = 0 }

    # モニター2番に対する設定 (例: 上部に32px、左側に64pxの領域を予約)
    # monitor_2 = { top = 32, bottom = 0, left = 64, right = 0 }


#==============================================================================
# 2. [[ignores]] - 無視ルール
# ここで指定した条件に一致するウィンドウは、以降の[[rules]]で処理されません。
# OSのUIコンポーネントなど、動かしたくないウィンドウを指定するのに役立ちます。
#==============================================================================
[[ignores]]
  name = "OS標準UIと特定のアプリを無視"
  
  # "OR": いずれかの条件に一致すれば無視
  # "AND": すべての条件に一致した場合のみ無視
  logic = "OR"
  
  conditions = [
    # --- プロセス名で無視 ---
    { process = "SystemSettings.exe" },      # Windows 設定
    { process = "TextInputHost.exe" },       # テキスト入力UI
    { process = "SearchHost.exe" },          # 検索UI
    { process = "ShellExperienceHost.exe" }, # シェルUI
    { process = "StartMenuExperienceHost.exe" }, # スタートメニュー
    
    # --- ウィンドウクラス名で無視 ---
    { class = "Shell_TrayWnd" },        # タスクバー
    { class = "VirtualDesktopHotkeySwitcher" }, # 仮想デスクトップスイッチャー
  ]


#==============================================================================
# 3. [[rules]] - 個別ルール
# ウィンドウをどのように配置するかを具体的に定義します。
# ルールはこのファイルの上から順にチェックされ、最初に一致したものが適用されます。
#==============================================================================

# --- サンプルルール 1: 基本的な配置 (電卓) ---
# ウィンドウタイトルで「電卓」を識別し、画面中央に固定サイズで配置します。
[[rules]]
  name = "電卓を中央に"

  [rules.condition]
    title = "電卓" # ウィンドウタイトルに「電卓」が含まれていれば一致

  [rules.action]
    anchor = "MiddleCenter"  # ウィンドウの中心を基点に
    move_to = "MiddleCenter" # モニターの作業領域の中心に移動
    resize_to = { width = 320, height = 480 } # サイズを 320x480 ピクセルに

# --- サンプルルール 2: 正規表現と遅延実行 (メモ帳) ---
# 正規表現を使って複数のタイトルパターンに一致させ、3秒待ってから処理します。
[[rules]]
  name = "新規のメモ帳を左上に"

  [rules.condition]
    logic = "OR" # いずれかのタイトルに一致すればOK
    conditions = [
      { title = "regex:^タイトルなし - メモ帳$" }, # タイトルが完全に一致
      { title = "regex:^無題 - メモ帳$" }
    ]

  [rules.action]
    # anchorのデフォルトは"TopLeft"なので、ウィンドウの左上が基点になります。
    # `move_to`を座標で指定すると、monitor_offsetsは無視され、モニターの左上隅が(0,0)となります。
    move_to = { x = 10, y = 10 } # モニターの左上から (10px, 10px) の位置へ
    resize_to = { width = "40%", height = "70%" } # モニターの作業領域に対して40% x 70%のサイズに
    execution_delay = 1500 # 1500ms = 1.5秒待ってから実行

# --- サンプルルール 3: プロセス名、別モニター、オフセット (ペイント) ---
# プロセス名でペイントを識別し、2番モニターの右下に、少し余白を空けて配置します。
[[rules]]
  name = "ペイントを2番モニターの右下へ"

  [rules.condition]
    # プロセス名が "mspaint.exe" または "pbrush.exe" に一致
    process = "regex:(mspaint|pbrush)\\.exe"
    case_sensitive = false # 大文字・小文字を区別しない (デフォルト)

  [rules.action]
    target_monitor = 2       # 2番目のモニターを対象に
    anchor = "BottomRight"   # ウィンドウの右下を基点に
    move_to = "BottomRight"  # 対象モニターの作業領域の右下に移動
    offset = { x = -10, y = -10 } # そこからさらに (-10px, -10px) ずらす

# --- サンプルルール 4: 最大化・最小化 ---
# 特定のウィンドウを強制的に最大化、または最小化します。
[[rules]]
  name = "特定のアプリを最大化"
  [rules.condition]
    title = "MaximizeApp"
  [rules.action]
    maximize = "ON"

[[rules]]
  name = "特定のアプリを最小化"
  [rules.condition]
    title = "MinimizeApp"
  [rules.action]
    minimize = "ON"

# --- サンプルルール 5: 仮想デスクトップへの移動 ---
# 特定のウィンドウを指定した仮想デスクトップへ移動します。
[[rules]]
  name = "特定のウィンドウを仮想デスクトップ2へ"
  [rules.condition]
    title = "仮想デスクトップ2へ移動"
  [rules.action]
    target_workspace = 2 # 2番の仮想デスクトップへ移動
    # move_toやresize_toと組み合わせることも可能です
    anchor = "TopLeft"
    move_to = "TopLeft"

# --- サンプルルール 6: ウィンドウクラス名での指定 ---
# ゲームや一部のアプリケーションでは、ウィンドウクラス名での指定が有効な場合があります。
# (クラス名は`log.txt`をDEBUGレベルにして確認できます)
[[rules]]
  name = "ウィンドウクラス名での指定"
  [rules.condition]
    class = "ApplicationClassName"
  [rules.action]
    anchor = "MiddleCenter"
    move_to = "MiddleCenter"

"""
        try:
            with open(self.filepath, "w", encoding="utf-8") as f:
                f.write(default_content)
            logging.info(f"デフォルトの設定ファイルを '{self.filepath}' に生成しました。")
        except Exception as e:
            logging.error(f"デフォルト設定ファイルの生成に失敗しました: {e}", exc_info=True)

    @property
    def globals(self):
        return self.model.globals.model_dump()

    @property
    def rules(self):
        return [rule.model_dump() for rule in self.model.rules]

    @property
    def ignores(self):
        return [ignore.model_dump() for ignore in self.model.ignores]

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

    def _get_target_monitor(self, rule_action, window):
        """ルールとウィンドウ情報に基づき、ターゲットモニターを決定する"""
        target_monitor_num = rule_action.get("target_monitor")
        if target_monitor_num is not None:
            if isinstance(target_monitor_num, int) and 1 <= target_monitor_num <= len(self.monitors):
                monitor = self.monitors[target_monitor_num - 1]
                logging.debug(f"ルール指定により、ターゲットモニター {target_monitor_num} を使用します。")
                return monitor
            else:
                logging.warning(f"指定されたターゲットモニター番号 '{target_monitor_num}' は無効です（有効範囲: 1～{len(self.monitors)}）。フォールバックしてモニターを自動検出します。")
        
        return self.get_window_monitor(window)

    def _get_work_area(self, monitor, is_absolute_move):
        """モニターの作業領域を計算する"""
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
        return work_area_x, work_area_y, work_area_width, work_area_height

    def _calculate_new_size(self, resize_to, work_area_width, work_area_height, window):
        """新しいウィンドウサイズを計算する"""
        w_val = resize_to.get("width")
        h_val = resize_to.get("height")
        width = self._parse_value(w_val, work_area_width) if w_val is not None else window.width
        height = self._parse_value(h_val, work_area_height) if h_val is not None else window.height
        if width is None or height is None:
            logging.warning("サイズ指定の解析に失敗したため、現在のサイズを維持します。")
            width = window.width if width is None else width
            height = window.height if height is None else height
        return width, height

    def _calculate_new_position(self, move_to, work_area_x, work_area_y, work_area_width, work_area_height, monitor):
        """新しいウィンドウの基準位置を計算する"""
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
        return base_x, base_y

    def get_target_rect(self, rule_action, window):
        """ルールとウィンドウ情報に基づき、最終的な座標とサイズ (x, y, w, h) を計算する"""
        try:
            monitor = self._get_target_monitor(rule_action, window)
            is_absolute_move = isinstance(rule_action.get("move_to"), dict)
            
            work_area_x, work_area_y, work_area_width, work_area_height = self._get_work_area(monitor, is_absolute_move)

            resize_to = rule_action.get("resize_to") or {}
            width, height = self._calculate_new_size(resize_to, work_area_width, work_area_height, window)

            move_to = rule_action.get("move_to")
            base_x, base_y = self._calculate_new_position(move_to, work_area_x, work_area_y, work_area_width, work_area_height, monitor)

            if base_x is None and base_y is None:
                logging.debug("移動指定がないため、現在の位置を基準とします。")
                return window.left, window.top, width, height
            
            final_x = base_x if base_x is not None else window.left
            final_y = base_y if base_y is not None else window.top

            anchor_name = rule_action.get("anchor", "TopLeft")
            anchor_x_ratio, anchor_y_ratio = ANCHOR_POINTS.get(anchor_name, (0.0, 0.0))
            logging.debug(f"ウィンドウのアンカー: {anchor_name} ({anchor_x_ratio}, {anchor_y_ratio})")
            final_x -= int(width * anchor_x_ratio)
            final_y -= int(height * anchor_y_ratio)
            logging.debug(f"ウィンドウアンカー適用後 -> ({final_x}, {final_y})")

            rule_offset = rule_action.get("offset") or {}
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
        # 処理済みウィンドウを、適用されたルール名と共に辞書で管理する
        self.processed_windows = {}
        self.is_paused = False
        self.lock = threading.Lock()

        # クリーンアップタスクをスケジュールする
        asyncio.run_coroutine_threadsafe(self._cleanup_processed_windows_periodically(), self.loop)

    def clear_log(self):
        """ログファイルをクリアする"""
        with self.lock:
            try:
                log_level_str = self.settings.globals.get("log_level", "INFO")
                log_level = get_log_level_from_string(log_level_str)
                setup_logging(level=log_level)
                logging.info("ログファイルをクリアしました。")
            except Exception as e:
                print(f"ログファイルのクリア中にエラーが発生しました: {e}")
                logging.error(f"ログファイルのクリア中にエラーが発生しました: {e}", exc_info=True)

    def toggle_pause(self):
        """一時停止と再開を切り替える"""
        with self.lock:
            self.is_paused = not self.is_paused
            paused = self.is_paused
        
        if paused:
            logging.info("処理を一時停止しました。")
        else:
            logging.info("処理を再開しました。")
            if self.settings.globals.get("apply_on_resume", True):
                self.processed_windows.clear()
                logging.info("すべてのウィンドウにルールを再適用します。")
                self.process_existing_windows()
            else:
                logging.info("新規ウィンドウのみルールを適用します（既存ウィンドウは対象外）。")
        return paused

    def reload_settings(self):
        """設定ファイルを再読み込みし、必要に応じてルールを再適用する"""
        logging.info("設定の再読み込みを開始します。")
        
        apply_on_reload_flag = False
        try:
            with self.lock:
                self.settings.load()
                monitors = get_monitors()
                self.calculator = Calculator(monitors, self.settings.globals)
                logging.info(f"{len(monitors)}個のモニター情報を更新しました。")

                log_level_str = self.settings.globals.get("log_level", "INFO")
                log_level = get_log_level_from_string(log_level_str)
                setup_logging(level=log_level)
                logging.info(f"ログレベルを「{log_level_str}」に設定しました。")
                
                apply_on_reload_flag = self.settings.globals.get("apply_on_reload", True)

        except Exception as e:
            logging.error(f"設定の再読み込み中に予期せぬエラーが発生しました: {e}", exc_info=True)
            return

        if apply_on_reload_flag:
            logging.info("すべてのウィンドウにルールを再適用します。")
            self.processed_windows.clear()
            threading.Thread(target=self.process_existing_windows, daemon=True).start()
        else:
            logging.info("新規ウィンドウのみルールを適用します（既存ウィンドウは対象外）。")
        
        logging.info("設定の再読み込みが完了しました。")

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
        """
        単一の条件ブロックをチェックする。
        ブロック内に複数の条件（title, processなど）がある場合、それらはANDとして評価される。
        何も条件が指定されていない場合は、Falseを返す。
        """
        title_pattern = condition.get("title")
        process_pattern = condition.get("process")
        class_pattern = condition.get("class_name")
        case_sensitive = condition.get("case_sensitive", False)

        if not title_pattern and not process_pattern and not class_pattern:
            return False

        if title_pattern:
            title_matched = False
            try:
                if title_pattern.startswith("regex:"):
                    pattern = title_pattern.replace("regex:", "", 1)
                    if re.search(pattern, window.title, 0 if case_sensitive else re.IGNORECASE):
                        title_matched = True
                elif case_sensitive:
                    if title_pattern in window.title:
                        title_matched = True
                else:
                    if title_pattern.lower() in window.title.lower():
                        title_matched = True
            except re.error as e:
                logging.warning(f'タイトル条件の正規表現 "{title_pattern}" が不正です: {e}')
            
            if not title_matched:
                return False

        if process_pattern:
            process_matched = False
            if process_name:
                try:
                    if process_pattern.startswith("regex:"):
                        pattern = process_pattern.replace("regex:", "", 1)
                        if re.fullmatch(pattern, process_name, 0 if case_sensitive else re.IGNORECASE):
                            process_matched = True
                    elif case_sensitive:
                        if process_pattern == process_name:
                            process_matched = True
                    else:
                        if process_pattern.lower() == process_name.lower():
                            process_matched = True
                except re.error as e:
                    logging.warning(f'プロセス条件の正規表現 "{process_pattern}" が不正です: {e}')

            if not process_matched:
                return False

        if class_pattern:
            class_matched = False
            if class_name:
                try:
                    if class_pattern.startswith("regex:"):
                        pattern = class_pattern.replace("regex:", "", 1)
                        if re.search(pattern, class_name, 0 if case_sensitive else re.IGNORECASE):
                            class_matched = True
                    elif case_sensitive:
                        if class_pattern in class_name:
                            class_matched = True
                    else:
                        if class_pattern.lower() in class_name.lower():
                            class_matched = True
                except re.error as e:
                    logging.warning(f'クラス条件の正規表現 "{class_pattern}" が不正です: {e}')
            
            if not class_matched:
                return False

        return True

    def _check_rule_conditions(self, window, process_name, class_name, rule_condition):
        """ルールの条件全体（AND/OR）をチェックする"""
        conditions = rule_condition.get("conditions")
        if not conditions:
            return self._check_single_condition(window, process_name, class_name, rule_condition)
        
        logic = rule_condition.get("logic", "AND").upper()
        try:
            if logic == "OR":
                return any(self._check_single_condition(window, process_name, class_name, c) for c in conditions)
            else:
                return all(self._check_single_condition(window, process_name, class_name, c) for c in conditions)
        except Exception as e:
            logging.error(f"ルール条件の評価中にエラーが発生しました: {e}", exc_info=True)
            return False

    def handle_window_event(self, hwnd, event):
        """WinEventHookからのコールバック。イベントタイプに応じてウィンドウを処理する"""
        with self.lock:
            if self.is_paused:
                return

            is_title_change_event = (event == win32con.EVENT_OBJECT_NAMECHANGE)
            
            if is_title_change_event and not self.settings.globals.get("recheck_on_title_change", False):
                return

            previously_applied_rule = self.processed_windows.get(hwnd)

            if not is_title_change_event and previously_applied_rule is not None:
                return

        window = None
        for attempt in range(3):
            time.sleep(0.02)
            try:
                temp_window = gw.Win32Window(hwnd)
                if temp_window.visible and not temp_window.isMinimized and temp_window.title:
                    window = temp_window
                    break
            except gw.PyGetWindowException:
                return
            except Exception:
                pass

        if not window:
            return
            
        try:
            if not (window.visible and not window.isMinimized and window.title):
                return

            try:
                class_name = win32gui.GetClassName(hwnd)
            except win32gui.error:
                class_name = None

            process_name = self._get_process_name(hwnd)
            
            event_name = "作成/表示" if not is_title_change_event else "タイトル変更"
            logging.debug(f"イベント受信 ({event_name}): タイトル='{window.title}', プロセス='{process_name}', クラス='{class_name}'")

            # 無視ルールは常に最優先
            for ignore_rule in self.settings.ignores:
                if self._check_rule_conditions(window, process_name, class_name, ignore_rule):
                    ignore_name = ignore_rule.get("name", "無名無視ルール")
                    logging.info(f"無視ルール '{ignore_name}' に一致したため、ウィンドウ '{window.title}' の処理をスキップします。")
                    with self.lock:
                        self.processed_windows[hwnd] = "ignored" # 無視したことも記録
                    return

            # ルール評価
            matched_rule = None
            for rule in self.settings.rules:
                if self._check_rule_conditions(window, process_name, class_name, rule.get("condition", {})):
                    matched_rule = rule
                    break
            
            if matched_rule:
                rule_name = matched_rule.get("name", "無名ルール")
                # タイトル変更イベントで、かつ前回適用されたルールと同じ場合はアクションをスキップ
                if is_title_change_event and previously_applied_rule == rule_name:
                    logging.debug(f"タイトルは変更されましたが、前回と同じルール '{rule_name}' にマッチしたため、アクションは再実行しません。")
                    return

                # 新規適用、または別のルールへの変更
                log_prefix = "新規ルール適用:" if not previously_applied_rule else f"ルール変更 ({previously_applied_rule} -> {rule_name}):"
                logging.info(f"{log_prefix} '{window.title}' にルール '{rule_name}' を適用します。")
                
                with self.lock:
                    self.processed_windows[hwnd] = rule_name
                
                asyncio.run_coroutine_threadsafe(self._apply_rule_async(matched_rule, window), self.loop)

            elif previously_applied_rule:
                # どのルールにもマッチしなくなった場合
                logging.info(f"ウィンドウ '{window.title}' はどのルールにもマッチしなくなったため、追跡を解除します。(旧ルール: {previously_applied_rule})")
                self._discard_window(hwnd)

        except gw.PyGetWindowException:
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
                    # 既存ウィンドウは新規作成イベントとして扱う
                    self.handle_window_event(window._hWnd, win32con.EVENT_OBJECT_CREATE)
        except Exception as e:
            logging.error(f"既存ウィンドウの処理中にエラー: {e}", exc_info=True)

    async def _apply_rule_async(self, rule, window):
        """非同期で単一のルールをウィンドウに適用する"""
        rule_name = rule.get("name", "無名ルール")
        action = rule.get("action", {})
        
        try:
            # execution_delay はアクションの実行を遅延させる
            delay_ms = action.get("execution_delay")
            if isinstance(delay_ms, int) and delay_ms > 0:
                logging.info(f" -> アクション実行を {delay_ms}ms 遅延します。")
                await asyncio.sleep(delay_ms / 1000)
                logging.info(f" -> {delay_ms}ms の遅延が完了しました。")

            if not win32gui.IsWindow(window._hWnd) or not window.visible or window.isMinimized:
                logging.warning(f'アクション実行前にウィンドウ "{window.title}" が無効になったため、処理を中断します。')
                self._discard_window(window._hWnd)
                return

            with self.lock:
                if not win32gui.IsWindow(window._hWnd) or not window.visible or window.isMinimized:
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
                elif action.get("move_to") or action.get("resize_to"):
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
        """処理済み辞書からウィンドウを安全に削除する"""
        with self.lock:
            self.processed_windows.pop(hwnd, None)

    async def _cleanup_processed_windows_periodically(self):
        """processed_windows 辞書を定期的にクリーンアップする"""
        cleanup_interval = self.settings.globals.get("cleanup_interval_seconds", 300)
        logging.info(f"{cleanup_interval}秒ごとに無効なウィンドウハンドルのクリーンアップを実行します。")
        
        while True:
            await asyncio.sleep(cleanup_interval)
            with self.lock:
                if not self.processed_windows:
                    continue

                logging.debug(f"クリーンアップ開始: 現在 {len(self.processed_windows)}個のウィンドウを追跡中。")
                
                invalid_hwnds = {hwnd for hwnd in self.processed_windows if not win32gui.IsWindow(hwnd)}
                
                if invalid_hwnds:
                    for hwnd in invalid_hwnds:
                        self.processed_windows.pop(hwnd, None)
                    logging.info(f"{len(invalid_hwnds)}個の無効なウィンドウハンドルをクリーンアップしました。")

# --- Win32 イベントフック (ctypes) ---
WINEVENTPROC = ctypes.WINFUNCTYPE(
    None, wintypes.HANDLE, wintypes.DWORD, wintypes.HWND, wintypes.LONG,
    wintypes.LONG, wintypes.DWORD, wintypes.DWORD)

class WinEventHook(threading.Thread):
    def __init__(self, callback):
        super().__init__(name="WinEventHookThread", daemon=True)
        self.callback = callback
        self.hooks = []
        self.running = False
        self.user32 = ctypes.windll.user32
        self.event_proc_obj = WINEVENTPROC(self.event_proc)

    def run(self):
        """イベントフックを開始し、メッセージループに入る"""
        self.running = True
        try:
            # 複数のイベントフックをセットアップ
            self.hooks.append(self.user32.SetWinEventHook(
                win32con.EVENT_OBJECT_CREATE, win32con.EVENT_OBJECT_SHOW,
                0, self.event_proc_obj, 0, 0, win32con.WINEVENT_OUTOFCONTEXT | win32con.WINEVENT_SKIPOWNPROCESS))
            
            self.hooks.append(self.user32.SetWinEventHook(
                win32con.EVENT_OBJECT_NAMECHANGE, win32con.EVENT_OBJECT_NAMECHANGE,
                0, self.event_proc_obj, 0, 0, win32con.WINEVENT_OUTOFCONTEXT | win32con.WINEVENT_SKIPOWNPROCESS))

            if not all(self.hooks):
                logging.error("Win32 イベントフックの開始に失敗しました。")
                return

            logging.info("Win32 イベントフックを開始しました。(CREATE, SHOW, NAMECHANGE)")
            msg = wintypes.MSG()
            while self.user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
                self.user32.TranslateMessage(ctypes.byref(msg))
                self.user32.DispatchMessageW(ctypes.byref(msg))

        except Exception as e:
            logging.critical(f"WinEventフックスレッドで致命的なエラー: {e}", exc_info=True)
        finally:
            logging.info("WinEventフックスレッドが終了しました。")

    def stop(self):
        """イベントフックを解除し、メッセージループを終了する"""
        for h in self.hooks:
            if h:
                self.user32.UnhookWinEvent(h)
        self.hooks = []
        logging.info("すべてのWin32 イベントフックを解除しました。")

        if self.running:
            self.user32.PostThreadMessageW(self.ident, win32con.WM_QUIT, 0, 0)
        self.running = False

    def event_proc(self, hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
        """イベントコールバック関数"""
        if idObject != win32con.OBJID_WINDOW or idChild != 0:
            return
        if not hwnd or not win32gui.IsWindow(hwnd) or win32gui.GetParent(hwnd) != 0:
            return
        try:
            # コールバックにイベントタイプを渡す
            self.callback(hwnd, event)
        except Exception as e:
            logging.error(f"イベント処理コールバックでエラー (HWND: {hwnd}, Event: {event}): {e}", exc_info=True)

# --- システムトレイ ---
class Tray:
    def __init__(self, window_manager, win_event_hook, application_path):
        self.window_manager = window_manager
        self.win_event_hook = win_event_hook

        try:
            self.icon_running = Image.open(os.path.join(application_path, "window_mover.ico"))
            self.icon_paused = Image.open(os.path.join(application_path, "window_mover_pause.ico"))
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
            pystray.MenuItem("ログを開く", self._open_log_action),
            pystray.MenuItem("ログをクリア", self._clear_log_action),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("終了", self._exit_action)
        )
        self.icon = pystray.Icon("window_mover", self.icon_running, "Window Mover", menu)

    def _open_log_action(self, icon, item):
        """ログファイルをnotepad.exeで開く"""
        try:
            # LOG_FILEはmain()で絶対パスに更新されているグローバル変数
            if os.path.exists(LOG_FILE):
                subprocess.run(["notepad.exe", LOG_FILE])
                logging.info(f"ログファイル '{LOG_FILE}' を開きました。")
            else:
                logging.warning(f"ログファイル '{LOG_FILE}' が見つかりません。")
        except FileNotFoundError:
            logging.error("notepad.exeが見つかりませんでした。パスが通っているか確認してください。")
        except Exception as e:
            logging.error(f"ログファイルを開く際にエラーが発生しました: {e}", exc_info=True)

    def _clear_log_action(self, icon, item):
        """ログファイルをクリアする"""
        self.window_manager.clear_log()

    def _toggle_pause_action(self, icon, item):
        """一時停止と再開を切り替える"""
        if self.window_manager.toggle_pause():
            self.icon.icon = self.icon_paused
        else:
            self.icon.icon = self.icon_running

    def _create_default_image(self, width, height, color1, color2):
        """デフォルトのアイコン画像を生成する"""
        image = Image.new("RGB", (width, height), color1)
        dc = ImageDraw.Draw(image)
        dc.rectangle((width // 2, 0, width, height // 2), fill=color2)
        dc.rectangle((0, height // 2, width // 2, height), fill=color2)
        return image

    def _reload_settings_action(self, icon, item):
        """設定ファイルを再読み込みする"""
        self.window_manager.reload_settings()

    def _exit_action(self, icon, item):
        logging.info("アプリケーションを終了します。")
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
    # --- パス設定 ---
    # 実行ファイル（.exe）またはスクリプト（.py）の場所を基準にパスを解決する
    if getattr(sys, 'frozen', False):
        # PyInstallerでバンドルされた場合、実行ファイルのディレクトリ
        application_path = os.path.dirname(sys.executable)
    else:
        # スクリプトとして実行された場合、このファイルのディレクトリ
        application_path = os.path.dirname(os.path.abspath(__file__))

    # ユーザーが編集する設定ファイルは、application_path を基準にする
    settings_file_path = os.path.join(application_path, SETTINGS_FILE)
    
    # ログファイルも同じ場所に生成する
    global LOG_FILE
    LOG_FILE = os.path.join(application_path, LOG_FILE)
    
    # --- 処理開始 ---
    setup_logging()
    
    async_worker = None
    win_event_hook = None
    
    try:
        settings = Settings(settings_file_path)
        log_level_str = settings.globals.get("log_level", "INFO")
        log_level = get_log_level_from_string(log_level_str)
        setup_logging(level=log_level)
        logging.info(f"ログレベルを「{log_level_str}」に設定しました。")
        logging.info(f"設定ファイルを '{settings_file_path}' から読み込みました。")

        logging.info("アプリケーションを開始します。")

        # 非同期処理用のワーカースレッドを開始
        async_worker = AsyncWorker()
        async_worker.start()

        window_manager = WindowManager(settings, async_worker.loop)
        
        # Win32イベントフックを開始
        win_event_hook = WinEventHook(window_manager.handle_window_event)
        win_event_hook.start()
        
        # 起動時のウィンドウ処理
        def startup_task():
            time.sleep(1) # イベントフックとの競合を避けるための待機
            window_manager.process_existing_windows()
        
        threading.Thread(target=startup_task, daemon=True).start()

        tray = Tray(window_manager, win_event_hook, application_path)
        tray.run() # これはブロッキング呼び出し

    except Exception as e:
        logging.critical(f"アプリケーションの起動中に致命的なエラーが発生しました: {e}", exc_info=True)
    finally:
        logging.info("アプリケーションのシャットダウン処理を開始します。")
        if win_event_hook and win_event_hook.is_alive():
            win_event_hook.stop()
            win_event_hook.join(timeout=2)
            if win_event_hook.is_alive():
                logging.warning("WinEventHookThread が時間内に終了しませんでした。")
        if async_worker and async_worker.is_alive():
            async_worker.stop()
            async_worker.join(timeout=2)
            if async_worker.is_alive():
                logging.warning("AsyncWorkerThread が時間内に終了しませんでした。")
        logging.info("アプリケーションが終了しました。")

if __name__ == "__main__":
    main()