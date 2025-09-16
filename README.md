# Window Mover

Window Moverは、ユーザーが定義したルールに基づいて、Windows上のアプリケーションウィンドウを自動的に移動・リサイズするためのユーティリティです。

`settings.toml`という単一のファイルにルールを記述するだけで、ウィンドウの配置を柔軟に自動化できます。

![image](https://user-images.githubusercontent.com/1234567/123456789-abcdef.png)  <!-- この行は後で実際のスクリーンショットに置き換えてください -->

## 主な機能

- **ルールベースのウィンドウ管理**: ウィンドウのタイトル、プロセス名、クラス名に基づいて、移動やリサイズを自動化します。
- **柔軟な設定**: TOML形式の設定ファイルにより、直感的かつ柔軟にルールを定義できます。
- **マルチモニター対応**: 特定のモニターへのウィンドウ移動をサポートします。
- **仮想デスクトップ対応**: 指定した仮想デスクトップへウィンドウを移動させることができます。
- **多彩な配置オプション**:
    - 9点のアンカーポイント（ウィンドウの基点）を指定可能。
    - `px`（ピクセル）または `%`（パーセンテージ）でのサイズ・位置指定。
    - 最終的な位置からの微調整（オフセット）も可能です。
- **システムトレイ常駐**: タスクトレイのアイコンから「一時停止/再開」「設定の再読み込み」「終了」を簡単に行えます。

## 動作環境

- Windows
- Python 3.8以上

## インストールと実行

1.  **リポジトリをクローンまたはダウンロードします。**

2.  **必要なライブラリをインストールします。**
    ```sh
    pip install -r requirements.txt
    ```

3.  **アプリケーションを実行します。**
    ```sh
    python main.py
    ```
    実行後、アプリケーションはシステムトレイに常駐します。

## 設定方法 (`settings.toml`)

このアプリケーションのすべての動作は `settings.toml` ファイルで制御します。ファイルは大きく分けて `[global]` `[[ignores]]` `[[rules]]` の3つのセクションで構成されます。

---

### 1. `[global]` - 全体設定

アプリケーション全体の動作を定義します。

```toml
[global]
# ログレベル: "DEBUG", "INFO", "WARNING", "ERROR"
log_level = "INFO"

# 起動時に既存ウィンドウへルールを適用するか
apply_on_startup = true
# 設定再読み込み時に既存ウィンドウへルールを適用するか
apply_on_reload = true
# 一時停止から再開する時に既存ウィンドウへルールを適用するか
apply_on_resume = false

# 無効なウィンドウハンドルを追跡リストから削除する間隔（秒）
cleanup_interval_seconds = 300

# モニターごとのオフセット設定
[global.monitor_offsets]
  # デフォルトのオフセット
  default = { top = 0, bottom = 0, left = 0, right = 0 }
  # モニター1のオフセット（例: タスクバーが下にある場合）
  monitor_1 = { top = 0, bottom = 48, left = 0, right = 0 }
```

- `monitor_offsets`: モニターの作業領域から、さらに除外したい領域（タスクバーなど）をピクセル単位で指定します。`move_to`でアンカーポイント（`"MiddleCenter"`など）を指定した際に、このオフセットを考慮した上で中央配置などが行われます。

---

### 2. `[[ignores]]` - 無視ルール

ここに記述された条件に一致するウィンドウは、後続のどのルールにもマッチしなくなり、処理対象外となります。OSのUIコンポーネントなど、動かしたくないウィンドウを指定するのに便利です。

```toml
[[ignores]]
  name = "OSコンポーネントを無視"
  logic = "OR" # "OR"または"AND"
  conditions = [
    { process = "SystemSettings.exe" },      # Windows 設定
    { class = "Shell_TrayWnd" },             # タスクバー
    { title = "regex:^検索$" }                # 正規表現を使ったタイトル指定
  ]
```

- `logic`: `conditions`内の複数の条件をどのように評価するかを指定します。
    - `"OR"`: いずれかの条件に一致すれば、無視リストにマッチします。
    - `"AND"`: すべての条件に一致した場合のみ、無視リストにマッチします。
- `conditions`:
    - `title`: ウィンドウのタイトル（部分一致、`regex:`で正規表現も可）
    - `process`: プロセス名（完全一致、`regex:`で正規表現も可）
    - `class`: ウィンドウクラス名（部分一致、`regex:`で正規表現も可）
    - `case_sensitive`: `true`にすると大文字・小文字を区別します（デフォルトは`false`）

---

### 3. `[[rules]]` - 個別ルール

ウィンドウをどのように動かすかを定義するメインのセクションです。ルールは上から順に評価され、最初に条件が一致したルールが1度だけ適用されます。

#### ルールの構成

各ルールは `name` `condition` `action` の3つのパートで構成されます。

```toml
[[rules]]
  name = "（ルールの名前）"

  [rules.condition]
  # ... 条件 ...

  [rules.action]
  # ... 動作 ...
```

#### `[rules.condition]` - 条件

無視ルールと同様の条件を記述します。

```toml
# 単一条件
[rules.condition]
title = "電卓"

# 複数条件 (AND/OR)
[rules.condition]
logic = "OR"
conditions = [
  { title = "regex:^タイトルなし - メモ帳$" },
  { title = "regex:^無題 - メモ帳$" }
]
```

#### ヒント: ルールの条件指定

- **ゲームやUDPアプリケーション、ランチャー経由で起動するアプリ**などは、プロセス名が期待通りでない場合があります。このようなアプリケーションを対象にする場合は、`process`や`class`での指定がうまく機能しないことがあります。
- その代わり、**`title`（ウィンドウタイトル）での指定を推奨します。** ウィンドウタイトルは最も安定して取得できる識別子であることが多いため、正規表現 (`regex:`) を活用することで、より確実にウィンドウを捉えることができます。

#### `[rules.action]` - 動作

ウィンドウをどのように操作するかを定義します。

```toml
[rules.action]
# ウィンドウのどの点を基準にするか (9点から選択)
anchor = "MiddleCenter"

# どこに移動させるか
move_to = "MiddleCenter"
# または、絶対/相対座標で指定
# move_to = { x = "10%", y = "10px" }

# どのサイズにするか
resize_to = { width = 320, height = "50%" }
# w, hでの省略形も可
# resize_to = { w = 320, h = 480 }

# 最終位置からの微調整
offset = { x = -20, y = 0 }

# 移動先のモニター番号 (1から始まる)
target_monitor = 2

# 移動先の仮想デスクトップ番号 (1から始まる)
target_workspace = 2

# 最大化/最小化
maximize = "ON" # または "OFF"
minimize = "ON" # または "OFF"

# 処理実行までの遅延時間 (ミリ秒)
execution_delay = 1000
```

- `anchor` / `move_to`で指定可能なアンカーポイント:
  - `TopLeft`, `TopCenter`, `TopRight`
  - `MiddleLeft`, `MiddleCenter`, `MiddleRight`
  - `BottomLeft`, `BottomCenter`, `BottomRight`

- `move_to`を座標（`{x=..., y=...}`）で指定した場合、`global.monitor_offsets`は適用されません。モニターの左上隅が`(0, 0)`となります。

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細は `LICENSE` ファイルをご覧ください。

## アイコン

- `window_mover.ico`
- `window_mover_pause.ico`

これらのアイコンは [（ここにアイコンの出所や作者を記述）] によって作成されました。
