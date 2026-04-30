# -*- coding: utf-8 -*-
"""
Kindle書籍 自動スクリーンショットツール
========================================
pyautoguiを使ってKindle書籍を自動でスクショするツール。

参考: https://elekibear.com/post/20200225_01

追加機能:
  - マウスで座標を測定するヘルパー
  - PDF変換機能（スクショをPDFにまとめる）
  - 重複検出（ページが変わらなくなったら自動停止）
  - プログレスバー表示
  - 設定ファイル（JSON）の保存/読み込み
  - 画像のリサイズ・最適化
"""

import pyautogui
import time
import os
import sys
import json
import datetime
import hashlib
from pathlib import Path
import tkinter as tk

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import img2pdf
    HAS_IMG2PDF = True
except ImportError:
    HAS_IMG2PDF = False

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False


# ============================================================
# デフォルト設定
# ============================================================
DEFAULT_CONFIG = {
    "page": 0,                    # ページ数 (0 = 自動停止モード)
    "x1": 0,                      # 取得範囲：左上 X
    "y1": 0,                      # 取得範囲：左上 Y
    "x2": 1920,                   # 取得範囲：右下 X
    "y2": 1080,                   # 取得範囲：右下 Y
    "span": 0.8,                  # スクショ間隔 (秒)
    "wait_before_start": 5,       # 開始前の待機時間 (秒)
    "output_folder_prefix": "output",    # 出力フォルダ名の接頭辞
    "output_file_prefix": "page",        # 出力ファイル名の接頭辞
    "image_format": "png",        # 画像形式 (png / jpg)
    "jpg_quality": 90,            # JPG品質 (1-100)
    "resize_width": 0,            # リサイズ横幅 (0 = リサイズなし)
    "auto_stop": True,            # 重複検出で自動停止するか
    "auto_stop_threshold": 3,     # 何回連続で同一画像なら停止するか
    "generate_pdf": True,         # 完了後にPDFを生成するか
    "page_direction": "right",    # ページ送りキー (right / left)
}

CONFIG_FILE = "kindle_config.json"


# ============================================================
# ユーティリティ関数
# ============================================================

def clear_screen():
    """コンソール画面をクリア"""
    os.system("cls" if os.name == "nt" else "clear")


def print_header():
    """ヘッダーを表示"""
    print("=" * 60)
    print("  📖 Kindle 自動スクリーンショットツール")
    print("=" * 60)
    print()


def print_menu():
    """メインメニューを表示"""
    print("  メニューを選択してください:")
    print()
    print("    [1] 📸 スクリーンショットを開始する")
    print("    [2] 🖱️  キャプチャ座標を測定する")
    print("    [3] ⚙️  設定を変更する")
    print("    [4] 📄 既存画像からPDFを生成する")
    print("    [5] 📋 現在の設定を表示する")
    print("    [6] 🚪 終了")
    print()


def load_config() -> dict:
    """設定ファイルを読み込む（なければデフォルト）"""
    config = DEFAULT_CONFIG.copy()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            config.update(saved)
            print(f"  ✅ 設定ファイルを読み込みました: {CONFIG_FILE}")
        except Exception as e:
            print(f"  ⚠️ 設定ファイルの読み込みに失敗: {e}")
    return config


def save_config(config: dict):
    """設定ファイルを保存する"""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        print(f"  ✅ 設定を保存しました: {CONFIG_FILE}")
    except Exception as e:
        print(f"  ⚠️ 設定の保存に失敗: {e}")


def display_config(config: dict):
    """現在の設定を表示"""
    print()
    print("  ─── 現在の設定 ────────────────────────")
    print(f"    ページ数          : {config['page']} {'(自動停止モード)' if config['page'] == 0 else ''}")
    print(f"    取得範囲 左上     : ({config['x1']}, {config['y1']})")
    print(f"    取得範囲 右下     : ({config['x2']}, {config['y2']})")
    print(f"    スクショ間隔      : {config['span']} 秒")
    print(f"    開始前待機時間    : {config['wait_before_start']} 秒")
    print(f"    出力フォルダ接頭辞: {config['output_folder_prefix']}")
    print(f"    出力ファイル接頭辞: {config['output_file_prefix']}")
    print(f"    画像形式          : {config['image_format']}")
    if config['image_format'] == 'jpg':
        print(f"    JPG品質           : {config['jpg_quality']}")
    print(f"    リサイズ横幅      : {config['resize_width'] if config['resize_width'] > 0 else 'なし (原寸)'}")
    print(f"    自動停止          : {'ON' if config['auto_stop'] else 'OFF'}")
    if config['auto_stop']:
        print(f"    自動停止しきい値  : {config['auto_stop_threshold']} 回連続")
    print(f"    PDF生成           : {'ON' if config['generate_pdf'] else 'OFF'}")
    print(f"    ページ送りキー    : {'→ (右)' if config['page_direction'] == 'right' else '← (左)'}")
    print("  ─────────────────────────────────────────")
    print()


def image_hash(img) -> str:
    """画像のハッシュを計算（重複検出用）"""
    return hashlib.md5(img.tobytes()).hexdigest()


# ============================================================
# 座標測定（Snipping Tool風 ドラッグ選択）
# ============================================================

def _run_region_selector() -> tuple:
    """
    Snipping Tool風のドラッグ選択オーバーレイを表示し、
    ユーザーが選択した矩形の座標 (x1, y1, x2, y2) を返す。
    キャンセル時は None を返す。
    """
    result = {"coords": None}

    # --- スクリーンショットを撮影してオーバーレイの背景にする ---
    screenshot = pyautogui.screenshot()

    root = tk.Tk()
    root.title("範囲選択")
    root.attributes("-fullscreen", True)
    root.attributes("-topmost", True)
    root.configure(cursor="crosshair")

    # スクリーンサイズ
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()

    canvas = tk.Canvas(root, width=sw, height=sh,
                       highlightthickness=0, cursor="crosshair")
    canvas.pack(fill=tk.BOTH, expand=True)

    # 背景画像（デスクトップのスクリーンショット）を表示
    from PIL import ImageTk
    bg_image = ImageTk.PhotoImage(screenshot)
    canvas.create_image(0, 0, anchor=tk.NW, image=bg_image)

    # 半透明の暗いオーバーレイ
    overlay = canvas.create_rectangle(0, 0, sw, sh,
                                      fill="black", stipple="gray50",
                                      outline="")

    # 選択矩形と座標表示
    rect_id = None
    start_x = start_y = 0
    label_id = None
    # 明るいくり抜き領域用のリスト
    cutout_id = None

    # --- ガイドテキスト ---
    canvas.create_text(sw // 2, 40,
                       text="📐 ドラッグでキャプチャ範囲を選択  |  Escでキャンセル",
                       fill="white", font=("Yu Gothic UI", 16, "bold"))

    def on_press(event):
        nonlocal start_x, start_y, rect_id, label_id, cutout_id
        start_x, start_y = event.x, event.y
        # 選択矩形（赤い枠線）
        if rect_id:
            canvas.delete(rect_id)
        rect_id = canvas.create_rectangle(
            start_x, start_y, start_x, start_y,
            outline="#FF4444", width=2, dash=(6, 3)
        )
        # くり抜き画像
        if cutout_id:
            canvas.delete(cutout_id)
        # サイズ表示ラベル
        if label_id:
            canvas.delete(label_id)

    def on_drag(event):
        nonlocal label_id, cutout_id
        cur_x, cur_y = event.x, event.y
        canvas.coords(rect_id, start_x, start_y, cur_x, cur_y)

        # サイズ表示を更新
        w = abs(cur_x - start_x)
        h = abs(cur_y - start_y)
        if label_id:
            canvas.delete(label_id)
        label_id = canvas.create_text(
            (start_x + cur_x) // 2, min(start_y, cur_y) - 15,
            text=f"{w} × {h} px",
            fill="white", font=("Yu Gothic UI", 12, "bold")
        )

        # くり抜き: 選択範囲を明るく見せる
        if cutout_id:
            canvas.delete(cutout_id)
        lx = min(start_x, cur_x)
        ly = min(start_y, cur_y)
        rx = max(start_x, cur_x)
        ry = max(start_y, cur_y)
        # 選択範囲に元画像を切り抜いて重ねる
        if rx - lx > 1 and ry - ly > 1:
            cropped = screenshot.crop((lx, ly, rx, ry))
            cutout_photo = ImageTk.PhotoImage(cropped)
            cutout_id = canvas.create_image(lx, ly, anchor=tk.NW,
                                            image=cutout_photo)
            # 参照を保持（ガベージコレクション防止）
            canvas._cutout_photo = cutout_photo
            # 枠線を最前面に
            canvas.tag_raise(rect_id)
            if label_id:
                canvas.tag_raise(label_id)

    def on_release(event):
        end_x, end_y = event.x, event.y
        # 座標を正規化（左上・右下）
        x1 = min(start_x, end_x)
        y1 = min(start_y, end_y)
        x2 = max(start_x, end_x)
        y2 = max(start_y, end_y)
        # 最小サイズチェック
        if (x2 - x1) > 10 and (y2 - y1) > 10:
            result["coords"] = (x1, y1, x2, y2)
        root.destroy()

    def on_escape(event):
        root.destroy()

    canvas.bind("<ButtonPress-1>", on_press)
    canvas.bind("<B1-Motion>", on_drag)
    canvas.bind("<ButtonRelease-1>", on_release)
    root.bind("<Escape>", on_escape)

    root.mainloop()
    return result["coords"]


def measure_coordinates(config: dict):
    """Snipping Tool風のドラッグ選択でキャプチャ座標を測定する"""
    print()
    print("  🖱️  キャプチャ座標の測定（ドラッグ選択）")
    print("  ─────────────────────────────")
    print("  画面全体が暗くなります。")
    print("  マウスをドラッグしてキャプチャ範囲を選択してください。")
    print("  Escキーでキャンセルできます。")
    print()
    input("  Enter を押すと選択画面を開きます...")

    coords = _run_region_selector()

    if coords is None:
        print("  ⏭️ キャンセルされました。")
        print()
        return

    x1, y1, x2, y2 = coords
    print(f"  ✅ 左上座標: ({x1}, {y1})")
    print(f"  ✅ 右下座標: ({x2}, {y2})")
    print(f"  📐 キャプチャ範囲: {x2 - x1} x {y2 - y1} px")
    print()

    apply = input("  この座標を設定に反映しますか？ (y/n): ").strip().lower()
    if apply == "y":
        config["x1"] = x1
        config["y1"] = y1
        config["x2"] = x2
        config["y2"] = y2
        save_config(config)
        print("  ✅ 座標を設定に反映しました。")
    else:
        print("  ⏭️ 反映をスキップしました。")
    print()


# ============================================================
# 設定の変更
# ============================================================

def edit_config(config: dict):
    """対話形式で設定を変更する"""
    print()
    print("  ⚙️  設定の変更")
    print("  ─────────────────────────────")
    print("  変更したい項目の番号を入力してください。")
    print("  (Enter で現在の値を維持します)")
    print()

    items = [
        ("page",                "ページ数 (0=自動停止)",      int),
        ("x1",                  "取得範囲 左上 X",             int),
        ("y1",                  "取得範囲 左上 Y",             int),
        ("x2",                  "取得範囲 右下 X",             int),
        ("y2",                  "取得範囲 右下 Y",             int),
        ("span",                "スクショ間隔 (秒)",           float),
        ("wait_before_start",   "開始前の待機時間 (秒)",       int),
        ("output_folder_prefix","出力フォルダ接頭辞",          str),
        ("output_file_prefix",  "出力ファイル接頭辞",          str),
        ("image_format",        "画像形式 (png/jpg)",          str),
        ("jpg_quality",         "JPG品質 (1-100)",             int),
        ("resize_width",        "リサイズ横幅 (0=原寸)",       int),
        ("auto_stop",           "自動停止 (true/false)",       bool),
        ("auto_stop_threshold", "自動停止しきい値 (回)",       int),
        ("generate_pdf",        "PDF生成 (true/false)",        bool),
        ("page_direction",      "ページ送りキー (right/left)", str),
    ]

    for i, (key, label, typ) in enumerate(items, 1):
        current = config[key]
        val = input(f"  [{i:2d}] {label} (現在: {current}): ").strip()
        if val == "":
            continue
        try:
            if typ == bool:
                config[key] = val.lower() in ("true", "1", "yes", "on")
            else:
                config[key] = typ(val)
        except ValueError:
            print(f"       ⚠️ 無効な値です。スキップします。")

    save_config(config)
    print()
    print("  ✅ 設定を更新しました。")
    display_config(config)


# ============================================================
# スクリーンショット取得
# ============================================================

def take_screenshots(config: dict):
    """Kindleのスクリーンショットを自動取得する"""
    print()
    print("  📸 スクリーンショットを開始します")
    print("  ─────────────────────────────")
    display_config(config)

    page = config["page"]
    x1, y1 = config["x1"], config["y1"]
    x2, y2 = config["x2"], config["y2"]
    span = config["span"]
    wait = config["wait_before_start"]
    fmt = config["image_format"]
    prefix_folder = config["output_folder_prefix"]
    prefix_file = config["output_file_prefix"]
    auto_stop = config["auto_stop"]
    threshold = config["auto_stop_threshold"]
    direction = config["page_direction"]
    resize_w = config["resize_width"]
    jpg_quality = config["jpg_quality"]

    width = x2 - x1
    height = y2 - y1

    if width <= 0 or height <= 0:
        print("  ❌ エラー: キャプチャ範囲が不正です。座標を確認してください。")
        return None

    is_auto_mode = (page == 0)
    if is_auto_mode:
        print("  📌 自動停止モード: ページが変わらなくなったら自動で停止します。")
        if not auto_stop:
            print("  ⚠️  auto_stop が OFF ですが、ページ数=0 のため自動停止を有効にします。")
            auto_stop = True
    else:
        print(f"  📌 {page} ページ分のスクショを取得します。")

    print()
    print(f"  ⏳ {wait} 秒後にスクショを開始します。")
    print("     この間にKindleのウィンドウをアクティブにしてください！")
    print("     （Ctrl+C で中断できます）")
    print()

    for i in range(wait, 0, -1):
        print(f"    開始まで {i} 秒...", end="\r")
        time.sleep(1)
    print("    🚀 スクリーンショット開始！          ")
    print()

    # Kindleウィンドウにフォーカスを当てる（キャプチャ範囲の中央をクリック）
    center_x = x1 + width // 2
    center_y = y1 + height // 2
    pyautogui.click(center_x, center_y)
    time.sleep(0.3)

    # 出力フォルダ作成
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = f"{prefix_folder}_{timestamp}"
    os.makedirs(folder_name, exist_ok=True)

    captured_files = []
    prev_hash = None
    same_count = 0
    max_pages = page if not is_auto_mode else 9999  # 自動モードでは上限を大きく

    # プログレスバーの設定
    if HAS_TQDM and not is_auto_mode:
        pbar = tqdm(total=page, desc="  📸 スクショ中", unit="ページ",
                    bar_format="  {desc}: {bar:30} {n_fmt}/{total_fmt} [{elapsed}<{remaining}]")
    else:
        pbar = None

    try:
        p = 0
        while p < max_pages:
            # スクリーンショット取得
            screenshot = pyautogui.screenshot(region=(x1, y1, width, height))
            current_hash = image_hash(screenshot)

            # ファイル名
            ext = fmt if fmt in ("png", "jpg") else "png"
            out_filename = f"{prefix_file}_{str(p + 1).zfill(4)}.{ext}"
            out_path = os.path.join(folder_name, out_filename)

            # リサイズ処理
            if HAS_PIL and resize_w > 0:
                ratio = resize_w / screenshot.width
                new_height = int(screenshot.height * ratio)
                screenshot = screenshot.resize((resize_w, new_height), Image.LANCZOS)

            # 保存
            if ext == "jpg":
                # JPG保存（RGBに変換が必要）
                if HAS_PIL:
                    screenshot = screenshot.convert("RGB")
                screenshot.save(out_path, quality=jpg_quality)
            else:
                screenshot.save(out_path)

            captured_files.append(out_path)
            prev_hash = current_hash
            p += 1

            # プログレス表示
            if pbar:
                pbar.update(1)
            elif is_auto_mode:
                print(f"    📄 {out_filename} を保存しました (計 {p} ページ)", end="\r")

            # ページ送り (press = keyDown + keyUp で1回押して離す)
            pyautogui.press(direction)

            # ────────────────────────────────────────
            # ページが実際に切り替わるまで待機する
            # (固定待機ではなく画面変化を検出する方式)
            # ────────────────────────────────────────
            POLL_INTERVAL = 0.15       # 画面チェック間隔 (秒)
            TIMEOUT = 5.0              # 最大待機時間 (秒)
            STABLE_WAIT = span         # 変化後の安定待ち (設定のspan値)

            elapsed = 0.0
            page_changed = False

            # フェーズ1: 画面が変わるまで待つ
            while elapsed < TIMEOUT:
                time.sleep(POLL_INTERVAL)
                elapsed += POLL_INTERVAL
                check = pyautogui.screenshot(region=(x1, y1, width, height))
                check_hash = image_hash(check)
                if check_hash != prev_hash:
                    page_changed = True
                    break

            if page_changed:
                # フェーズ2: 画面が安定するまで待つ
                #   (アニメーション中の中間フレームを避ける)
                stable_hash = None
                while elapsed < TIMEOUT:
                    time.sleep(POLL_INTERVAL)
                    elapsed += POLL_INTERVAL
                    check = pyautogui.screenshot(region=(x1, y1, width, height))
                    new_hash = image_hash(check)
                    if new_hash == stable_hash:
                        # 2回連続で同じ = アニメーション完了
                        break
                    stable_hash = new_hash
                # 安定後に少し待って確実にする
                time.sleep(STABLE_WAIT)
                # ページ変化成功 → same_countリセット
                same_count = 0
            else:
                # タイムアウト: ページが変わらなかった (最終ページの可能性)
                same_count += 1
                if auto_stop and same_count >= threshold:
                    print(f"\n  🛑 ページが {threshold} 回連続で変化しませんでした。最終ページと判断し停止します。")
                    break
                time.sleep(STABLE_WAIT)

    except KeyboardInterrupt:
        print("\n\n  ⚠️ ユーザーにより中断されました。")

    finally:
        if pbar:
            pbar.close()

    print()
    print(f"  ✅ 完了！ {len(captured_files)} ページ分のスクショを保存しました。")
    print(f"  📁 出力先: {os.path.abspath(folder_name)}")
    print()

    return folder_name if captured_files else None


# ============================================================
# PDF生成
# ============================================================

def generate_pdf(folder_path: str, output_name: str = None):
    """指定フォルダ内の画像をPDFにまとめる"""
    if not HAS_IMG2PDF:
        print("  ⚠️ img2pdf がインストールされていません。")
        print("     pip install img2pdf を実行してください。")
        return

    if not os.path.isdir(folder_path):
        print(f"  ❌ フォルダが見つかりません: {folder_path}")
        return

    # 画像ファイルを取得（ソート済み）
    valid_ext = {".png", ".jpg", ".jpeg"}
    images = sorted([
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if Path(f).suffix.lower() in valid_ext
    ])

    if not images:
        print(f"  ❌ {folder_path} に画像ファイルが見つかりません。")
        return

    # PNG画像をRGBに変換（img2pdfはRGBA非対応のため）
    converted_images = []
    temp_files = []
    for img_path in images:
        if img_path.lower().endswith(".png") and HAS_PIL:
            img = Image.open(img_path)
            if img.mode == "RGBA":
                rgb_img = img.convert("RGB")
                temp_path = img_path + ".tmp.jpg"
                rgb_img.save(temp_path, quality=95)
                converted_images.append(temp_path)
                temp_files.append(temp_path)
                continue
        converted_images.append(img_path)

    if output_name is None:
        output_name = os.path.basename(folder_path) + ".pdf"

    pdf_path = os.path.join(folder_path, output_name)

    print(f"  📄 PDF生成中... ({len(converted_images)} ページ)")

    try:
        with open(pdf_path, "wb") as f:
            f.write(img2pdf.convert(converted_images))
        print(f"  ✅ PDFを生成しました: {os.path.abspath(pdf_path)}")
    except Exception as e:
        print(f"  ❌ PDF生成に失敗しました: {e}")
    finally:
        # 一時ファイルを削除
        for temp_file in temp_files:
            try:
                os.remove(temp_file)
            except OSError:
                pass
    print()


def pdf_from_existing():
    """既存の画像フォルダからPDFを生成する"""
    print()
    print("  📄 既存画像からPDF生成")
    print("  ─────────────────────────────")

    # outputフォルダの一覧を表示
    output_dirs = sorted([
        d for d in os.listdir(".")
        if os.path.isdir(d) and d.startswith("output")
    ])

    if not output_dirs:
        print("  ❌ output で始まるフォルダが見つかりません。")
        print()
        return

    print("  利用可能なフォルダ:")
    for i, d in enumerate(output_dirs, 1):
        count = len([f for f in os.listdir(d) if Path(f).suffix.lower() in {".png", ".jpg", ".jpeg"}])
        print(f"    [{i}] {d}  ({count} 枚)")
    print()

    choice = input("  フォルダ番号を入力 (またはフォルダパスを直接入力): ").strip()
    try:
        idx = int(choice) - 1
        folder = output_dirs[idx]
    except (ValueError, IndexError):
        folder = choice

    generate_pdf(folder)


# ============================================================
# メインループ
# ============================================================

def main():
    # pyautogui の安全設定
    pyautogui.FAILSAFE = True  # 画面左上にカーソルを移動で緊急停止
    pyautogui.PAUSE = 0.1

    config = load_config()

    while True:
        clear_screen()
        print_header()
        print_menu()

        choice = input("  > ").strip()

        if choice == "1":
            folder = take_screenshots(config)
            if folder and config.get("generate_pdf", False):
                ans = input("  📄 PDFを生成しますか？ (y/n): ").strip().lower()
                if ans == "y":
                    generate_pdf(folder)
            input("  Enter で メニューに戻ります...")

        elif choice == "2":
            measure_coordinates(config)
            input("  Enter で メニューに戻ります...")

        elif choice == "3":
            edit_config(config)
            input("  Enter で メニューに戻ります...")

        elif choice == "4":
            pdf_from_existing()
            input("  Enter で メニューに戻ります...")

        elif choice == "5":
            display_config(config)
            input("  Enter で メニューに戻ります...")

        elif choice == "6":
            print("  👋 終了します。")
            break

        else:
            print("  ⚠️ 無効な選択です。1-6 の数字を入力してください。")
            time.sleep(1)


if __name__ == "__main__":
    main()
