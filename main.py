import json
import os
import psutil
import win32gui
import win32process
import win32con
import logging

CONFIG_PATH = "config.json"

# -------------------------
# ログ設定
# -------------------------
logging.basicConfig(
    filename="app.log",
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def log(msg, debug=False):
    print(msg)
    if debug:
        logging.debug(msg)
    else:
        logging.info(msg)

# -------------------------
# 設定読み込み
# -------------------------
def load_config(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

# -------------------------
# パス一致するプロセスを探す
# -------------------------
def find_processes_by_path(target_path, debug=False):
    target_path = os.path.normcase(target_path)
    matches = []

    for p in psutil.process_iter(['pid', 'exe']):
        try:
            exe = p.info['exe']
            if exe and os.path.normcase(exe) == target_path:
                matches.append(p)
        except Exception as e:
            log(f"  → プロセス情報取得エラー: {e}", debug)

    return matches

# -------------------------
# PID → ウィンドウ最小化
# -------------------------
def minimize_windows_by_pid(pid, debug=False):
    found = False

    def callback(hwnd, _):
        nonlocal found
        try:
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
            if window_pid == pid:
                found = True

                # 状態チェック
                placement = win32gui.GetWindowPlacement(hwnd)
                show_cmd = placement[1]

                if show_cmd == win32con.SW_SHOWMINIMIZED:
                    log(f"    → すでに最小化されています: {hex(hwnd)}", debug)
                else:
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                    log(f"    → 最小化しました: {hex(hwnd)}", debug)

        except Exception as e:
            log(f"    → ウィンドウ処理エラー: {e}", debug)

    win32gui.EnumWindows(callback, None)

    if not found:
        log("    → ウィンドウが見つかりません", debug)

# -------------------------
# メイン処理
# -------------------------
def main():
    config = load_config(CONFIG_PATH)
    debug = config.get("debug", False)
    targets = config.get("targets", [])

    log("=== 実行開始 ===", debug)

    for exe_path in targets:
        log(f"Checking: {exe_path}", debug)

        processes = find_processes_by_path(exe_path, debug)
        if not processes:
            log("  → プロセスが見つかりません", debug)
            continue

        for proc in processes:
            log(f"  → PID {proc.pid} を処理中", debug)
            minimize_windows_by_pid(proc.pid, debug)

    log("=== 完了 ===", debug)

if __name__ == "__main__":
    main()
