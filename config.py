import json
from pathlib import Path
from shutil import copyfile
from typing import Any, Dict, List, Tuple

import win32gui
import win32process


CONFIG_PATH = Path("config.json")
CONFIG_EXAMPLE_PATH = Path("config.json.example")


def ensure_config() -> bool:
    if CONFIG_PATH.exists():
        return True
    if CONFIG_EXAMPLE_PATH.exists():
        copyfile(CONFIG_EXAMPLE_PATH, CONFIG_PATH)
        print("Auto-created config.json from config.json.example")
        return True
    print("Missing both config.json and config.json.example")
    return False


def load_config_raw() -> Dict[str, Any]:
    with CONFIG_PATH.open("r", encoding="utf-8") as file:
        return json.load(file)


def save_config_raw(raw: Dict[str, Any]) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as file:
        json.dump(raw, file, ensure_ascii=False, indent=2)


def list_visible_windows() -> List[Dict[str, Any]]:
    windows: List[Dict[str, Any]] = []

    def enum_handler(hwnd: int, result: List[Dict[str, Any]]) -> None:
        if not win32gui.IsWindowVisible(hwnd):
            return

        title = win32gui.GetWindowText(hwnd).strip()
        if not title:
            return

        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top
        if width <= 0 or height <= 0:
            return

        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        result.append(
            {
                "hwnd": hwnd,
                "title": title,
                "class_name": win32gui.GetClassName(hwnd),
                "pid": pid,
                "rect": (left, top, right, bottom),
                "size": f"{width}x{height}",
            }
        )

    win32gui.EnumWindows(enum_handler, windows)
    return windows


def print_windows(windows: List[Dict[str, Any]]) -> None:
    print("Visible windows:")
    for idx, item in enumerate(windows, start=1):
        left, top, right, bottom = item["rect"]
        print(
            f"[{idx:03d}] {item['title']} | class={item['class_name']} | "
            f"pid={item['pid']} | hwnd={item['hwnd']} | pos=({left},{top}) | "
            f"size={item['size']}"
        )


def activate_window(hwnd: int) -> Tuple[bool, str]:
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, 9)
        else:
            win32gui.ShowWindow(hwnd, 5)
        win32gui.SetForegroundWindow(hwnd)
        return True, "Window activated. Please confirm visually."
    except Exception as exc:
        return False, f"Window activation failed: {exc}"


def main() -> int:
    if not ensure_config():
        return 1

    windows = list_visible_windows()
    if not windows:
        print("No visible windows found. Open the target window and try again.")
        return 1

    print_windows(windows)
    while True:
        text = input("\nInput window index (Enter to cancel): ").strip()
        if not text:
            print("Cancelled.")
            return 1
        if not text.isdigit():
            print("Invalid input. Please enter a number.")
            continue

        index = int(text)
        if index < 1 or index > len(windows):
            print("Index out of range.")
            continue

        selected = windows[index - 1]
        ok, msg = activate_window(int(selected["hwnd"]))
        print(msg)
        if not ok:
            continue

        confirm = input("Use this window? (y/n): ").strip().lower()
        if confirm != "y":
            print("Not confirmed. Pick again.")
            continue

        raw = load_config_raw()
        raw["window_title"] = selected["title"]
        raw["window_hwnd"] = int(selected["hwnd"])
        raw["window_pid"] = int(selected["pid"])
        raw["window_class"] = selected["class_name"]
        save_config_raw(raw)

        print("\nSaved to config.json:")
        print(f"window_title = {selected['title']}")
        print(f"window_hwnd  = {selected['hwnd']}")
        print(f"window_pid   = {selected['pid']}")
        print(f"window_class = {selected['class_name']}")
        print("Now run napcat_screenshot_ai.py (or run_venv.bat).")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
