import base64
import ctypes
import json
import os
import time
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests
import win32con
import win32gui
import win32process
import win32ui
from PIL import Image, ImageGrab


CONFIG_PATH = Path("config.json")


@dataclass
class AppConfig:
    napcat_base_url: str
    napcat_access_token: str
    napcat_send_max_retries: int
    target_qq: int
    capture_fullscreen: bool
    capture_all_screens: bool
    window_title: str
    window_hwnd: Optional[int]
    window_pid: Optional[int]
    window_class: Optional[str]
    interval_minutes: float
    openai_base_url: str
    openai_api_key: str
    model: str
    prompt: str


class NapCatApiError(RuntimeError):
    def __init__(self, action: str, raw: Dict[str, Any]):
        self.action = action
        self.raw = raw
        self.retcode = raw.get("retcode")
        self.status = raw.get("status")
        self.msg = raw.get("msg")
        self.message = raw.get("message")
        self.wording = raw.get("wording")
        detail = self.wording or self.message or self.msg or str(raw)
        super().__init__(
            f"NapCat API {action} failed: retcode={self.retcode}, status={self.status}, detail={detail}"
        )


def runtime_tag() -> str:
    try:
        mtime = int(Path(__file__).stat().st_mtime)
    except Exception:
        mtime = 0
    return f"pid={os.getpid()} mtime={mtime}"


def _to_bool(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "y", "on"}
    return bool(value)


def load_config_raw(path: Path = CONFIG_PATH) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


def save_config_raw(raw: Dict[str, Any], path: Path = CONFIG_PATH) -> None:
    with path.open("w", encoding="utf-8") as file:
        json.dump(raw, file, ensure_ascii=False, indent=2)


def load_config(path: Path = CONFIG_PATH) -> AppConfig:
    raw = load_config_raw(path)

    required_keys = [
        "napcat_base_url",
        "target_qq",
        "interval_minutes",
        "openai_base_url",
        "openai_api_key",
        "model",
    ]
    missing = [key for key in required_keys if key not in raw]
    if missing:
        raise ValueError(f"config.json missing keys: {', '.join(missing)}")

    interval = float(raw["interval_minutes"])
    if interval <= 0:
        raise ValueError("interval_minutes must be > 0")

    capture_fullscreen = _to_bool(raw.get("capture_fullscreen", False), default=False)
    capture_all_screens = _to_bool(raw.get("capture_all_screens", True), default=True)
    window_title = str(raw.get("window_title", ""))
    window_hwnd = int(raw["window_hwnd"]) if raw.get("window_hwnd") not in (None, "") else None
    window_pid = int(raw["window_pid"]) if raw.get("window_pid") not in (None, "") else None
    window_class = str(raw["window_class"]) if raw.get("window_class") else None

    if not capture_fullscreen and not window_title and window_hwnd is None:
        raise ValueError("window_title/window_hwnd is required when capture_fullscreen=false")

    return AppConfig(
        napcat_base_url=str(raw["napcat_base_url"]).rstrip("/"),
        napcat_access_token=str(raw.get("napcat_access_token", "")).strip(),
        napcat_send_max_retries=max(1, int(raw.get("napcat_send_max_retries", 5))),
        target_qq=int(raw["target_qq"]),
        capture_fullscreen=capture_fullscreen,
        capture_all_screens=capture_all_screens,
        window_title=window_title,
        window_hwnd=window_hwnd,
        window_pid=window_pid,
        window_class=window_class,
        interval_minutes=interval,
        openai_base_url=str(raw["openai_base_url"]),
        openai_api_key=str(raw["openai_api_key"]),
        model=str(raw["model"]),
        prompt=str(
            raw.get(
                "prompt",
                "Please describe the screenshot in detail, extract visible text, and summarize key points.",
            )
        ),
    )


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
            }
        )

    win32gui.EnumWindows(enum_handler, windows)
    return windows


def _pick_single(candidates: List[Dict[str, Any]], reason: str) -> Tuple[Optional[int], Optional[str]]:
    if len(candidates) == 1:
        return int(candidates[0]["hwnd"]), None
    if len(candidates) > 1:
        details = "; ".join(
            f"pid={item['pid']} hwnd={item['hwnd']} class={item['class_name']} title={item['title']}"
            for item in candidates[:5]
        )
        return None, f"Ambiguous window ({reason}), {len(candidates)} matches: {details}"
    return None, None


def resolve_window_handle(config: AppConfig) -> Tuple[Optional[int], Optional[str]]:
    windows = list_visible_windows()
    if not windows:
        return None, "No visible windows found."

    title = config.window_title.strip()

    if config.window_hwnd is not None:
        direct = [item for item in windows if int(item["hwnd"]) == int(config.window_hwnd)]
        if direct:
            one = direct[0]
            if config.window_pid is not None and int(one["pid"]) != int(config.window_pid):
                return None, "Configured hwnd found but pid mismatched. Re-run config.py."
            if config.window_class and str(one["class_name"]) != str(config.window_class):
                return None, "Configured hwnd found but class mismatched. Re-run config.py."
            return int(one["hwnd"]), None

    strict = windows
    if config.window_pid is not None:
        strict = [item for item in strict if int(item["pid"]) == int(config.window_pid)]
        if config.window_class:
            strict = [item for item in strict if str(item["class_name"]) == str(config.window_class)]

    if title:
        exact = [item for item in strict if str(item["title"]) == title]
        hwnd, err = _pick_single(exact, "exact title + configured pid/class")
        if hwnd is not None:
            return hwnd, None
        if err is not None:
            if config.window_pid is None:
                return None, f"{err}. Run run_config.bat and select the exact target window."
            return None, err

        fuzzy = [item for item in strict if title.lower() in str(item["title"]).lower()]
        hwnd, err = _pick_single(fuzzy, "fuzzy title + configured pid/class")
        if hwnd is not None:
            return hwnd, None
        if err is not None:
            if config.window_pid is None:
                return None, f"{err}. Run run_config.bat and select the exact target window."
            return None, err

    if config.window_pid is not None or config.window_class:
        return None, "Configured target window not found. Please run config.py to reselect."

    exact_legacy = [item for item in windows if str(item["title"]) == title]
    hwnd, err = _pick_single(exact_legacy, "exact title")
    if hwnd is not None:
        return hwnd, None
    if err is not None:
        return None, f"{err}. Run run_config.bat and select the exact target window."

    fuzzy_legacy = [item for item in windows if title.lower() in str(item["title"]).lower()]
    hwnd, err = _pick_single(fuzzy_legacy, "fuzzy title")
    if hwnd is not None:
        return hwnd, None
    if err is not None:
        return None, f"{err}. Run run_config.bat and select the exact target window."

    return None, f"Window not found: {title}"


def capture_window(config: AppConfig) -> Tuple[Optional[Image.Image], Optional[str]]:
    hwnd, find_error = resolve_window_handle(config)
    if not hwnd:
        return None, find_error or f"Window not found: {config.window_title}"

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top
    if width <= 0 or height <= 0:
        return None, f"Invalid window size: {width}x{height}"

    # Preferred: screen grab by actual window bounds (more complete on cmd/legacy windows).
    # Fallback to PrintWindow/DC copy when screen grab is unavailable.
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass

    time.sleep(0.2)
    try:
        try:
            screen_image = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
        except TypeError:
            screen_image = ImageGrab.grab(bbox=(left, top, right, bottom))
        if screen_image is not None:
            return screen_image.convert("RGB"), None
    except Exception:
        pass

    hwnd_dc = mfc_dc = save_dc = bitmap = None
    try:
        hwnd_dc = win32gui.GetWindowDC(hwnd)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
        save_dc.SelectObject(bitmap)

        success = 0
        if hasattr(win32gui, "PrintWindow"):
            success = win32gui.PrintWindow(hwnd, save_dc.GetSafeHdc(), 0)
        else:
            try:
                success = ctypes.windll.user32.PrintWindow(int(hwnd), int(save_dc.GetSafeHdc()), 0)
            except Exception:
                success = 0

        if int(success) != 1:
            save_dc.BitBlt((0, 0), (width, height), mfc_dc, (0, 0), win32con.SRCCOPY)

        bmp_info = bitmap.GetInfo()
        bmp_bytes = bitmap.GetBitmapBits(True)
        image = Image.frombuffer(
            "RGB",
            (bmp_info["bmWidth"], bmp_info["bmHeight"]),
            bmp_bytes,
            "raw",
            "BGRX",
            0,
            1,
        )
        return image, None
    except Exception as exc:
        return None, f"Capture failed: {exc}"
    finally:
        if bitmap is not None:
            win32gui.DeleteObject(bitmap.GetHandle())
        if save_dc is not None:
            save_dc.DeleteDC()
        if mfc_dc is not None:
            mfc_dc.DeleteDC()
        if hwnd_dc is not None:
            win32gui.ReleaseDC(hwnd, hwnd_dc)


def capture_fullscreen(config: AppConfig) -> Tuple[Optional[Image.Image], Optional[str]]:
    try:
        try:
            image = ImageGrab.grab(all_screens=config.capture_all_screens)
        except TypeError:
            image = ImageGrab.grab()
        if image is None:
            return None, "Full screen capture returned empty image"
        return image.convert("RGB"), None
    except Exception as exc:
        return None, f"Full screen capture failed: {exc}"


def capture_image(config: AppConfig) -> Tuple[Optional[Image.Image], Optional[str]]:
    if config.capture_fullscreen:
        return capture_fullscreen(config)
    return capture_window(config)


def image_to_base64(image: Image.Image) -> str:
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def _content_to_text(content: Any) -> str:
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts: List[str] = []
        for item in content:
            if isinstance(item, str):
                parts.append(item)
            elif isinstance(item, dict):
                if item.get("type") == "text" and item.get("text"):
                    parts.append(str(item["text"]))
                elif item.get("text"):
                    parts.append(str(item["text"]))
        return "\n".join(part for part in parts if part).strip()
    if isinstance(content, dict):
        if "text" in content and content["text"] is not None:
            return str(content["text"])
        return json.dumps(content, ensure_ascii=False)
    if content is None:
        return ""
    return str(content)


def _extract_text_from_sse_string(raw_text: str) -> str:
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    if not any(line.startswith("data:") for line in lines):
        return raw_text

    chunks: List[str] = []
    upstream_error: Optional[str] = None

    for line in lines:
        if not line.startswith("data:"):
            continue
        payload = line[5:].strip()
        if not payload or payload == "[DONE]":
            continue

        try:
            item = json.loads(payload)
        except Exception:
            chunks.append(payload)
            continue

        if isinstance(item, dict) and isinstance(item.get("error"), dict):
            err = item["error"]
            upstream_error = (
                f"{err.get('message') or err.get('type') or 'unknown error'}"
                f" (code={err.get('code')})"
            )
            continue

        if isinstance(item, dict):
            choices = item.get("choices")
            if isinstance(choices, list):
                for choice in choices:
                    if not isinstance(choice, dict):
                        continue
                    delta = choice.get("delta")
                    if isinstance(delta, dict):
                        delta_text = _content_to_text(delta.get("content"))
                        if delta_text:
                            chunks.append(delta_text)
                    message = choice.get("message")
                    if isinstance(message, dict):
                        msg_text = _content_to_text(message.get("content"))
                        if msg_text:
                            chunks.append(msg_text)
                    text = _content_to_text(choice.get("text"))
                    if text:
                        chunks.append(text)

    text_result = "".join(chunks).strip()
    if text_result:
        return text_result
    if upstream_error:
        raise RuntimeError(f"Upstream model error: {upstream_error}")
    raise RuntimeError("Streaming response contains no text content")


def _extract_text_from_response(response: Any) -> str:
    if isinstance(response, str):
        stripped = response.strip()
        if stripped.startswith("data:") or "\ndata:" in stripped:
            return _extract_text_from_sse_string(stripped)
        try:
            parsed = json.loads(stripped)
            if isinstance(parsed, dict) and isinstance(parsed.get("error"), dict):
                err = parsed["error"]
                raise RuntimeError(
                    f"Upstream model error: {err.get('message') or err.get('type')} "
                    f"(code={err.get('code')})"
                )
        except RuntimeError:
            raise
        except Exception:
            pass
        return response

    if isinstance(response, dict):
        choices = response.get("choices")
        if isinstance(choices, list) and choices:
            first = choices[0]
            if isinstance(first, dict):
                message = first.get("message")
                if isinstance(message, dict):
                    text = _content_to_text(message.get("content"))
                    if text:
                        return text
                text = _content_to_text(first.get("text"))
                if text:
                    return text
        for key in ("output_text", "text", "content"):
            if key in response:
                text = _content_to_text(response.get(key))
                if text:
                    return text
        raise ValueError(f"Unsupported dict response format, keys={list(response.keys())[:10]}")

    if hasattr(response, "choices"):
        choices = getattr(response, "choices")
        if choices:
            first = choices[0]
            message = getattr(first, "message", None)
            if message is not None:
                text = _content_to_text(getattr(message, "content", None))
                if text:
                    return text
            text = _content_to_text(getattr(first, "text", None))
            if text:
                return text
        raise ValueError("Response has choices but no readable text content")

    if hasattr(response, "model_dump"):
        return _extract_text_from_response(response.model_dump())

    return _content_to_text(response)


def _build_napcat_headers(config: AppConfig) -> Dict[str, str]:
    headers: Dict[str, str] = {}
    if config.napcat_access_token:
        headers["Authorization"] = f"Bearer {config.napcat_access_token}"
    return headers


def call_napcat_api(config: AppConfig, action: str, payload: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    api_url = f"{config.napcat_base_url}/{action.lstrip('/')}"
    body = payload or {}
    headers = _build_napcat_headers(config)

    response = requests.post(api_url, json=body, headers=headers, timeout=15)
    response.raise_for_status()

    try:
        data = response.json()
    except Exception as exc:
        text_preview = response.text[:500]
        raise RuntimeError(
            f"NapCat returned non-JSON response for {action}: "
            f"http={response.status_code}, body={text_preview}"
        ) from exc

    if isinstance(data, dict):
        retcode = data.get("retcode")
        if retcode not in (0, "0"):
            raise NapCatApiError(action, data)
        return data

    raise RuntimeError(f"NapCat API {action} returned invalid JSON type: {type(data).__name__}")


def _is_napcat_timeout_error(exc: Exception) -> bool:
    if isinstance(exc, NapCatApiError):
        if str(exc.retcode) == "200":
            return True
        detail = f"{exc.wording or ''} {exc.message or ''} {exc.msg or ''}".lower()
        return "timeout" in detail
    text = str(exc).lower()
    return "timeout" in text


def _call_napcat_with_retry(
    config: AppConfig,
    action: str,
    payload: Dict[str, Any],
    retries: Optional[int] = None,
) -> Dict[str, Any]:
    max_retries = retries or config.napcat_send_max_retries
    last_error: Optional[Exception] = None

    for attempt in range(1, max_retries + 1):
        try:
            return call_napcat_api(config, action, payload)
        except Exception as exc:
            last_error = exc
            if attempt >= max_retries:
                break
            if not _is_napcat_timeout_error(exc):
                break
            wait_seconds = min(5, attempt)
            print(
                f"NapCat {action} timeout (attempt {attempt}/{max_retries}): {exc}. "
                f"Retrying in {wait_seconds}s..."
            )
            time.sleep(wait_seconds)

    raise RuntimeError(f"NapCat {action} failed after {max_retries} attempts: {last_error}")


def _send_private_with_fallback(config: AppConfig, payload: Dict[str, Any]) -> Dict[str, Any]:
    try:
        return _call_napcat_with_retry(config, "send_private_msg", payload)
    except Exception as first_error:
        print(f"send_private_msg failed: {first_error}. Trying send_msg fallback...")
        fallback_payload = {
            "message_type": "private",
            "user_id": payload["user_id"],
            "message": payload["message"],
        }
        return _call_napcat_with_retry(config, "send_msg", fallback_payload)


def check_napcat_connection(config: AppConfig) -> None:
    try:
        result = call_napcat_api(config, "get_login_info")
        user_id = None
        nickname = None
        if isinstance(result.get("data"), dict):
            user_id = result["data"].get("user_id")
            nickname = result["data"].get("nickname")
        print(f"NapCat connection OK. login_user_id={user_id}, nickname={nickname}")
        if user_id is not None and int(user_id) == int(config.target_qq):
            print("Warning: target_qq equals current login QQ. Some clients may not deliver self private messages.")
    except Exception as exc:
        print(f"NapCat connection check failed: {exc}")
        return

    try:
        friends_result = call_napcat_api(config, "get_friend_list")
        friends = friends_result.get("data")
        if isinstance(friends, list):
            friend_ids = {
                int(item.get("user_id"))
                for item in friends
                if isinstance(item, dict) and item.get("user_id") is not None
            }
            if int(config.target_qq) not in friend_ids and (user_id is None or int(config.target_qq) != int(user_id)):
                print(
                    f"Warning: target_qq={config.target_qq} not found in friend list. "
                    "send_private_msg may fail."
                )
    except Exception as exc:
        print(f"Friend list check skipped: {exc}")


def analyze_with_ai(config: AppConfig, image: Image.Image) -> str:
    try:
        from openai import OpenAI
    except ImportError as exc:
        raise RuntimeError("openai package not installed. Run: pip install -r requirements.txt") from exc

    client = OpenAI(base_url=config.openai_base_url, api_key=config.openai_api_key)
    image_b64 = image_to_base64(image)
    max_retries = 5
    last_error: Optional[Exception] = None
    last_response_type = "None"

    for attempt in range(1, max_retries + 1):
        try:
            response = client.chat.completions.create(
                model=config.model,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": config.prompt},
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:image/png;base64,{image_b64}"},
                            },
                        ],
                    }
                ],
                temperature=0.7,
                max_tokens=1500,
                stream=False,
            )
            last_response_type = type(response).__name__
            text = _extract_text_from_response(response)
            if text:
                return text
            raise ValueError("Model returned empty text content")
        except Exception as exc:
            last_error = exc
            if attempt >= max_retries:
                break
            wait_seconds = min(8, attempt * 2)
            print(
                f"Model request failed (attempt {attempt}/{max_retries}): {exc}. "
                f"Retrying in {wait_seconds}s..."
            )
            time.sleep(wait_seconds)

    raise RuntimeError(
        f"Model request/parse failed after {max_retries} attempts "
        f"(last_response_type={last_response_type}): {last_error}"
    )


def send_to_qq(config: AppConfig, text: str, image: Optional[Image.Image] = None) -> None:
    # Send text and image separately to reduce timeout risk on large mixed payloads.
    text_message = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Screenshot result:\n{text}"
    text_payload = {"user_id": config.target_qq, "message": text_message}
    result = _send_private_with_fallback(config, text_payload)
    message_id = None
    if isinstance(result.get("data"), dict):
        message_id = result["data"].get("message_id")
    print(f"NapCat text message sent. message_id={message_id}")

    # Intentionally keep text-only delivery to avoid private message image timeout issues.
    _ = image


def run_once(config: AppConfig) -> None:
    if config.capture_fullscreen:
        target_hint = f"FULLSCREEN (all_screens={config.capture_all_screens})"
    else:
        target_hint = f"title='{config.window_title}'"
        if config.window_pid is not None:
            target_hint += f", pid={config.window_pid}"
        if config.window_hwnd is not None:
            target_hint += f", hwnd={config.window_hwnd}"
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Capturing window: {target_hint}")

    image, capture_error = capture_image(config)
    if image is None:
        error_message = capture_error or "Unknown capture error"
        print(error_message)
        try:
            send_to_qq(config, f"Capture failed: {error_message}")
        except Exception as exc:
            print(f"Failed to send capture error to QQ: {exc}")
        return

    try:
        print("Calling vision model...")
        analysis = analyze_with_ai(config, image)
    except Exception as exc:
        analysis = f"AI analysis failed: {exc}"
        print(analysis)

    try:
        print("Sending message to QQ...")
        send_to_qq(config, analysis, None)
        print("Message sent.")
    except Exception as exc:
        print(f"Failed to send message: {exc}")


def run_scheduler() -> None:
    print("NapCat screenshot + AI recognizer started.")
    print("Config is reloaded before every run; updates apply on the next cycle.")
    print(f"Runtime: {runtime_tag()} | script={Path(__file__).resolve()}")

    while True:
        try:
            cfg = load_config()
        except Exception as exc:
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Config error: {exc}")
            time.sleep(10)
            continue

        check_napcat_connection(cfg)
        run_once(cfg)

        sleep_seconds = max(1, int(cfg.interval_minutes * 60))
        print(f"Sleeping {sleep_seconds} seconds...")
        for _ in range(sleep_seconds):
            time.sleep(1)


def main() -> int:
    run_scheduler()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
