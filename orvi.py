"""
ORVI Simulator - шуточная программа, имитирующая симптомы простуды на компьютере.

После паузы 3-7 секунд случайно выполняется одно из трёх "действий" (по ~33.3% каждое):
  - чих: открывается случайное приложение (браузер, диспетчер задач, экранная клавиатура, проводник)
  - кашель: все видимые окна на экране на секунду начинают "трястись"
  - сморкание: системная громкость случайно немного повышается или понижается

Дополнительно, независимо от основного действия, есть небольшой шанс "поднятия температуры" -
экран на несколько секунд слегка подсвечивается красным.

При первом запуске программа регистрирует видимое задание в Планировщике заданий
Windows. Оно запускается при входе текущего пользователя в интерактивный сеанс;
права SYSTEM и запись Run не используются.
Пока программа работает, она блокирует события мыши. Ctrl+Alt+Shift+M
возвращает управление мышью и удаляет задание, не останавливая эффекты.

Остановить запущенный сейчас процесс: горячая клавиша Ctrl+Alt+Shift+Q, либо завершение
процесса через Диспетчер задач (имя процесса - ORVISimulator.exe после сборки).

Убрать из автозапуска: запустить `ORVISimulator.exe --uninstall`
(удаляет только задание; текущий процесс продолжает работать).

ВАЖНО: предназначено только для запуска на собственном компьютере либо на компьютере
человека, который знает и согласен на розыгрыш. Программа не скрывает и не маскирует
задание автозапуска, не устанавливается как служба, не пытается противодействовать
удалению и не собирает никакие данные - только визуальные/звуковые эффекты и одно
обычное задание Планировщика для текущего пользователя.
"""

import ctypes
import ctypes.wintypes as wintypes
import itertools
import os
import random
import subprocess
import sys
import threading
import time
import webbrowser
import winreg

import pythoncom
import win32api
import win32con
import win32event
import win32gui
import win32com.client

# ---------------------------------------------------------------------------
# Настройки
# ---------------------------------------------------------------------------

INTERVAL_MIN_SEC = 3.0
INTERVAL_MAX_SEC = 7.0

FEVER_CHANCE = 0.18          # шанс дополнительного эффекта "температуры" в каждом цикле
COUGH_SHAKE_DURATION = 1.2   # сколько секунд "трясётся" экран при кашле
COUGH_SHAKE_AMPLITUDE = 10   # амплитуда дрожания окон в пикселях

FEVER_MIN_DURATION = 2.5
FEVER_MAX_DURATION = 4.5
FEVER_MIN_ALPHA = 0.12
FEVER_MAX_ALPHA = 0.30

VOLUME_PRESSES_MIN = 4
VOLUME_PRESSES_MAX = 9

INTERNAL_WINDOW_MARK = "__orvi_internal__"

MUTEX_NAME = "Global\\ORVISimulator_SingleInstance_Mutex"

QUIT_HOTKEY_ID = 1
DISABLE_HOTKEY_ID = 2
HOTKEY_MODIFIERS = win32con.MOD_CONTROL | win32con.MOD_ALT | win32con.MOD_SHIFT
QUIT_VK = ord("Q")
DISABLE_VK = ord("M")
WH_MOUSE_LL = 14

# Классы окон, которые нельзя трогать при "кашле" (таскбар, рабочий стол и т.п.)
SHAKE_CLASS_BLACKLIST = {
    "Shell_TrayWnd",
    "Shell_SecondaryTrayWnd",
    "Progman",
    "WorkerW",
    "Button",
    "NotifyIconOverflowWindow",
    "Windows.UI.Core.CoreWindow",
}

stop_event = threading.Event()


# ---------------------------------------------------------------------------
# Автозапуск: видимое задание Планировщика для интерактивного сеанса пользователя
# ---------------------------------------------------------------------------

AUTOSTART_TASK_NAME = "ORVISimulator"
LEGACY_RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
TASK_LOGON_TRIGGER = 9
TASK_EXEC_ACTION = 0
TASK_CREATE_OR_UPDATE = 6
TASK_LOGON_INTERACTIVE_TOKEN = 3
TASK_INSTANCES_IGNORE_NEW = 2


def _get_exe_path():
    if getattr(sys, "frozen", False):
        return sys.executable
    return os.path.abspath(__file__)


def _task_action():
    if getattr(sys, "frozen", False):
        return _get_exe_path(), ""
    return sys.executable, subprocess.list2cmdline([_get_exe_path()])


def _remove_legacy_run_entry():
    """Очистить запись автозагрузки, созданную предыдущими версиями проекта."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, LEGACY_RUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, AUTOSTART_TASK_NAME)
    except OSError:
        pass


def _register_autostart_task():
    service = win32com.client.Dispatch("Schedule.Service")
    service.Connect()
    folder = service.GetFolder("\\")
    definition = service.NewTask(0)
    user = win32api.GetUserNameEx(win32api.NameSamCompatible)

    definition.RegistrationInfo.Description = "ORVI Simulator: запуск при входе пользователя"
    definition.Principal.UserId = user
    definition.Principal.LogonType = TASK_LOGON_INTERACTIVE_TOKEN
    definition.Principal.RunLevel = 0  # Наименьшие права, без повышения до администратора.
    definition.Settings.Enabled = True
    definition.Settings.Hidden = False
    definition.Settings.ExecutionTimeLimit = "PT0S"
    definition.Settings.DisallowStartIfOnBatteries = False
    definition.Settings.StopIfGoingOnBatteries = False
    definition.Settings.MultipleInstances = TASK_INSTANCES_IGNORE_NEW

    trigger = definition.Triggers.Create(TASK_LOGON_TRIGGER)
    trigger.UserId = user
    trigger.Enabled = True

    executable, arguments = _task_action()
    action = definition.Actions.Create(TASK_EXEC_ACTION)
    action.Path = executable
    action.Arguments = arguments
    action.WorkingDirectory = os.path.dirname(_get_exe_path())

    folder.RegisterTaskDefinition(
        AUTOSTART_TASK_NAME, definition, TASK_CREATE_OR_UPDATE,
        user, None, TASK_LOGON_INTERACTIVE_TOKEN,
    )


def enable_autostart():
    """Создать задание входа без прав SYSTEM и без хранения пароля."""
    pythoncom.CoInitialize()
    try:
        _register_autostart_task()
        return True
    except Exception:
        return False
    finally:
        pythoncom.CoUninitialize()


def _delete_autostart_task():
    service = win32com.client.Dispatch("Schedule.Service")
    service.Connect()
    service.GetFolder("\\").DeleteTask(AUTOSTART_TASK_NAME, 0)


def disable_autostart():
    pythoncom.CoInitialize()
    try:
        _delete_autostart_task()
    except Exception:
        pass
    finally:
        pythoncom.CoUninitialize()
        _remove_legacy_run_entry()


# ---------------------------------------------------------------------------
# Звук: проигрывание mp3 через MCI (winmm.dll), без сторонних библиотек
# ---------------------------------------------------------------------------

def _resource_path(relative_path):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)


SNEEZE_SOUND = _resource_path(os.path.join("sounds", "sneeze.mp3"))
COUGH_SOUND = _resource_path(os.path.join("sounds", "cough.mp3"))
BLOW_NOSE_SOUND = _resource_path(os.path.join("sounds", "blow_nose.mp3"))

_sound_alias_counter = itertools.count()
_MAX_SOUND_LIFETIME_SEC = 8.0  # с запасом дольше самого длинного звука


def _play_sound_async(path):
    if not os.path.isfile(path):
        return

    alias = f"orvisound{next(_sound_alias_counter)}"
    winmm = ctypes.windll.winmm

    try:
        winmm.mciSendStringW(f'open "{path}" type mpegvideo alias {alias}', None, 0, None)
        winmm.mciSendStringW(f"play {alias}", None, 0, None)
    except Exception:
        return

    def _cleanup():
        time.sleep(_MAX_SOUND_LIFETIME_SEC)
        try:
            winmm.mciSendStringW(f"close {alias}", None, 0, None)
        except Exception:
            pass

    threading.Thread(target=_cleanup, daemon=True).start()


# ---------------------------------------------------------------------------
# Чих: открыть случайное приложение
# ---------------------------------------------------------------------------

def _open_default_browser():
    webbrowser.open("about:blank")


def _open_task_manager():
    subprocess.Popen(["taskmgr.exe"], shell=False)


def _open_onscreen_keyboard():
    subprocess.Popen(["osk.exe"], shell=False)


def _open_explorer():
    subprocess.Popen(["explorer.exe"], shell=False)


SNEEZE_ACTIONS = [
    _open_default_browser,
    _open_task_manager,
    _open_onscreen_keyboard,
    _open_explorer,
]


def do_sneeze():
    _play_sound_async(SNEEZE_SOUND)
    action = random.choice(SNEEZE_ACTIONS)
    try:
        action()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Кашель: дрожание экрана (дрожат все видимые окна)
# ---------------------------------------------------------------------------

def _enum_shakeable_windows():
    windows = []

    def callback(hwnd, _extra):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        if win32gui.IsIconic(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd)
        if not title or title.startswith(INTERNAL_WINDOW_MARK):
            return True
        try:
            cls = win32gui.GetClassName(hwnd)
        except Exception:
            return True
        if cls in SHAKE_CLASS_BLACKLIST:
            return True
        windows.append(hwnd)
        return True

    win32gui.EnumWindows(callback, None)
    return windows


def do_cough():
    _play_sound_async(COUGH_SOUND)
    windows = _enum_shakeable_windows()
    if not windows:
        return

    originals = {}
    for hwnd in windows:
        try:
            originals[hwnd] = win32gui.GetWindowRect(hwnd)
        except Exception:
            pass

    if not originals:
        return

    end_time = time.time() + COUGH_SHAKE_DURATION
    while time.time() < end_time and not stop_event.is_set():
        for hwnd, (left, top, _right, _bottom) in originals.items():
            try:
                dx = random.randint(-COUGH_SHAKE_AMPLITUDE, COUGH_SHAKE_AMPLITUDE)
                dy = random.randint(-COUGH_SHAKE_AMPLITUDE, COUGH_SHAKE_AMPLITUDE)
                win32gui.SetWindowPos(
                    hwnd, None, left + dx, top + dy, 0, 0,
                    win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE,
                )
            except Exception:
                pass
        time.sleep(0.03)

    for hwnd, (left, top, _right, _bottom) in originals.items():
        try:
            win32gui.SetWindowPos(
                hwnd, None, left, top, 0, 0,
                win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE,
            )
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Сморкание: изменить громкость
# ---------------------------------------------------------------------------

VK_VOLUME_UP = 0xAF
VK_VOLUME_DOWN = 0xAE


def do_blow_nose():
    _play_sound_async(BLOW_NOSE_SOUND)
    key = random.choice([VK_VOLUME_UP, VK_VOLUME_DOWN])
    presses = random.randint(VOLUME_PRESSES_MIN, VOLUME_PRESSES_MAX)
    for _ in range(presses):
        if stop_event.is_set():
            return
        win32api.keybd_event(key, 0, 0, 0)
        win32api.keybd_event(key, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.05)


# ---------------------------------------------------------------------------
# "Температура": лёгкая красная подсветка экрана
# ---------------------------------------------------------------------------

def do_fever():
    import tkinter as tk

    duration = random.uniform(FEVER_MIN_DURATION, FEVER_MAX_DURATION)
    max_alpha = random.uniform(FEVER_MIN_ALPHA, FEVER_MAX_ALPHA)

    root = tk.Tk()
    root.title(INTERNAL_WINDOW_MARK + "fever")
    try:
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        root.geometry(f"{screen_w}x{screen_h}+0+0")
        root.configure(bg="red")
        root.attributes("-alpha", 0.0)

        # Клики не должны мешать работе - окно не активируется и не получает фокус.
        steps = 20
        half = max(duration / 2.0, 0.1)

        for i in range(steps + 1):
            if stop_event.is_set():
                break
            alpha = max_alpha * (i / steps)
            root.attributes("-alpha", alpha)
            root.update()
            time.sleep(half / steps)

        if not stop_event.is_set():
            time.sleep(0.2)

        for i in range(steps, -1, -1):
            if stop_event.is_set():
                break
            alpha = max_alpha * (i / steps)
            root.attributes("-alpha", alpha)
            root.update()
            time.sleep(half / steps)
    except Exception:
        pass
    finally:
        try:
            root.destroy()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Управление: Ctrl+Alt+Shift+Q — выход; Ctrl+Alt+Shift+M — снять блокировку
# ---------------------------------------------------------------------------

def control_listener(ready_event, active_event):
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    hook = None
    mouse_blocked = threading.Event()
    registered_hotkeys = []

    # Прототипы важны для 64-битной сборки: дескриптор хука и результат
    # CallNextHookEx не должны обрезаться до 32 бит.
    hook_proc_type = ctypes.WINFUNCTYPE(
        ctypes.c_ssize_t, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM,
    )
    user32.SetWindowsHookExW.argtypes = [
        ctypes.c_int, hook_proc_type, wintypes.HANDLE, wintypes.DWORD,
    ]
    user32.SetWindowsHookExW.restype = wintypes.HANDLE
    user32.CallNextHookEx.argtypes = [
        wintypes.HANDLE, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM,
    ]
    user32.CallNextHookEx.restype = ctypes.c_ssize_t
    user32.UnhookWindowsHookEx.argtypes = [wintypes.HANDLE]
    user32.UnhookWindowsHookEx.restype = wintypes.BOOL
    kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
    kernel32.GetModuleHandleW.restype = wintypes.HANDLE

    def mouse_hook(n_code, w_param, l_param):
        if n_code >= 0 and mouse_blocked.is_set():
            return 1  # Погасить перемещение, кнопки и колесо мыши.
        return user32.CallNextHookEx(hook, n_code, w_param, l_param)

    callback = hook_proc_type(mouse_hook)

    try:
        for hotkey_id, virtual_key in (
            (QUIT_HOTKEY_ID, QUIT_VK),
            (DISABLE_HOTKEY_ID, DISABLE_VK),
        ):
            if not user32.RegisterHotKey(None, hotkey_id, HOTKEY_MODIFIERS, virtual_key):
                return
            registered_hotkeys.append(hotkey_id)

        installed = enable_autostart()
        _remove_legacy_run_entry()
        if not installed:
            return

        hook = user32.SetWindowsHookExW(
            WH_MOUSE_LL, callback, kernel32.GetModuleHandleW(None), 0,
        )
        if not hook:
            disable_autostart()
            return
        mouse_blocked.set()

        active_event.set()
        ready_event.set()

        msg = wintypes.MSG()
        while not stop_event.is_set():
            result = user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1)  # PM_REMOVE
            if result and msg.message == win32con.WM_HOTKEY:
                if msg.wParam == QUIT_HOTKEY_ID:
                    stop_event.set()
                    break
                if msg.wParam == DISABLE_HOTKEY_ID and DISABLE_HOTKEY_ID in registered_hotkeys:
                    mouse_blocked.clear()
                    if user32.UnhookWindowsHookEx(hook):
                        hook = None
                    disable_autostart()
                    user32.UnregisterHotKey(None, DISABLE_HOTKEY_ID)
                    registered_hotkeys.remove(DISABLE_HOTKEY_ID)
            time.sleep(0.01)
    finally:
        mouse_blocked.clear()
        ready_event.set()
        if hook:
            user32.UnhookWindowsHookEx(hook)
        for hotkey_id in registered_hotkeys:
            user32.UnregisterHotKey(None, hotkey_id)
        stop_event.set()


# ---------------------------------------------------------------------------
# Основной цикл
# ---------------------------------------------------------------------------

def scheduler_loop():
    actions = [do_sneeze, do_cough, do_blow_nose]

    while not stop_event.is_set():
        wait_time = random.uniform(INTERVAL_MIN_SEC, INTERVAL_MAX_SEC)
        if stop_event.wait(wait_time):
            break

        action = random.choice(actions)
        try:
            action()
        except Exception:
            pass

        if stop_event.is_set():
            break

        if random.random() < FEVER_CHANCE:
            try:
                do_fever()
            except Exception:
                pass


def acquire_single_instance_lock():
    mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
    last_error = win32api.GetLastError()
    if last_error == 183:  # ERROR_ALREADY_EXISTS
        return None
    return mutex


def main():
    if len(sys.argv) > 1 and sys.argv[1].lower() in ("--uninstall", "/uninstall", "-u"):
        disable_autostart()
        sys.exit(0)

    mutex = acquire_single_instance_lock()
    if mutex is None:
        sys.exit(0)

    ready_event = threading.Event()
    active_event = threading.Event()
    listener = threading.Thread(
        target=control_listener, args=(ready_event, active_event), daemon=True,
    )
    listener.start()
    ready_event.wait(timeout=5)
    if not active_event.is_set():
        return  # Без обеих горячих клавиш и хука мыши запуск небезопасен.

    try:
        scheduler_loop()
    finally:
        stop_event.set()
        listener.join(timeout=2)


if __name__ == "__main__":
    main()
