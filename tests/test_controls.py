"""Проверки аварийного отключения без установленного Windows API."""

import ctypes
import importlib.util
from pathlib import Path
import sys
import threading
import types
import unittest
from unittest import mock


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def load_orvi():
    winreg = types.ModuleType("winreg")
    win32api = types.ModuleType("win32api")
    win32event = types.ModuleType("win32event")
    win32gui = types.ModuleType("win32gui")
    win32con = types.ModuleType("win32con")
    pythoncom = types.ModuleType("pythoncom")
    pythoncom.CoInitialize = lambda: None
    pythoncom.CoUninitialize = lambda: None
    win32com = types.ModuleType("win32com")
    win32com.__path__ = []
    win32com_client = types.ModuleType("win32com.client")
    win32com.client = win32com_client
    win32con.MOD_CONTROL = 2
    win32con.MOD_ALT = 1
    win32con.MOD_SHIFT = 4
    win32con.WM_HOTKEY = 0x312

    with mock.patch.dict(sys.modules, {
        "winreg": winreg,
        "win32api": win32api,
        "win32event": win32event,
        "win32gui": win32gui,
        "win32con": win32con,
        "pythoncom": pythoncom,
        "win32com": win32com,
        "win32com.client": win32com_client,
    }):
        spec = importlib.util.spec_from_file_location("orvi_under_test", PROJECT_ROOT / "orvi.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
    return module


class ApiFunction:
    def __init__(self, impl):
        self.impl = impl
        self.argtypes = None
        self.restype = None

    def __call__(self, *args):
        return self.impl(*args)


class ControlsTest(unittest.TestCase):
    def setUp(self):
        self.orvi = load_orvi()

    def test_task_action_uses_python_for_source_and_exe_for_build(self):
        with mock.patch.object(self.orvi.sys, "frozen", False, create=True), \
             mock.patch.object(self.orvi.sys, "executable", r"C:\Python\pythonw.exe"), \
             mock.patch.object(self.orvi, "_get_exe_path", return_value=r"C:\ORVI Simulator\orvi.py"):
            self.assertEqual(
                self.orvi._task_action(),
                (r"C:\Python\pythonw.exe", '"C:\\ORVI Simulator\\orvi.py"'),
            )

        with mock.patch.object(self.orvi.sys, "frozen", True, create=True), \
             mock.patch.object(self.orvi, "_get_exe_path", return_value=r"C:\ORVI\ORVISimulator.exe"):
            self.assertEqual(self.orvi._task_action(), (r"C:\ORVI\ORVISimulator.exe", ""))

    def test_task_is_visible_and_runs_in_user_session_on_logon(self):
        module = self.orvi
        events = []
        triggers = []
        actions = []
        folder = types.SimpleNamespace(
            RegisterTaskDefinition=lambda *args: events.append(("register", args)),
        )
        definition = types.SimpleNamespace(
            RegistrationInfo=types.SimpleNamespace(),
            Principal=types.SimpleNamespace(),
            Settings=types.SimpleNamespace(),
            Triggers=types.SimpleNamespace(Create=lambda kind: triggers.append((kind, types.SimpleNamespace())) or triggers[-1][1]),
            Actions=types.SimpleNamespace(Create=lambda kind: actions.append((kind, types.SimpleNamespace())) or actions[-1][1]),
        )
        service = types.SimpleNamespace(
            Connect=lambda: events.append("connected"),
            GetFolder=lambda path: folder if path == "\\" else None,
            NewTask=lambda flags: definition,
        )

        with mock.patch.object(module.pythoncom, "CoInitialize", side_effect=lambda: events.append("com on")), \
             mock.patch.object(module.pythoncom, "CoUninitialize", side_effect=lambda: events.append("com off")), \
             mock.patch.object(module.win32com.client, "Dispatch", return_value=service, create=True), \
             mock.patch.object(module.win32api, "GetUserNameEx", return_value=r"PC\Student", create=True), \
             mock.patch.object(module.win32api, "NameSamCompatible", 2, create=True), \
             mock.patch.object(module, "_task_action", return_value=(r"C:\ORVI\ORVISimulator.exe", "")), \
             mock.patch.object(module, "_get_exe_path", return_value=r"C:\ORVI\ORVISimulator.exe"):
            self.assertTrue(module.enable_autostart())

        self.assertEqual(triggers[0][0], module.TASK_LOGON_TRIGGER)
        self.assertEqual(triggers[0][1].UserId, r"PC\Student")
        self.assertEqual(actions[0][1].Path, r"C:\ORVI\ORVISimulator.exe")
        self.assertFalse(definition.Settings.Hidden)
        self.assertEqual(definition.Settings.ExecutionTimeLimit, "PT0S")
        self.assertEqual(definition.Principal.LogonType, module.TASK_LOGON_INTERACTIVE_TOKEN)
        self.assertEqual(definition.Principal.RunLevel, 0)
        registered = events[-2][1]
        self.assertEqual(registered[0], module.AUTOSTART_TASK_NAME)
        self.assertEqual(registered[5], module.TASK_LOGON_INTERACTIVE_TOKEN)
        self.assertEqual(events[-1], "com off")

    def test_uninstall_deletes_task_and_old_run_entry(self):
        module = self.orvi
        events = []
        folder = types.SimpleNamespace(
            DeleteTask=lambda name, flags: events.append(("deleted", name, flags)),
        )
        service = types.SimpleNamespace(
            Connect=lambda: None,
            GetFolder=lambda path: folder if path == "\\" else None,
        )

        with mock.patch.object(module.pythoncom, "CoInitialize", side_effect=lambda: events.append("com on")), \
             mock.patch.object(module.pythoncom, "CoUninitialize", side_effect=lambda: events.append("com off")), \
             mock.patch.object(module.win32com.client, "Dispatch", return_value=service, create=True), \
             mock.patch.object(module, "_remove_legacy_run_entry", side_effect=lambda: events.append("old Run removed")):
            module.disable_autostart()

        self.assertIn(("deleted", module.AUTOSTART_TASK_NAME, 0), events)
        self.assertEqual(events[-2:], ["com off", "old Run removed"])

    def test_disable_hotkey_restores_mouse_and_removes_autostart_without_quitting(self):
        module = self.orvi
        events = []
        hook_callback = []
        registered = []
        message_ids = iter([
            module.DISABLE_HOTKEY_ID,
            module.DISABLE_HOTKEY_ID,  # Уже поставленное в очередь повторное нажатие.
            module.QUIT_HOTKEY_ID,
        ])

        def register(_hwnd, hotkey_id, _modifiers, _key):
            registered.append(hotkey_id)
            return True

        def set_hook(_kind, callback, _module_handle, _thread_id):
            hook_callback.append(callback)
            events.append("hooked")
            return 42

        def install_task():
            self.assertEqual(hook_callback, [])
            events.append("autostart on")
            return True

        def next_message(msg_pointer, _hwnd, _min, _max, _remove):
            msg_pointer._obj.message = module.win32con.WM_HOTKEY
            msg_pointer._obj.wParam = next(message_ids)
            if msg_pointer._obj.wParam == module.DISABLE_HOTKEY_ID and "unhooked" not in events:
                self.assertEqual(hook_callback[0](0, 0, 0), 1)
            else:
                self.assertFalse(module.stop_event.is_set())
                self.assertEqual(hook_callback[0](0, 0, 0), 0)
            return 1

        user32 = types.SimpleNamespace(
            RegisterHotKey=register,
            UnregisterHotKey=lambda _hwnd, hotkey_id: events.append(("unregistered", hotkey_id)),
            SetWindowsHookExW=ApiFunction(set_hook),
            CallNextHookEx=ApiFunction(lambda *_args: 0),
            UnhookWindowsHookEx=ApiFunction(lambda _hook: events.append("unhooked") or True),
            PeekMessageW=next_message,
        )
        kernel32 = types.SimpleNamespace(GetModuleHandleW=ApiFunction(lambda _name: 123))
        ready_event = threading.Event()
        active_event = threading.Event()

        with mock.patch.object(ctypes, "windll", types.SimpleNamespace(user32=user32, kernel32=kernel32), create=True), \
             mock.patch.object(ctypes, "WINFUNCTYPE", lambda *_args: lambda fn: fn, create=True), \
             mock.patch.object(module, "enable_autostart", side_effect=install_task), \
             mock.patch.object(module, "disable_autostart", side_effect=lambda: events.append("autostart off")), \
             mock.patch.object(module, "_remove_legacy_run_entry", side_effect=lambda: events.append("old Run removed")), \
             mock.patch.object(module.time, "sleep"):
            module.control_listener(ready_event, active_event)

        self.assertTrue(ready_event.is_set())
        self.assertTrue(active_event.is_set())
        self.assertEqual(registered, [module.QUIT_HOTKEY_ID, module.DISABLE_HOTKEY_ID])
        self.assertLess(events.index("unhooked"), events.index("autostart off"))
        self.assertEqual(events.count("autostart off"), 1)
        self.assertEqual(events.count("old Run removed"), 1)
        self.assertIn(("unregistered", module.DISABLE_HOTKEY_ID), events)
        self.assertIn(("unregistered", module.QUIT_HOTKEY_ID), events)

    def test_mouse_is_not_blocked_if_escape_hotkey_is_unavailable(self):
        module = self.orvi
        hooked = []
        unregistered = []
        user32 = types.SimpleNamespace(
            RegisterHotKey=lambda _hwnd, hotkey_id, _modifiers, _key: hotkey_id == module.QUIT_HOTKEY_ID,
            UnregisterHotKey=lambda _hwnd, hotkey_id: unregistered.append(hotkey_id),
            SetWindowsHookExW=ApiFunction(lambda *_args: hooked.append(True)),
            CallNextHookEx=ApiFunction(lambda *_args: 0),
            UnhookWindowsHookEx=ApiFunction(lambda *_args: True),
        )
        kernel32 = types.SimpleNamespace(GetModuleHandleW=ApiFunction(lambda _name: 123))
        ready_event = threading.Event()
        active_event = threading.Event()

        with mock.patch.object(ctypes, "windll", types.SimpleNamespace(user32=user32, kernel32=kernel32), create=True), \
             mock.patch.object(ctypes, "WINFUNCTYPE", lambda *_args: lambda fn: fn, create=True):
            module.control_listener(ready_event, active_event)

        self.assertTrue(ready_event.is_set())
        self.assertFalse(active_event.is_set())
        self.assertEqual(hooked, [])
        self.assertEqual(unregistered, [module.QUIT_HOTKEY_ID])

    def test_mouse_is_not_blocked_if_scheduled_task_cannot_be_created(self):
        module = self.orvi
        hooked = []
        user32 = types.SimpleNamespace(
            RegisterHotKey=lambda *_args: True,
            UnregisterHotKey=lambda *_args: None,
            SetWindowsHookExW=ApiFunction(lambda *_args: hooked.append(True)),
            CallNextHookEx=ApiFunction(lambda *_args: 0),
            UnhookWindowsHookEx=ApiFunction(lambda *_args: True),
        )
        kernel32 = types.SimpleNamespace(GetModuleHandleW=ApiFunction(lambda _name: 123))
        ready_event = threading.Event()
        active_event = threading.Event()

        with mock.patch.object(ctypes, "windll", types.SimpleNamespace(user32=user32, kernel32=kernel32), create=True), \
             mock.patch.object(ctypes, "WINFUNCTYPE", lambda *_args: lambda fn: fn, create=True), \
             mock.patch.object(module, "enable_autostart", return_value=False), \
             mock.patch.object(module, "_remove_legacy_run_entry"):
            module.control_listener(ready_event, active_event)

        self.assertTrue(ready_event.is_set())
        self.assertFalse(active_event.is_set())
        self.assertEqual(hooked, [])
