"""
mic_hotkeys_cominit_fixed.py
Global hotkeys to control microphone volume (Windows).
This version ensures COM is initialized inside each hotkey handler thread.

Hotkeys:
  Ctrl+Alt+Up    -> increase mic volume by 5%
  Ctrl+Alt+Down  -> decrease mic volume by 5%
  Ctrl+Alt+M     -> toggle mute/unmute
  Ctrl+Alt+Q     -> stop hotkeys (clean shutdown)
"""
import sys
import time
import platform
import threading
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize
from comtypes.client import CreateObject
from comtypes import GUID
from functools import wraps

if platform.system() != "Windows":
    raise SystemExit("This script runs only on Windows.")

# 3rd-party libraries
try:
    import keyboard   # pip install keyboard
except Exception:
    raise SystemExit("Install required package: pip install keyboard")

try:
    from pycaw.pycaw import IAudioEndpointVolume, IMMDeviceEnumerator
except Exception:
    raise SystemExit("Install required packages: pip install pycaw comtypes")

# Constants
eCapture = 1
eConsole = 0
STEP_PERCENT = 5.0

# Stop event for clean shutdown
stop_event = threading.Event()

# Helper: create IMMDeviceEnumerator (typed)
def _create_mmdevice_enumerator():
    try:
        return CreateObject("MMDeviceEnumerator.MMDeviceEnumerator", interface=IMMDeviceEnumerator)
    except Exception:
        clsid = GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")
        return CreateObject(clsid, interface=IMMDeviceEnumerator)

def _get_volume_interface_for_default():
    """
    Returns an IAudioEndpointVolume pointer for the default capture device.
    Caller must ensure COM is initialized in the calling thread.
    """
    enumerator = _create_mmdevice_enumerator()
    default_device = enumerator.GetDefaultAudioEndpoint(eCapture, eConsole)
    iface = default_device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    return cast(iface, POINTER(IAudioEndpointVolume))

def _percent_to_scalar(p):
    return max(0.0, min(1.0, p / 100.0))

def _scalar_to_percent(s):
    return max(0.0, min(100.0, s * 100.0))

# Decorator to ensure COM initialized per-thread for hotkey handlers
def ensure_com(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        CoInitialize()
        try:
            return func(*args, **kwargs)
        except Exception:
            # Print traceback for easier debugging
            import traceback
            print("Exception in handler:", file=sys.stderr)
            traceback.print_exc()
        finally:
            # Uninitialize COM in this thread when handler completes
            try:
                CoUninitialize()
            except Exception:
                pass
    return wrapper

# Handlers (each will run with COM initialized)
@ensure_com
def increase_volume():
    vol = _get_volume_interface_for_default()
    cur = float(vol.GetMasterVolumeLevelScalar())
    cur_pct = _scalar_to_percent(cur)
    new_pct = min(100.0, cur_pct + STEP_PERCENT)
    vol.SetMasterVolumeLevelScalar(_percent_to_scalar(new_pct), None)
    print(f"[+] Mic volume -> {new_pct:.0f}%")

@ensure_com
def decrease_volume():
    vol = _get_volume_interface_for_default()
    cur = float(vol.GetMasterVolumeLevelScalar())
    cur_pct = _scalar_to_percent(cur)
    new_pct = max(0.0, cur_pct - STEP_PERCENT)
    vol.SetMasterVolumeLevelScalar(_percent_to_scalar(new_pct), None)
    print(f"[-] Mic volume -> {new_pct:.0f}%")

@ensure_com
def toggle_mute():
    vol = _get_volume_interface_for_default()
    cur_mute = bool(vol.GetMute())
    new_mute = not cur_mute
    vol.SetMute(1 if new_mute else 0, None)
    # Print the new state (M when muted, U when unmuted)
    print(f"[{'M' if new_mute else 'U'}] Mic muted -> {new_mute}")

def quit_program():
    """
    Signal the main loop to stop and unhook hotkeys.
    Avoids calling sys.exit() directly from a hotkey handler.
    """
    print("Stopping hotkeys (quit requested)...")
    try:
        keyboard.unhook_all_hotkeys()
    except Exception:
        pass
    stop_event.set()

def register_hotkeys():
    # Avoid duplicate registrations if re-run
    try:
        keyboard.unhook_all_hotkeys()
    except Exception:
        pass

    keyboard.add_hotkey('ctrl+alt+up', increase_volume)
    keyboard.add_hotkey('ctrl+alt+down', decrease_volume)
    keyboard.add_hotkey('ctrl+alt+m', toggle_mute)
    keyboard.add_hotkey('ctrl+alt+q', quit_program)

if __name__ == "__main__":
    print("Registering hotkeys...")
    register_hotkeys()

    # Try to print initial state (initialize COM briefly here)
    CoInitialize()
    try:
        try:
            vol = _get_volume_interface_for_default()
            print(
                f"Initial mic volume: {_scalar_to_percent(float(vol.GetMasterVolumeLevelScalar())):.0f}%"
                f"  muted={bool(vol.GetMute())}"
            )
        except Exception as e:
            print("Could not read initial mic state:", e)
    finally:
        CoUninitialize()

    print("Hotkeys active. Press Ctrl+Alt+Up/Down to change mic, Ctrl+Alt+M to toggle mute, Ctrl+Alt+Q to quit.")
    try:
        # Main loop: wait until quit_program() sets stop_event
        while not stop_event.wait(timeout=0.1):
            # keep the main thread responsive; checks stop_event every 0.1s
            pass
    except KeyboardInterrupt:
        print("Interrupted by user; exiting.")
    finally:
        # clean up in any case
        try:
            keyboard.unhook_all_hotkeys()
        except Exception:
            pass
        print("Exited mic hotkeys.")
