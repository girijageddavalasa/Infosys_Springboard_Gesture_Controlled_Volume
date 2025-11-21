"""
mic_hotkeys_stable_v2.py
Fixes: 'Access Violation' / COM Release errors.
Strategy: Initializes Audio Interface ONCE on the main thread and caches it.
"""
import sys
import threading
import platform
from ctypes import POINTER, cast
from functools import wraps

# Windows COM libraries
from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize
from comtypes.client import CreateObject
from comtypes import GUID

# --- Kivy Imports ---
try:
    from kivy.app import App
    from kivy.uix.boxlayout import BoxLayout
    from kivy.clock import Clock
    from kivy.config import Config
    from kivy.lang import Builder
except ImportError:
    raise SystemExit("Install required package: pip install kivy")

# --- Config ---
Config.set('graphics', 'width', '450')
Config.set('graphics', 'height', '350')
Config.set('graphics', 'resizable', '0')
Config.set('input', 'mouse', 'mouse,multitouch_on_demand')

if platform.system() != "Windows":
    raise SystemExit("This script runs only on Windows.")

# 3rd-party libraries
try:
    import keyboard
except Exception:
    raise SystemExit("Install required package: pip install keyboard")

try:
    from pycaw.pycaw import IAudioEndpointVolume, IMMDeviceEnumerator
except Exception:
    raise SystemExit("Install required packages: pip install pycaw comtypes")

# --- Audio Helper Functions ---
# We define these globally but will use them carefully
eCapture = 1
#eRender = 0
eConsole = 0
STEP_PERCENT = 5.0

def _create_mmdevice_enumerator():
    try:
        return CreateObject("MMDeviceEnumerator.MMDeviceEnumerator", interface=IMMDeviceEnumerator)
    except Exception:
        clsid = GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")
        return CreateObject(clsid, interface=IMMDeviceEnumerator)

def _get_new_interface():
    """Creates a NEW interface. Use sparingly."""
    enumerator = _create_mmdevice_enumerator()
    default_device = enumerator.GetDefaultAudioEndpoint(eCapture, eConsole)
    iface = default_device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    return cast(iface, POINTER(IAudioEndpointVolume))

def _scalar_to_percent(s):
    return max(0.0, min(100.0, s * 100.0))

def _percent_to_scalar(p):
    return max(0.0, min(1.0, p / 100.0))

# --- Hotkey Handlers (Background Threads) ---
# Hotkeys run on separate threads, so they MUST create their own short-lived connections.
# We wrap them to ensure safe creation/destruction.

def safe_hotkey_execution(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            CoInitialize() # Init COM for this thread
            # Create a fresh interface just for this click
            vol = _get_new_interface()
            
            # Execute logic
            func(vol, *args, **kwargs)
            
            # Explicitly release helps prevent GC errors in threads
            del vol
        except Exception as e:
            print(f"Hotkey Error: {e}")
        finally:
            CoUninitialize() # Clean up COM
    return wrapper

@safe_hotkey_execution
def increase_volume(vol):
    cur = float(vol.GetMasterVolumeLevelScalar())
    new_pct = min(100.0, _scalar_to_percent(cur) + STEP_PERCENT)
    vol.SetMasterVolumeLevelScalar(_percent_to_scalar(new_pct), None)
    print(f"[Hotkey] Vol Up -> {new_pct:.0f}%")

@safe_hotkey_execution
def decrease_volume(vol):
    cur = float(vol.GetMasterVolumeLevelScalar())
    new_pct = max(0.0, _scalar_to_percent(cur) - STEP_PERCENT)
    vol.SetMasterVolumeLevelScalar(_percent_to_scalar(new_pct), None)
    print(f"[Hotkey] Vol Down -> {new_pct:.0f}%")

@safe_hotkey_execution
def toggle_mute(vol):
    new_mute = not bool(vol.GetMute())
    vol.SetMute(1 if new_mute else 0, None)
    print(f"[Hotkey] Mute -> {new_mute}")

def quit_program():
    print("Quitting...")
    try:
        keyboard.unhook_all_hotkeys()
    except:
        pass
    App.get_running_app().stop()

def register_hotkeys():
    try:
        keyboard.unhook_all_hotkeys()
    except:
        pass
    keyboard.add_hotkey('ctrl+alt+up', increase_volume)
    keyboard.add_hotkey('ctrl+alt+down', decrease_volume)
    keyboard.add_hotkey('ctrl+alt+m', toggle_mute)
    keyboard.add_hotkey('ctrl+alt+q', quit_program)

# --- GUI Layout (KV) ---
KV_DESIGN = """
#:import Factory kivy.factory.Factory

#:set bg_color (0.11, 0.11, 0.13, 1)
#:set card_color (0.16, 0.16, 0.18, 1)
#:set accent_color (0.3, 0.6, 1.0, 1)

<RoundedButton@Button>:
    background_color: 0,0,0,0
    bg_rect_color: (0.8, 0.2, 0.2, 1)
    canvas.before:
        Color:
            rgba: self.bg_rect_color
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [6,]

<MicController>:
    orientation: 'vertical'
    padding: 25
    spacing: 15
    canvas.before:
        Color:
            rgba: bg_color
        Rectangle:
            pos: self.pos
            size: self.size

    Label:
        text: "SYSTEM MIC CONTROL"
        font_size: '14sp'
        bold: True
        color: (0.5, 0.5, 0.5, 1)
        size_hint_y: 0.1

    BoxLayout:
        orientation: 'vertical'
        size_hint_y: 0.5
        padding: 20
        spacing: 10
        canvas.before:
            Color:
                rgba: card_color
            RoundedRectangle:
                pos: self.pos
                size: self.size
                radius: [10,]

        Label:
            id: status_lbl
            text: "--"
            font_size: '36sp'
            bold: True
            color: accent_color
        
        ProgressBar:
            id: vol_bar
            max: 100
            value: 0
            size_hint_y: None
            height: '20dp'
        
        Label:
            id: vol_text_lbl
            text: "Initializing..."
            font_size: '12sp'
            color: (0.6, 0.6, 0.6, 1)

    Label:
        text: "Ctrl+Alt+Up/Down: Volume  |  Ctrl+Alt+M: Mute"
        color: (0.5, 0.5, 0.5, 1)
        font_size: '11sp'
        size_hint_y: 0.1

    RoundedButton:
        text: "QUIT APPLICATION"
        font_size: '13sp'
        bold: True
        size_hint_y: 0.2
        bg_rect_color: (0.2, 0.2, 0.25, 1)
        on_press: root.quit_app()
"""

class MicController(BoxLayout):
    def quit_app(self):
        quit_program()

    def update_ui(self, vol_pct, is_muted):
        self.ids.vol_bar.value = vol_pct
        if is_muted:
            self.ids.status_lbl.text = "MUTED"
            self.ids.status_lbl.color = (1, 0.3, 0.3, 1) # Red
            self.ids.vol_text_lbl.text = "Microphone is Inactive"
        else:
            self.ids.status_lbl.text = f"{vol_pct:.0f}%"
            self.ids.status_lbl.color = (0.3, 0.6, 1.0, 1) # Blue
            self.ids.vol_text_lbl.text = "Microphone Active"

class MicApp(App):
    def build(self):
        self.title = "Mic Control"
        Builder.load_string(KV_DESIGN)
        self.layout = MicController()
        return self.layout

    def on_start(self):
        """
        INITIALIZATION LOGIC
        1. Init COM for Main Thread
        2. Create Interface ONCE (Cached)
        3. Start Clock
        """
        print("App Starting...")
        register_hotkeys()

        # 1. Initialize COM for the MainThread (GUI)
        CoInitialize()

        # 2. Create the persistent connection
        try:
            self.cached_interface = _get_new_interface()
            print("Audio Interface Connected Successfully.")
        except Exception as e:
            print(f"Error connecting to audio: {e}")
            self.cached_interface = None

        # 3. Start updating GUI every 0.2 seconds
        Clock.schedule_interval(self.check_mic_status, 0.2)

    def on_stop(self):
        """CLEANUP LOGIC"""
        print("App Stopping...")
        try:
            keyboard.unhook_all_hotkeys()
        except:
            pass
        
        # Release the cached interface
        self.cached_interface = None
        
        # Uninitialize COM
        CoUninitialize()

    def check_mic_status(self, dt):
        """
        Polls the CACHED interface.
        Does NOT create new objects, preventing GC crashes.
        """
        if not self.cached_interface:
            return

        try:
            # Reuse the existing connection
            vol = self.cached_interface
            current_vol = _scalar_to_percent(float(vol.GetMasterVolumeLevelScalar()))
            current_mute = bool(vol.GetMute())
            
            self.layout.update_ui(current_vol, current_mute)
        except Exception:
            # If connection is lost (rare), try to reconnect once
            try:
                self.cached_interface = _get_new_interface()
            except:
                pass

if __name__ == "__main__":
    MicApp().run()