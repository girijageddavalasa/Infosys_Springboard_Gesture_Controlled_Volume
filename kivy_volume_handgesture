"""
mic_gesture_v3.py
Features: 
1. Kivy GUI for Mic Status.
2. Keyboard Hotkeys (Ctrl+Alt+Up/Down/M/Q).
3. MediaPipe Hand Tracking (Thumb-Index Pinch to Control Volume).
"""
import sys
import threading
import platform
import math
import time
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
    from kivy.properties import BooleanProperty
except ImportError:
    raise SystemExit("Install required package: pip install kivy")

# --- Computer Vision Imports ---
try:
    import cv2
    import mediapipe as mp
except ImportError:
    raise SystemExit("Install required packages: pip install opencv-python mediapipe")

# --- Config ---
Config.set('graphics', 'width', '450')
Config.set('graphics', 'height', '420') # Increased height slightly for toggle button
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
eCapture = 1
eConsole = 0
STEP_PERCENT = 5.0

def _create_mmdevice_enumerator():
    try:
        return CreateObject("MMDeviceEnumerator.MMDeviceEnumerator", interface=IMMDeviceEnumerator)
    except Exception:
        clsid = GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")
        return CreateObject(clsid, interface=IMMDeviceEnumerator)

def _get_new_interface():
    """Creates a NEW interface."""
    enumerator = _create_mmdevice_enumerator()
    default_device = enumerator.GetDefaultAudioEndpoint(eCapture, eConsole)
    iface = default_device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    return cast(iface, POINTER(IAudioEndpointVolume))

def _scalar_to_percent(s):
    return max(0.0, min(100.0, s * 100.0))

def _percent_to_scalar(p):
    return max(0.0, min(1.0, p / 100.0))

# --- Hotkey Handlers (Background Threads) ---
def safe_hotkey_execution(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            CoInitialize()
            vol = _get_new_interface()
            func(vol, *args, **kwargs)
            del vol
        except Exception as e:
            print(f"Hotkey Error: {e}")
        finally:
            CoUninitialize()
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
    bg_rect_color: (0.2, 0.2, 0.25, 1)
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
        text: "SYSTEM MIC & GESTURE CONTROL"
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

    BoxLayout:
        size_hint_y: 0.15
        orientation: 'horizontal'
        spacing: 10
        
        RoundedButton:
            id: btn_cam
            text: "ENABLE CAMERA" if not root.camera_active else "DISABLE CAMERA"
            bg_rect_color: (0.3, 0.6, 0.3, 1) if root.camera_active else (0.2, 0.2, 0.25, 1)
            on_press: root.toggle_camera()

        RoundedButton:
            text: "QUIT APP"
            bg_rect_color: (0.8, 0.2, 0.2, 1)
            on_press: root.quit_app()

    Label:
        text: "Pinch Thumb & Index to control volume | Ctrl+Alt+Up/Down"
        color: (0.5, 0.5, 0.5, 1)
        font_size: '11sp'
        size_hint_y: 0.1
"""

class MicController(BoxLayout):
    camera_active = BooleanProperty(False)

    def quit_app(self):
        quit_program()

    def toggle_camera(self):
        app = App.get_running_app()
        if not app.is_camera_running:
            app.start_camera_thread()
            self.camera_active = True
        else:
            app.stop_camera_thread()
            self.camera_active = False

    def update_ui(self, vol_pct, is_muted):
        self.ids.vol_bar.value = vol_pct
        if is_muted:
            self.ids.status_lbl.text = "MUTED"
            self.ids.status_lbl.color = (1, 0.3, 0.3, 1)
            self.ids.vol_text_lbl.text = "Microphone is Inactive"
        else:
            self.ids.status_lbl.text = f"{vol_pct:.0f}%"
            self.ids.status_lbl.color = (0.3, 0.6, 1.0, 1)
            self.ids.vol_text_lbl.text = "Microphone Active"

class MicApp(App):
    def build(self):
        self.title = "Mic Gesture Control"
        Builder.load_string(KV_DESIGN)
        self.layout = MicController()
        return self.layout

    def on_start(self):
        print("App Starting...")
        register_hotkeys()

        # 1. Main Thread COM
        CoInitialize()
        try:
            self.cached_interface = _get_new_interface()
            print("Main Thread Audio Interface Connected.")
        except Exception:
            self.cached_interface = None

        # 2. Camera Thread State
        self.is_camera_running = False
        self.cam_thread = None

        # 3. Start Polling UI
        Clock.schedule_interval(self.check_mic_status, 0.2)

    def on_stop(self):
        print("App Stopping...")
        self.stop_camera_thread()
        try:
            keyboard.unhook_all_hotkeys()
        except:
            pass
        self.cached_interface = None
        CoUninitialize()

    # --- Camera Logic ---
    def start_camera_thread(self):
        if self.is_camera_running:
            return
        self.is_camera_running = True
        self.cam_thread = threading.Thread(target=self._camera_loop, daemon=True)
        self.cam_thread.start()

    def stop_camera_thread(self):
        self.is_camera_running = False
        if self.cam_thread:
            self.cam_thread.join(timeout=1.0)
            self.cam_thread = None
        cv2.destroyAllWindows()

    def _camera_loop(self):
        """
        Runs in a separate thread using MediaPipe.
        Maintains its OWN COM connection for speed.
        """
        # 1. Init COM for this thread
        CoInitialize()
        try:
            cam_vol_interface = _get_new_interface()
        except Exception as e:
            print(f"Camera Thread Audio Error: {e}")
            CoUninitialize()
            return

        # 2. Setup MediaPipe
        mp_hands = mp.solutions.hands
        hands = mp_hands.Hands(
            min_detection_confidence=0.7,
            min_tracking_confidence=0.7,
            max_num_hands=1
        )
        cap = cv2.VideoCapture(0)
        
        # Pinch Calibration
        min_dist = 0.02 # Fingers touching
        max_dist = 0.20 # Fingers spread (adjust based on hand distance from cam)
        smooth_vol = 0  # For smoothing output

        print("Camera Thread Started.")

        while self.is_camera_running and cap.isOpened():
            success, img = cap.read()
            if not success:
                break

            img = cv2.flip(img, 1) # Mirror view
            img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            results = hands.process(img_rgb)

            if results.multi_hand_landmarks:
                for hand_lms in results.multi_hand_landmarks:
                    # Get Thumb Tip (4) and Index Tip (8)
                    x1, y1 = hand_lms.landmark[4].x, hand_lms.landmark[4].y
                    x2, y2 = hand_lms.landmark[8].x, hand_lms.landmark[8].y

                    # Calculate Distance (hypotenuse)
                    length = math.hypot(x2 - x1, y2 - y1)

                    # Map distance to volume range (0 - 100)
                    # Formula: (value - min) / (max - min)
                    vol_val = (length - min_dist) / (max_dist - min_dist)
                    vol_val = max(0.0, min(1.0, vol_val)) # Clamp 0 to 1

                    # Smoothing (Exponential Moving Average) to reduce jitter
                    current_scalar = float(cam_vol_interface.GetMasterVolumeLevelScalar())
                    
                    # Only update if change is significant (deadband) or intentional movement
                    # But for smooth slider, we update with interpolation
                    target_vol = vol_val
                    smooth_vol = (current_scalar * 0.7) + (target_vol * 0.3)

                    # Set Volume
                    try:
                        cam_vol_interface.SetMasterVolumeLevelScalar(smooth_vol, None)
                    except Exception:
                        pass # Ignore intermittent COM errors

                    # --- Visual Feedback on CV2 Window ---
                    h, w, c = img.shape
                    cx1, cy1 = int(x1 * w), int(y1 * h)
                    cx2, cy2 = int(x2 * w), int(y2 * h)
                    
                    # Draw Line
                    cv2.circle(img, (cx1, cy1), 10, (255, 0, 255), cv2.FILLED)
                    cv2.circle(img, (cx2, cy2), 10, (255, 0, 255), cv2.FILLED)
                    cv2.line(img, (cx1, cy1), (cx2, cy2), (255, 0, 255), 3)
                    
                    # Calculate center for text
                    cx, cy = (cx1 + cx2) // 2, (cy1 + cy2) // 2
                    cv2.putText(img, f'{int(smooth_vol*100)}%', (cx, cy - 20), 
                                cv2.FONT_HERSHEY_COMPLEX, 1, (255, 0, 0), 2)

            # Show small preview window
            cv2.imshow("Gesture Control", img)
            if cv2.waitKey(1) & 0xFF == 27: # Esc key
                break

        # Cleanup
        cap.release()
        cv2.destroyAllWindows()
        del cam_vol_interface
        CoUninitialize()
        print("Camera Thread Stopped.")

    # --- UI Loop ---
    def check_mic_status(self, dt):
        if not self.cached_interface:
            try:
                self.cached_interface = _get_new_interface()
            except:
                return

        try:
            vol = self.cached_interface
            current_vol = _scalar_to_percent(float(vol.GetMasterVolumeLevelScalar()))
            current_mute = bool(vol.GetMute())
            self.layout.update_ui(current_vol, current_mute)
        except Exception:
            self.cached_interface = None

if __name__ == "__main__":
    MicApp().run()
