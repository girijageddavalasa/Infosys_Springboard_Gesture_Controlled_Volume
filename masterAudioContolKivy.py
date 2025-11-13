import kivy
kivy.require('1.11.1')

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.slider import Slider
from kivy.uix.label import Label
from kivy.clock import Clock
from kivy.properties import NumericProperty, BooleanProperty, StringProperty
from kivy.core.window import Window

from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import comtypes

# Define the volume step for key presses (5%)
VOLUME_STEP = 0.05


# --- pycaw Backend Functions ---

def get_volume_interface():
    """Gets the main audio endpoint volume interface."""
    try:
        comtypes.CoInitialize()
        devices = AudioUtilities.GetSpeakers()
        return devices.EndpointVolume
    except Exception as e:
        print(f"Error accessing audio device: {e}")

        # Fallback mock object for testing without real audio device
        class MockVolume:
            def GetMasterVolumeLevelScalar(self): return 0.5
            def SetMasterVolumeLevelScalar(self, level, guid): pass
            def GetMute(self): return 0
            def SetMute(self, mute, guid): pass

        return MockVolume()


VOLUME_INTERFACE = get_volume_interface()


def get_current_volume_scalar():
    try:
        return VOLUME_INTERFACE.GetMasterVolumeLevelScalar()
    except Exception:
        return 0.0


def set_volume_scalar(level_scalar):
    try:
        VOLUME_INTERFACE.SetMasterVolumeLevelScalar(level_scalar, None)
    except Exception as e:
        print(f"Error setting volume: {e}")


def get_mute_status():
    try:
        return VOLUME_INTERFACE.GetMute() == 1
    except Exception:
        return False


def set_mute_status(is_muted):
    try:
        VOLUME_INTERFACE.SetMute(1 if is_muted else 0, None)
    except Exception as e:
        print(f"Error setting mute: {e}")


# --- Kivy Application ---

class VolumeControlWidget(BoxLayout):
    """Root widget for the volume control application."""

    volume_scalar = NumericProperty(get_current_volume_scalar())
    is_muted = BooleanProperty(get_mute_status())
    volume_label_text = StringProperty(f"Volume: {get_current_volume_scalar() * 100:.0f}%")
    status_text = StringProperty("Use Up/Down Arrows to adjust volume. Press 'M' to mute.")

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = 20
        self.spacing = 15

        # --- 1. Volume Label ---
        self.volume_label = Label(text=self.volume_label_text, size_hint_y=0.2, font_size='24sp')
        self.bind(volume_label_text=self.volume_label.setter('text'))  # <-- FIXED: keeps label updated
        self.add_widget(self.volume_label)

        # --- 2. Volume Slider ---
        self.volume_slider = Slider(
            min=0, max=1, value=self.volume_scalar, step=0.01,
            size_hint_y=0.4
        )
        self.volume_slider.bind(value=self.update_volume_from_gui)
        self.add_widget(self.volume_slider)

        # Keep slider synced with property (fix for external change updates)
        self.bind(volume_scalar=self.volume_slider.setter('value'))

        # --- 3. Mute/Unmute Button ---
        self.mute_button = Button(
            text='Mute (M)', size_hint_y=0.2,
            on_press=self.toggle_mute
        )
        self.add_widget(self.mute_button)

        # --- 4. Status Bar ---
        self.status_bar = Label(text=self.status_text, size_hint_y=0.1, font_size='12sp')
        self.bind(status_text=self.status_bar.setter('text'))
        self.add_widget(self.status_bar)

        # --- 5. Kivy Property Bindings ---
        self.bind(volume_scalar=self.update_label)
        self.bind(is_muted=self.update_gui_on_mute)

        # --- 6. Keyboard Binding ---
        Window.bind(on_key_down=self.key_press)

        # --- 7. Auto-Update Check ---
        Clock.schedule_interval(self.check_external_change, 0.25)

    # --- Volume + Mute Methods ---

    def update_volume_from_gui(self, instance, value):
        set_volume_scalar(value)
        self.volume_scalar = value
        self.status_text = f"Volume set via Slider: {value * 100:.0f}%"

    def update_volume_from_key(self, change_scalar):
        current_scalar = self.volume_scalar
        new_scalar = max(0.0, min(1.0, current_scalar + change_scalar))
        set_volume_scalar(new_scalar)
        self.volume_scalar = new_scalar
        self.status_text = f"Volume set via Keyboard: {new_scalar * 100:.0f}%"

    def toggle_mute(self, instance=None):
        new_mute_status = not self.is_muted
        set_mute_status(new_mute_status)
        self.is_muted = new_mute_status
        self.status_text = f"Mute Toggled: {'Muted' if new_mute_status else 'Unmuted'}"

    def key_press(self, window, keycode, scancode, codepoint, modifier):
        if keycode == 273:  # Up Arrow
            self.update_volume_from_key(VOLUME_STEP)
            return True
        elif keycode == 274:  # Down Arrow
            self.update_volume_from_key(-VOLUME_STEP)
            return True
        elif codepoint and codepoint.lower() == 'm':  # M key
            self.toggle_mute()
            return True
        return False

    # --- Reactive Updates ---

    def update_label(self, instance, value):
        self.volume_label_text = f'Volume: {value * 100:.0f}%'

    def update_gui_on_mute(self, instance, is_muted):
        if is_muted:
            self.mute_button.text = 'Unmute (Muted)'
            self.mute_button.background_color = (0.8, 0.2, 0.2, 1)
        else:
            self.mute_button.text = 'Mute (M)'
            self.mute_button.background_color = (0.2, 0.8, 0.2, 1)

    def check_external_change(self, dt):
        current_external_scalar = get_current_volume_scalar()
        current_external_mute = get_mute_status()

        if abs(self.volume_scalar - current_external_scalar) > 0.005:
            self.volume_scalar = current_external_scalar
            self.status_text = "Volume changed externally!"

        if self.is_muted != current_external_mute:
            self.is_muted = current_external_mute


# --- App Class ---

class VolumeApp(App):
    def build(self):
        self.title = 'System Volume Controller (Kivy)'
        return VolumeControlWidget()


if __name__ == '__main__':
    try:
        VolumeApp().run()
    finally:
        comtypes.CoUninitialize()
