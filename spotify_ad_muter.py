import time
import threading
import sys
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from PIL import Image, ImageDraw
import pystray
import pyttsx3
import win32gui
import win32process
import psutil

# ================================
# Volume Control Setup (Windows)
# ================================
devices = AudioUtilities.GetSpeakers()
volume_interface = devices._dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(volume_interface, POINTER(IAudioEndpointVolume))

def mute():
    volume.SetMute(1, None)

def unmute():
    volume.SetMute(0, None)

# ================================
# Text-to-Speech Setup
# ================================
tts = pyttsx3.init()
tts.setProperty('rate', 170)
tts.setProperty('volume', 1)

def speak(message):
    print("🗣️ Tars:", message)
    tts.say(message)
    tts.runAndWait()

# ================================
# Tray Icon
# ================================
def create_icon(color):
    img = Image.new('RGB', (64, 64), color=color)
    draw = ImageDraw.Draw(img)
    draw.rectangle((16, 16, 48, 48), fill="white")
    return img

icon = pystray.Icon("SpotifyAdMuter")
icon.icon = create_icon("green")
icon.title = "Spotify Ad Muter"

# ================================
# Spotify Window Title Detection
# ================================
def get_spotify_title():
    titles = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                proc = psutil.Process(pid)
                if 'spotify' in proc.name().lower() and title:
                    titles.append(title)
            except:
                pass
    win32gui.EnumWindows(callback, None)
    return titles[0] if titles else ""

def watch_spotify():
    is_muted = False
    speak("Hey User, I'm Tars, your Spotify assistant is now online.")

    while True:
        title = get_spotify_title()

        # Song titles always have " - " (Artist - Song)
        # Ads show "Advertisement" or just "Spotify"
        is_ad = (title in ("Spotify", "Advertisement", "") or " - " not in title)

        if is_ad and not is_muted:
            mute()
            is_muted = True
            print(f"🔇 Ad detected (title: '{title}') → Muting")
            icon.icon = create_icon("red")
            speak("Hey User, ad is playing, muting now.")

        elif not is_ad and is_muted:
            unmute()
            is_muted = False
            print(f"🔊 Song resumed: {title}")
            icon.icon = create_icon("green")
            speak(f"Song is back. Now playing {title}.")

        else:
            print(f"▶ {title}" if title else "⏸ Spotify not running")

        time.sleep(2)

# ================================
# Tray Menu
# ================================
def quit_app(icon, item):
    speak("Goodbye User, Tars is shutting down.")
    icon.stop()
    sys.exit()

icon.menu = pystray.Menu(pystray.MenuItem("Quit", quit_app))

thread = threading.Thread(target=watch_spotify, daemon=True)
thread.start()

icon.run()