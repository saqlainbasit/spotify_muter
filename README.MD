# SpotifyAdMuter

Automatically mutes Spotify ads and unmutes when your song resumes. Runs silently in the system tray.

## Features

- Detects ads by reading the Spotify window title
- Mutes system volume during ads, unmutes when music resumes
- System tray icon — green when playing, red when muted
- Voice notifications via text-to-speech
- Lightweight, runs in the background

## Usage

### Run from source

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run:
   ```bash
   python spotify_ad_muter.py
   ```

### Run the executable

Just double-click `SpotifyAdMuter.exe`. No Python needed.

> Make sure Spotify is open before launching the app.
