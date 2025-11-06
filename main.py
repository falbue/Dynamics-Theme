"""Dynamics Theme - Automatic Windows theme switcher based on sunrise/sunset times."""

from typing import Optional, Tuple
from datetime import datetime, timezone
import ctypes
import sys
import threading
import time

import ephem
import pystray
import requests
import winreg
from PIL import Image

APP_NAME = 'Dynamics Theme'
VERSION = '1.3'

TRANSLATIONS = {
    'en': {
        'dark': 'Dark ☾',
        'light': 'Light ☼',
        'automatic': 'Automatic',
        'exit': 'Exit',
    },
    'ru': {
        'dark': 'Тёмная ☾',
        'light': 'Светлая ☼',
        'automatic': 'Автоматическая',
        'exit': 'Закрыть',
    },
    'es': {
        'dark': 'Oscuro ☾',
        'light': 'Claro ☼',
        'automatic': 'Automático',
        'exit': 'Salir',
    },
    'de': {
        'dark': 'Dunkel ☾',
        'light': 'Hell ☼',
        'automatic': 'Automatisch',
        'exit': 'Beenden',
    },
    'fr': {
        'dark': 'Sombre ☾',
        'light': 'Clair ☼',
        'automatic': 'Automatique',
        'exit': 'Quitter',
    },
    'it': {
        'dark': 'Scuro ☾',
        'light': 'Chiaro ☼',
        'automatic': 'Automatico',
        'exit': 'Esci',
    },
    'pt': {
        'dark': 'Escuro ☾',
        'light': 'Claro ☼',
        'automatic': 'Automático',
        'exit': 'Sair',
    },
    'zh': {
        'dark': '深色 ☾',
        'light': '浅色 ☼',
        'automatic': '自动',
        'exit': '退出',
    },
    'ja': {
        'dark': 'ダーク ☾',
        'light': 'ライト ☼',
        'automatic': '自動',
        'exit': '終了',
    },
    'ko': {
        'dark': '어두운 ☾',
        'light': '밝은 ☼',
        'automatic': '자동',
        'exit': '종료',
    },
}

stop_event: threading.Event = threading.Event()
icon: Optional[pystray.Icon] = None


def set_windows_theme(theme: str) -> bool:
    """Set Windows theme to light or dark mode and update tray icon.

    Args:
        theme: Theme to set ('light' or 'dark')

    Returns:
        True if theme was set successfully, False otherwise
    """
    global icon

    try:
        registry_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path, 0, winreg.KEY_WRITE) as key:
            if theme == "light":
                theme_value = 1
                icon_path = "lib/icon_light.png"
            elif theme == "dark":
                theme_value = 0
                icon_path = "lib/icon_dark.png"
            else:
                print(f"Invalid theme: '{theme}'. Choose 'light' or 'dark'.")
                return False

            winreg.SetValueEx(key, "AppsUseLightTheme", 0, winreg.REG_DWORD, theme_value)
            winreg.SetValueEx(key, "SystemUsesLightTheme", 0, winreg.REG_DWORD, theme_value)

            if icon:
                icon.icon = Image.open(icon_path)

        ctypes.windll.user32.SendMessageW(0xFFFF, 0x001A, 0, 0)
        print(f"Theme successfully changed to '{theme}'.")
        return True
    except Exception as e:
        print(f"Error setting theme: {e}")
        return False


def get_system_language() -> str:
    """Detect system language and return ISO 639 language code.

    Returns:
        Two-letter language code (e.g., 'en', 'ru', 'es')
    """
    try:
        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        GetUserDefaultUILanguage = kernel32.GetUserDefaultUILanguage
        GetUserDefaultUILanguage.restype = ctypes.c_uint
        lcid = GetUserDefaultUILanguage()

        LOCALE_NAME_MAX_LENGTH = 85
        LOCALE_SISO639LANGNAME = 0x59

        locale_name = ctypes.create_unicode_buffer(LOCALE_NAME_MAX_LENGTH)

        if kernel32.GetLocaleInfoW(lcid, LOCALE_SISO639LANGNAME, locale_name, LOCALE_NAME_MAX_LENGTH):
            return locale_name.value[:2].lower()

        return 'en'
    except Exception as e:
        print(f"Error detecting system language: {e}")
        return 'en'


def get_current_theme() -> Optional[int]:
    """Get current Windows theme setting.

    Returns:
        1 if light theme is active, 0 if dark theme is active, None on error
    """
    try:
        registry_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path, 0, winreg.KEY_READ) as key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0]
    except Exception as e:
        print(f"Error getting current theme: {e}")
        return None


def get_location() -> Tuple[Optional[float], Optional[float]]:
    """Get current geographic location using IP geolocation.

    Returns:
        Tuple of (latitude, longitude) or (None, None) on error
    """
    try:
        response = requests.get('https://ipinfo.io/json', timeout=10)
        response.raise_for_status()
        data = response.json()
        latitude, longitude = map(float, data['loc'].split(','))
        return latitude, longitude
    except Exception as e:
        print(f"Error getting location: {e}")
        return None, None


def get_sunrise_and_sunset(latitude: float, longitude: float) -> Tuple[Optional[datetime], Optional[datetime]]:
    """Calculate sunrise and sunset times for given coordinates.

    Args:
        latitude: Geographic latitude
        longitude: Geographic longitude

    Returns:
        Tuple of (sunrise_time_utc, sunset_time_utc) or (None, None) on error
    """
    try:
        observer = ephem.Observer()
        observer.lat = str(latitude)
        observer.lon = str(longitude)
        sunrise_time_utc = observer.next_rising(ephem.Sun()).datetime()
        sunset_time_utc = observer.next_setting(ephem.Sun()).datetime()
        return sunrise_time_utc, sunset_time_utc
    except Exception as e:
        print(f"Error calculating sunrise/sunset: {e}")
        return None, None


def sun_time_local(sunrise_datetime_utc: datetime, sunset_datetime_utc: datetime) -> Optional[Tuple[str, str]]:
    """Convert UTC sunrise/sunset times to local time strings.

    Args:
        sunrise_datetime_utc: Sunrise time in UTC
        sunset_datetime_utc: Sunset time in UTC

    Returns:
        Tuple of (sunrise_local, sunset_local) formatted as 'HH:MM:SS', or None on error
    """
    try:
        local_offset = datetime.now(timezone.utc).astimezone().utcoffset()
        if local_offset is None:
            return None
        sunrise_time_local = sunrise_datetime_utc + local_offset
        sunset_time_local = sunset_datetime_utc + local_offset
        return sunrise_time_local.strftime('%H:%M:%S'), sunset_time_local.strftime('%H:%M:%S')
    except Exception as e:
        print(f"Error converting to local time: {e}")
        return None


def automatic_data() -> Optional[Tuple[str, str]]:
    """Get sunrise and sunset times in local timezone.

    Returns:
        Tuple of (sunrise, sunset) times formatted as 'HH:MM:SS', or None on error
    """
    latitude, longitude = get_location()
    if latitude is None or longitude is None:
        return None

    sunrise_datetime_utc, sunset_datetime_utc = get_sunrise_and_sunset(latitude, longitude)
    if sunrise_datetime_utc is None or sunset_datetime_utc is None:
        return None

    return sun_time_local(sunrise_datetime_utc, sunset_datetime_utc)


def get_local_time() -> str:
    """Get current local time formatted as 'HH:MM:SS'.

    Returns:
        Current local time string
    """
    return datetime.now().strftime("%H:%M:%S")


def select_theme(theme: str) -> None:
    """Select and apply a theme, starting automatic mode if requested.

    Args:
        theme: Theme to apply ('auto', 'light', or 'dark')
    """
    stop_event.set()
    if theme == 'auto':
        start_automatic()
    else:
        print("Automatic mode disabled")
        set_windows_theme(theme)


def start_automatic() -> None:
    """Start automatic theme switching in a background thread."""
    print("Automatic mode enabled")
    global stop_event
    stop_event = threading.Event()
    thread = threading.Thread(target=automatic_theme, daemon=True)
    thread.start()


def automatic_theme() -> None:
    """Monitor time and automatically switch theme based on sunrise/sunset."""
    backoff = 30  # initial backoff in seconds
    max_backoff = 600  # maximum backoff in seconds (10 minutes)
    sun_times = automatic_data()
    
    while sun_times is None and not stop_event.is_set():
        print(f"Failed to get data for automatic mode. Retrying in {backoff} seconds...")
        time.sleep(backoff)
        backoff = min(backoff * 2, max_backoff)
        sun_times = automatic_data()
    
    if sun_times is None:
        print("Could not fetch sunrise/sunset data after multiple attempts. Exiting automatic mode.")
        return

    sunrise, sunset = sun_times
    current_applied_theme = None

    while not stop_event.is_set():
        local_time = get_local_time()
        print(f"Sunrise: {sunrise} | Current: {local_time} | Sunset: {sunset}")

        desired_theme = "light" if sunrise < local_time < sunset else "dark"
        
        # Only set theme if it needs to change
        if desired_theme != current_applied_theme:
            if set_windows_theme(desired_theme):
                current_applied_theme = desired_theme

        time.sleep(60)


def get_translations(language: str) -> dict:
    """Get translations for a specific language, with fallback to English.

    Args:
        language: Two-letter language code

    Returns:
        Dictionary of translations
    """
    return TRANSLATIONS.get(language, TRANSLATIONS['en'])


def create_tray_icon() -> None:
    """Create and run the system tray icon with localized menu."""
    global icon

    current_theme = get_current_theme()
    if current_theme is None:
        return

    icon_path = f"lib/icon_{'light' if current_theme else 'dark'}.png"
    icon = pystray.Icon("dynamics_theme", Image.open(icon_path), APP_NAME)

    language = get_system_language()
    translations = get_translations(language)

    menu_items = [
        pystray.MenuItem(translations['dark'], lambda: select_theme('dark')),
        pystray.MenuItem(translations['light'], lambda: select_theme('light')),
        pystray.MenuItem(translations['automatic'], lambda: select_theme('auto')),
        pystray.MenuItem(translations['exit'], hide_icon)
    ]

    icon.menu = pystray.Menu(*menu_items)
    start_automatic()
    icon.run()


def hide_icon() -> None:
    """Stop the tray icon and exit the application."""
    global icon
    stop_event.set()
    if icon:
        icon.stop()
    sys.exit(0)


if __name__ == "__main__":
    create_tray_icon()

