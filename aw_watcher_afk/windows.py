import ctypes
import time
from ctypes import POINTER, WINFUNCTYPE, Structure  # type: ignore
from ctypes.wintypes import BOOL, DWORD, UINT

from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioMeterInformation


class LastInputInfo(Structure):
    _fields_ = [("cbSize", UINT), ("dwTime", DWORD)]


def _getLastInputTick() -> int:
    prototype = WINFUNCTYPE(BOOL, POINTER(LastInputInfo))
    paramflags = ((1, "lastinputinfo"),)
    c_GetLastInputInfo = prototype(("GetLastInputInfo", ctypes.windll.user32), paramflags)  # type: ignore

    lastinput = LastInputInfo()
    lastinput.cbSize = ctypes.sizeof(LastInputInfo)
    assert 0 != c_GetLastInputInfo(lastinput)
    return lastinput.dwTime


def _getTickCount() -> int:
    prototype = WINFUNCTYPE(DWORD)
    paramflags = ()
    c_GetTickCount = prototype(("GetTickCount", ctypes.windll.kernel32), paramflags)  # type: ignore
    return c_GetTickCount()


def seconds_since_last_input():
    seconds_since_input = (_getTickCount() - _getLastInputTick()) / 1000
    return seconds_since_input


def is_audio_playing():
    """Check if any audio is currently playing on the system."""
    try:
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process:
                volume = session._ctl.QueryInterface(IAudioMeterInformation)
                if volume.GetPeakValue() > 0.01:
                    return True
    except Exception:
        pass
    return False


def seconds_since_last_input_with_audio():
    """
    Returns seconds since last input, but returns 0 if audio is playing.
    This considers the user as "active" when audio is playing.
    """
    if is_audio_playing():
        return 0.0
    return seconds_since_last_input()


if __name__ == "__main__":
    while True:
        time.sleep(1)
        seconds_since_last_input_with_audio()
