import time
from threading import Thread, Event
import win32gui
import win32con

from app_log import error

FILE_IN_USE = "File in Use"

# A helper class to close VB error dialog
class MsgBoxListener(Thread):

    def __init__(self, interval: int):
        Thread.__init__(self)
        self._titles = ["Microsoft Visual Basic", FILE_IN_USE]
        self._interval = interval
        self._stop_event = Event()

        self.file_in_use_dialog = False

    def stop(self): self._stop_event.set()

    @property
    def is_running(self): return not self._stop_event.is_set()

    def run(self):
        while self.is_running:
            try:
                time.sleep(self._interval)
                self._close_msgbox()
            except Exception as e:
                error(e)
                print(e, flush=True)


    def _close_msgbox(self):
        # find the top window by title
        hwnd = None
        found_title = None

        for title in self._titles:
            hwnd = win32gui.FindWindow(None, title)

            if hwnd:
                found_title = title
                break

        if not hwnd: return

        # find child button
        h_btn = win32gui.FindWindowEx(hwnd, None, 'Button', None)
        h_btn = win32gui.FindWindowEx(hwnd, h_btn, 'Button', None)

        if found_title == FILE_IN_USE:
            self.file_in_use_dialog = True
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

        if not h_btn: return

        # show text
        text = win32gui.GetWindowText(h_btn)

        # click button
        win32gui.PostMessage(h_btn, win32con.WM_LBUTTONDOWN, None, None)
        time.sleep(0.2)
        win32gui.PostMessage(h_btn, win32con.WM_LBUTTONUP, None, None)
        time.sleep(0.2)