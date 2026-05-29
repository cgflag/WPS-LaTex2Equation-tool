"""Windows drag-and-drop — windnd runs on OS thread; only queue paths here."""

from __future__ import annotations

import queue
import sys

_hooked_hwnds: set[int] = set()


def hook_dropfiles_once(widget, drop_queue: queue.Queue) -> bool:
    """
    Register windnd on a tk/CTk widget.
    Dropped paths are put on drop_queue; process them on the Tk main thread only.
    """
    if sys.platform != "win32":
        return False

    try:
        import windnd  # type: ignore[import-untyped]
    except ImportError:
        return False

    hwnd = widget.winfo_id()
    if hwnd in _hooked_hwnds:
        return False

    def os_thread_callback(paths) -> None:
        # MUST NOT call Tk/Python API here — Windows message thread, no GIL.
        try:
            drop_queue.put_nowait(list(paths))
        except queue.Full:
            pass
        except Exception:
            pass

    try:
        widget.update_idletasks()
        windnd.hook_dropfiles(widget, func=os_thread_callback, force_unicode=True)
        _hooked_hwnds.add(hwnd)
        return True
    except Exception:
        return False
