"""Dedicated single worker thread for COM operations with timeout & poisoned state."""

from __future__ import annotations

import os
import queue
import sys
import threading
from concurrent.futures import Future
from typing import Any, Callable, Optional, Tuple

from src.core.exceptions import ERROR_MESSAGES, LiveWordError

# ---------------------------------------------------------------------------
# Guarded COM Imports (C2: pywin32 is optional)
# ---------------------------------------------------------------------------
HAS_COM = False
try:
    if sys.platform == "win32":
        import pythoncom
        import win32api
        import win32com.client
        HAS_COM = True
except ImportError:
    HAS_COM = False


class ComWorker:
    """Dedicated single worker thread for COM operations with timeout & poisoned state."""

    def __init__(self, default_timeout: float = 20.0):
        self._default_timeout = float(os.getenv("LIVE_CALL_TIMEOUT", default_timeout))
        self._queue: queue.Queue[Optional[Tuple[Callable[..., Any], tuple, dict, Future]]] = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._poisoned = False
        self._lock = threading.Lock()
        self._running = False

    def start(self) -> None:
        """Start worker thread if not already active."""
        with self._lock:
            if self._thread is None or not self._thread.is_alive():
                self._running = True
                self._thread = threading.Thread(
                    target=self._worker_loop,
                    name="LiveWordCOMWorker",
                    daemon=True,
                )
                self._thread.start()

    def _worker_loop(self) -> None:
        """Thread worker loop running CoInitialize/CoUninitialize."""
        if HAS_COM and sys.platform == "win32":
            try:
                import win32service
                import win32con
                h_default = win32service.OpenDesktop("Default", 0, False, win32con.MAXIMUM_ALLOWED)
                h_default.SetThreadDesktop()
            except Exception:
                pass
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

        try:
            while self._running:
                item = self._queue.get()
                if item is None:
                    self._queue.task_done()
                    break

                func, args, kwargs, future = item
                try:
                    result = func(*args, **kwargs)
                    if not future.done():
                        future.set_result(result)
                except BaseException as exc:
                    if not future.done():
                        future.set_exception(exc)
                finally:
                    self._queue.task_done()
                    with self._lock:
                        # If a prior call timed out, its completion clears the poisoned state
                        if self._poisoned:
                            self._poisoned = False
        finally:
            if HAS_COM:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def run(self, func: Callable[..., Any], *args: Any, timeout: Optional[float] = None, **kwargs: Any) -> Any:
        """Submit a callable to the COM worker thread and await its result within timeout."""
        if self._thread is not None and threading.current_thread() is self._thread:
            return func(*args, **kwargs)

        self.start()
        effective_timeout = timeout if timeout is not None else self._default_timeout

        with self._lock:
            if self._poisoned:
                raise LiveWordError(ERROR_MESSAGES["worker_poisoned"])

        fut: Future = Future()
        self._queue.put((func, args, kwargs, fut))

        try:
            return fut.result(timeout=effective_timeout)
        except TimeoutError:
            with self._lock:
                self._poisoned = True
            msg = ERROR_MESSAGES["modal_dialog_or_timeout"].format(timeout=effective_timeout)
            raise LiveWordError(msg) from None

    def is_poisoned(self) -> bool:
        """Check if worker is currently in poisoned state."""
        with self._lock:
            return self._poisoned

    def clear_poison(self) -> None:
        """Manual reset of poisoned flag if needed."""
        with self._lock:
            self._poisoned = False

    def stop(self) -> None:
        """Stop the worker thread cleanly."""
        with self._lock:
            self._running = False
        self._queue.put(None)
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=2.0)
        self._thread = None


# Global singleton COM worker instance
com_worker = ComWorker()

