"""Application entry point for local run and PyInstaller packaging."""

import tkinter  # noqa: F401 - ensure PyInstaller includes Tcl/Tk runtime hooks.

from src.main_ui import main


if __name__ == "__main__":
    raise SystemExit(main())
