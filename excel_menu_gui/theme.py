# Backward-compat shim for app.gui.theme
import sys as _sys
from app.gui import theme as _mod
_sys.modules[__name__] = _mod
