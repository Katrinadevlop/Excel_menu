# Backward-compat shim for app.gui.ui_styles
import sys as _sys
from app.gui import ui_styles as _mod
_sys.modules[__name__] = _mod
