# Backward-compat shim for app.services.menu_template_filler
import sys as _sys
from app.services import menu_template_filler as _mod
_sys.modules[__name__] = _mod
