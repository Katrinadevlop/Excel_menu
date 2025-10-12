# Backward-compat shim for app.reports.presentation_handler
import sys as _sys
from app.reports import presentation_handler as _mod
_sys.modules[__name__] = _mod
