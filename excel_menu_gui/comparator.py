# Backward-compat shim for app.services.comparator
import sys as _sys
from app.services import comparator as _mod
_sys.modules[__name__] = _mod
