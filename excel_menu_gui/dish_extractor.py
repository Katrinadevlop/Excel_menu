# Backward-compat shim for app.services.dish_extractor
import sys as _sys
from app.services import dish_extractor as _mod
_sys.modules[__name__] = _mod
