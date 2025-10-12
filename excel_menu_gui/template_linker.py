# Backward-compat shim for app.services.template_linker
import sys as _sys
from app.services import template_linker as _mod
_sys.modules[__name__] = _mod
