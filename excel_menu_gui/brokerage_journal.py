# Backward-compat shim for app.reports.brokerage_journal
import sys as _sys
from app.reports import brokerage_journal as _mod
_sys.modules[__name__] = _mod
