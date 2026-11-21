"""
=============================================================================
Contact Intelligence Engine - Backward Compatibility Shim
=============================================================================
This module redirects all calls to the unified standalone `contact_manager.pyw`.
"""
import importlib.util
import sys
from pathlib import Path

_pyw_path = Path(__file__).parent / "contact_manager.pyw"
if _pyw_path.exists():
    spec = importlib.util.spec_from_file_location("contact_manager", _pyw_path)
    cm = importlib.util.module_from_spec(spec)
    sys.modules["contact_manager"] = cm
    spec.loader.exec_module(cm)

    globals().update({k: getattr(cm, k) for k in dir(cm) if not k.startswith("__")})
