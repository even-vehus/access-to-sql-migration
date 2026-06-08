"""Pytest setup: the migration scripts live in sibling folders (not a package),
so put each on sys.path to allow `import extract_access_db`, `import
inspect_artifacts`, etc."""

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent

for _sub in ("migration", "extract", "generators", "deploy"):
    _path = str(_ROOT / _sub)
    if _path not in sys.path:
        sys.path.insert(0, _path)
