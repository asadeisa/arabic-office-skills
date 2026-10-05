import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for skill in ("arabic-docx", "arabic-pptx", "arabic-pdf"):
    sys.path.insert(0, str(ROOT / skill / "scripts"))
