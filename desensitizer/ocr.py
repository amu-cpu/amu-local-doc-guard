from __future__ import annotations

from pathlib import Path


class PaddleOCRProvider:
    """Reserved adapter for a later scanned-PDF/OCR pass."""

    def __init__(self, **options):
        self.options = options

    def detect_text_regions(self, image_path: Path) -> list[dict]:
        raise NotImplementedError("OCR is reserved for a later version.")
