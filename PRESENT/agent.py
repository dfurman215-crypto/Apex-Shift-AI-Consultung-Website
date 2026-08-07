from __future__ import annotations

import json
import re

OLLAMA_BASE_URL = "http://127.0.0.1:11434"
DEFAULT_MODEL = "gemma3:4b"
OLLAMA_TIMEOUT_SECONDS = 180.0
ALLOWED_SLIDE_TYPES = {
    "hero", "metaphor", "comparison", "content", "process", "lifecycle",
    "collaboration", "roadmap", "layers", "ecosystem", "pricing",
    "benchmark", "statement",
}


class LocalAgent