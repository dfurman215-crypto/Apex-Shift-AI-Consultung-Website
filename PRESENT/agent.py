from __future__ import annotations

import json
import re

OLLAMA_BASE_URL = "http://127.0.0.1:11434"
DEFAULT_MODEL = "gemma3:4b"
OLLAMA_TIMEOUT_SECONDS = 180.0

ALLOWED_SLIDE_TYPES = {
    "hero",
    "metaphor",
    "comparison",
    "content",
    "process",
    "lifecycle",
    "collaboration",
    "roadmap",
    "layers",
    "ecosystem",
    "pricing",
    "benchmark",
    "statement",
}


class LocalAgentUnavailable(RuntimeError):
    """Raised when the local Ollama/Gemma planner cannot be used."""


def _client():
    try:
        from ollama import Client
    except ImportError as exc:
        raise LocalAgentUnavailable(
            "The local Ollama Python client is not installed. "
            "Run: python -m pip install -r requirements.txt"
        ) from exc

    return Client(host=OLLAMA_BASE_URL, timeout=OLLAMA_TIMEOUT_SECONDS)


def ollama_available(model: str = DEFAULT_MODEL) -> bool:
    try:
        _client().show(model)
        return True
    except Exception:
        return False


def _extract_slide_blueprint(markdown_text: str) -> str:
    """Keep Gemma focused on the numbered slide blueprint instead of the entire brief."""
    text = markdown_text.replace("\x0b", "\n").replace("\r\n", "\n").replace("\r", "\n")
    match = re.search(
        r"# Slide Blueprint\s*(.*?)(?:# Presentation Style Guide|\Z)",
        text,
        flags=re.S | re.I,
    )
    blueprint = match.group(1).strip() if match else text.strip()
    return blueprint[:18000]


def _deck_summary(deck_info: dict) -> str:
    """Create a compact text-only summary of the existing deck."""
    lines = []
    for slide in deck_info.get("slides", [])[:20]:
        text = str(slide.get("text", "") or "").strip()
        if "# Apex Shift PEARL Playbook" in text or "# Slide Blueprint" in text:
            text = "[embedded presentation brief]"
        text = re.sub(r"\s+", " ", text)
        if len(text) > 300:
            text = text[:300] + "..."
        lines.append(
            f"Slide {slide.get('number')}: "
            f"{text or '[visual / little extracted text]'}"
        )
    return "\n".join(lines)


def _as_list(value, limit: int = 8) -> list:
    if value is None:
        return []
    if isinstance(value, list):
        return value[:limit]
    return [value]


def _validate_spec(spec: dict) -> dict:
    if not isinstance(spec, dict):
        raise ValueError("Gemma did not return a JSON object.")

    raw_slides = spec.get("slides")
    if not isinstance(raw_slides, list) or not raw_slides:
        raise ValueError("Gemma returned no slides.")

    slides = []
    for index, raw in enumerate(raw_slides[:16], start=1):
        if not isinstance(raw, dict):
            continue

        slide = dict(raw)
        slide_type = str(slide.get("type", "content")).strip().lower()
        if slide_type not in ALLOWED_SLIDE_TYPES:
            slide_type = "content"

        slide["type"] = slide_type
        slide["headline"] = str(slide.get("headline") or f"Slide {index}")

        for key in (
            "items",
            "left",
            "right",
            "left_items",
            "right_items",
            "strategic",
            "tactical",
            "layers",
            "nodes",
            "factors",
            "metrics",
            "stages",
            "formula",
        ):
            if key in slide:
                slide[key] = _as_list(slide[key], 8)

        if slide_type == "comparison":
            slide.setdefault("left_title", "Current")
            slide.setdefault("right_title", "Target")
            slide.setdefault("left", [])
            slide.setdefault("right", [])
        elif slide_type == "content":
            slide.setdefault("items", [])
        elif slide_type in {"process", "lifecycle"}:
            slide.setdefault("stages", slide.get("items", []))
        elif slide_type == "collaboration":
            slide.setdefault("left_title", "Apex Shift")
            slide.setdefault("left_items", [])
            slide.setdefault("center", "Collaboration")
            slide.setdefault("right_title", "Client")
            slide.setdefault("right_items", [])
        elif slide_type == "roadmap":
            slide.setdefault("strategic_title", "Strategic Roadmap")
            slide.setdefault("strategic", [])
            slide.setdefault("tactical_title", "Tactical Roadmap")
            slide.setdefault("tactical", [])
        elif slide_type == "layers":
            slide.setdefault("center", "Pilot")
            slide.setdefault("layers", slide.get("items", []))
        elif slide_type == "ecosystem":
            slide.setdefault("center", "Integrated System")
            slide.setdefault("nodes", slide.get("items", []))
        elif slide_type == "pricing":
            slide.setdefault("formula", ["Scope", "Effort", "Price"])
            slide.setdefault("factors", slide.get("items", []))
        elif slide_type == "benchmark":
            slide.setdefault("metrics", [])

        slides.append(slide)

    if not slides:
        raise ValueError("Gemma returned no usable slide definitions.")

    return {
        "deck": str(spec.get("deck") or "PRESENT Fast Plan"),
        "version": "0.3",
        "slides": slides,
    }


def plan_deck_with_gemma(
    deck_info: dict,
    markdown_text: str,
    model: str = DEFAULT_MODEL,
) -> dict:
    """Use local Gemma for one focused text-only planning pass."""
    if not ollama_available(model):
        raise LocalAgentUnavailable(
            f"Local Ollama model '{model}' is not available."
        )

    blueprint = _extract_slide_blueprint(markdown_text)
    deck_summary = _deck_summary(deck_info)

    prompt = f"""
You are PRESENT's local slide-planning specialist.

TASK
Create a concise JSON slide plan from the numbered Slide Blueprint below.
Do not design the PowerPoint file itself. Python handles layout, shapes, images, and rendering.

GOAL
Produce a useful executive rough draft quickly.

RULES
1. Return JSON only.
2. Prefer one slide per numbered blueprint section.
3. Keep the blueprint's narrative order.
4. Use the most appropriate type from:
   hero, metaphor, comparison, content, process, lifecycle,
   collaboration, roadmap, layers, ecosystem, pricing, benchmark, statement.
5. Keep wording concise. Avoid paragraphs.
6. Never invent facts, prices, rates, statistics, citations, or claims.
7. PEARL Methodology describes how Apex Shift works.
8. Capability Lifecycle describes what is progressively built.
9. Keep those two concepts distinct.
10. Do not choose or analyze images. Python will reuse source-deck images separately.

JSON SHAPE
{{
  "deck": "PEARL",
  "slides": [
    {{
      "type": "hero",
      "headline": "...",
      "subheadline": "..."
    }}
  ]
}}

TYPE-SPECIFIC FIELDS
comparison: left_title, left[], right_title, right[]
content: items[]
process/lifecycle: stages[] using short strings or {{"label":"...","detail":"..."}}
collaboration: left_title, left_items[], center, center_detail, right_title, right_items[]
roadmap: strategic_title, strategic[], tactical_title, tactical[], cadence
layers: center, layers[]
ecosystem: center, nodes[]
pricing: formula[], factors[], payment
benchmark: metrics[] using {{"label":"...","value":"...","detail":"..."}}

EXISTING DECK SUMMARY
{deck_summary}

NUMBERED SLIDE BLUEPRINT
{blueprint}
""".strip()

    try:
        response = _client().chat(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            format="json",
            options={
                "temperature": 0.1,
                "num_predict": 2600,
            },
        )
    except Exception as exc:
        raise LocalAgentUnavailable(
            f"Local Gemma planning call failed or exceeded "
            f"{int(OLLAMA_TIMEOUT_SECONDS)} seconds: {exc}"
        ) from exc

    raw = str(response.message.content or "").strip()
    if not raw:
        raise ValueError("Gemma returned an empty response.")

    try:
        spec = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Gemma returned invalid JSON: {exc}") from exc

    return _validate_spec(spec)
