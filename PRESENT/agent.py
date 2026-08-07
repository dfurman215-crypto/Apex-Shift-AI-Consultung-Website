from __future__ import annotations

import json
from pathlib import Path

OLLAMA_BASE_URL = "http://127.0.0.1:11434"
DEFAULT_MODEL = "gemma3:4b"
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
    pass


def _client():
    try:
        from ollama import Client
    except ImportError as exc:
        raise LocalAgentUnavailable(
            "The local Ollama Python client is not installed. Run: python -m pip install -r requirements.txt"
        ) from exc
    return Client(host=OLLAMA_BASE_URL)


def ollama_available(model: str = DEFAULT_MODEL) -> bool:
    try:
        client = _client()
        client.show(model)
        return True
    except Exception:
        return False


def _deck_context(deck_info: dict) -> str:
    lines = []
    for slide in deck_info.get("slides", []):
        text = slide.get("text", "").strip()
        if "# Apex Shift PEARL Playbook" in text or "# Slide Blueprint" in text:
            text = "[Embedded presentation brief omitted here; supplied separately.]"
        if len(text) > 1000:
            text = text[:1000] + "..."
        lines.append(
            f"Slide {slide.get('number')}: {text or '[visual/no extracted text]'} "
            f"(shapes: {slide.get('shape_count', '?')})"
        )
    return "\n\n".join(lines)


def _asset_context(assets: list[dict] | None) -> str:
    if not assets:
        return "No reusable picture assets were extracted from the source deck."
    lines = [
        "Reusable picture assets are shown in the attached contact sheet. "
        "Each tile is labeled with its asset ID and source slide."
    ]
    for asset in assets:
        dims = ""
        if asset.get("width") and asset.get("height"):
            dims = f" {asset['width']}x{asset['height']}px"
        lines.append(f"{asset['id']}: source slide {asset['slide']},{dims} file={asset['filename']}")
    return "\n".join(lines)


def _as_list(value, limit=8):
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if not isinstance(value, list):
        return [str(value)]
    return value[:limit]


def _validate_spec(spec: dict) -> dict:
    if not isinstance(spec, dict):
        raise ValueError("Gemma did not return a JSON object.")
    slides = spec.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ValueError("Gemma returned no slides.")

    cleaned = []
    for index, raw_slide in enumerate(slides, start=1):
        if not isinstance(raw_slide, dict):
            raise ValueError(f"Slide {index} is not a JSON object.")
        slide = dict(raw_slide)
        slide_type = str(slide.get("type", "content")).lower().strip()
        if slide_type not in ALLOWED_SLIDE_TYPES:
            slide_type = "content"
        slide["type"] = slide_type
        slide.setdefault("headline", f"Slide {index}")

        for key in ("items", "left", "right", "left_items", "right_items", "strategic", "tactical", "layers", "nodes", "factors", "metrics", "stages"):
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
            slide.setdefault("right_title", "Client")
            slide.setdefault("center", "Collaboration")
        elif slide_type == "roadmap":
            slide.setdefault("strategic", [])
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

        cleaned.append(slide)

    return {
        "deck": str(spec.get("deck", "PRESENT Local Gemma Plan")),
        "version": str(spec.get("version", "0.2")),
        "slides": cleaned,
    }


def plan_deck_with_gemma(
    deck_info: dict,
    markdown_text: str,
    assets: list[dict] | None = None,
    contact_sheet: str | None = None,
    model: str = DEFAULT_MODEL,
) -> dict:
    if not ollama_available(model):
        raise LocalAgentUnavailable(
            f"Local Ollama model '{model}' is not available. Start Ollama and ensure the model is installed."
        )

    prompt = f"""
You are the local presentation planning agent for PRESENT.

Your job is to turn an EXISTING PowerPoint deck plus its embedded Markdown brief into a visually intentional executive slide plan. Python will render your JSON using native PowerPoint shapes and extracted source imagery.

DESIGN STANDARD
- Apple keynote simplicity plus McKinsey consulting clarity.
- One dominant idea per slide.
- Minimal text, strong hierarchy, whitespace, visual storytelling.
- Use the pearl imagery as a teaching metaphor, not decoration for its own sake.
- Do not merely turn every paragraph into bullets.
- Reuse source-deck images when they materially improve the story.
- Never invent facts, prices, statistics, citations, or claims.

SUPPORTED VISUAL ARCHETYPES
hero: cover or major opening statement. Fields: headline, subheadline, optional asset_id.
metaphor: source image paired with one conceptual message. Fields: headline, subheadline, message OR items, optional asset_id.
comparison: two-sided contrast. Fields: headline, left_title, left[], right_title, right[].
content: 2-6 concise executive idea cards. Fields: headline, subheadline, items[].
process: 3-5 sequential stages. Fields: headline, subheadline, stages[] where each stage is {{"label":"...","detail":"..."}}.
lifecycle: 3-5 maturity stages. Same stage format as process.
collaboration: two parties feeding a shared center. Fields: headline, left_title, left_items[], center, center_detail, right_title, right_items[].
roadmap: synchronized strategic and tactical lanes. Fields: headline, strategic_title, strategic[], tactical_title, tactical[], cadence.
layers: progressive layering around a nucleus. Fields: headline, center, layers[] where each layer may be {{"label":"...","detail":"..."}}.
ecosystem: central capability connected to surrounding nodes. Fields: headline, center, nodes[], optional asset_id.
pricing: pricing logic. Fields: headline, formula[], factors[], payment.
benchmark: market-context metric cards. Fields: headline, metrics[] where each metric is {{"label":"...","value":"...","detail":"..."}}.
statement: closing or major declaration. Fields: headline, subheadline, optional asset_id.

HARD RULES
1. Return JSON only. No markdown fences or commentary.
2. Prefer exactly one output slide for each numbered section in the brief's Slide Blueprint.
3. Preserve the intended narrative order.
4. Use the richest appropriate archetype rather than defaulting to content.
5. Keep individual labels short. Keep detail text to one sentence or less.
6. Use no more than 6 items/nodes/factors on one slide unless the brief clearly requires more.
7. If the contact sheet is attached, inspect it. Use asset_id only for an image that actually supports that slide.
8. Do not assign the same asset to every slide. A few strong image-led slides are better than visual clutter.
9. Keep PEARL Methodology (how Apex Shift works) distinct from Capability Lifecycle (what gets progressively built) if the brief makes that distinction.
10. Preserve source-supported benchmark numbers exactly when using them.

Required output shape:
{{
  "deck": "...",
  "version": "0.2",
  "slides": [ ... ]
}}

EXISTING DECK INVENTORY:
{_deck_context(deck_info)}

SOURCE IMAGE CATALOG:
{_asset_context(assets)}

EMBEDDED MARKDOWN BRIEF:
{markdown_text}
""".strip()

    message = {
        "role": "user",
        "content": prompt,
    }
    if contact_sheet and Path(contact_sheet).exists():
        message["images"] = [str(Path(contact_sheet))]

    try:
        response = _client().chat(
            model=model,
            messages=[message],
            format="json",
            options={"temperature": 0.12},
        )
    except Exception as exc:
        raise LocalAgentUnavailable(f"Local Gemma planning call failed: {exc}") from exc

    raw = response.message.content.strip()
    if not raw:
        raise ValueError("Gemma returned an empty response.")
    try:
        spec = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Gemma returned invalid JSON: {exc}") from exc
    return _validate_spec(spec)
