import json
import urllib.error
import urllib.request

OLLAMA_BASE_URL = "http://127.0.0.1:11434"
DEFAULT_MODEL = "gemma3:4b"
ALLOWED_SLIDE_TYPES = {"hero", "comparison", "content", "statement"}


class LocalAgentUnavailable(RuntimeError):
    pass


def _request_json(url: str, payload: dict | None = None, timeout: int = 120) -> dict:
    data = None if payload is None else json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="GET" if payload is None else "POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as exc:
        raise LocalAgentUnavailable(f"Local Ollama is unavailable: {exc}") from exc


def ollama_available(model: str = DEFAULT_MODEL) -> bool:
    try:
        result = _request_json(f"{OLLAMA_BASE_URL}/api/tags", timeout=5)
    except LocalAgentUnavailable:
        return False

    names = {item.get("name", "") for item in result.get("models", [])}
    return model in names or any(name.startswith(f"{model}:") for name in names)


def _deck_context(deck_info: dict) -> str:
    lines = []
    for slide in deck_info.get("slides", []):
        text = slide.get("text", "").strip()
        # The embedded Markdown itself is supplied separately, so avoid duplicating a huge block.
        if "# Apex Shift PEARL Playbook" in text or "# Slide Blueprint" in text:
            text = "[Embedded presentation brief omitted here; supplied separately.]"
        if len(text) > 1200:
            text = text[:1200] + "..."
        lines.append(f"Slide {slide.get('number')}: {text or '[visual/no extracted text]'}")
    return "\n\n".join(lines)


def _validate_spec(spec: dict) -> dict:
    if not isinstance(spec, dict):
        raise ValueError("Gemma did not return a JSON object.")

    slides = spec.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ValueError("Gemma returned no slides.")

    cleaned = []
    for index, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            raise ValueError(f"Slide {index} is not a JSON object.")
        slide_type = slide.get("type", "content")
        if slide_type not in ALLOWED_SLIDE_TYPES:
            slide_type = "content"
        slide["type"] = slide_type
        slide.setdefault("headline", f"Slide {index}")
        if slide_type == "comparison":
            slide.setdefault("left_title", "Current")
            slide.setdefault("left", [])
            slide.setdefault("right_title", "Target")
            slide.setdefault("right", [])
        elif slide_type == "content":
            items = slide.get("items", [])
            if isinstance(items, str):
                items = [items]
            slide["items"] = [str(item) for item in items][:10]
        cleaned.append(slide)

    return {
        "deck": str(spec.get("deck", "PRESENT Local Gemma Plan")),
        "version": str(spec.get("version", "0.1")),
        "slides": cleaned,
    }


def plan_deck_with_gemma(deck_info: dict, markdown_text: str, model: str = DEFAULT_MODEL) -> dict:
    if not ollama_available(model):
        raise LocalAgentUnavailable(
            f"Ollama model '{model}' is not available locally. Start Ollama and ensure the model is installed."
        )

    prompt = f"""
You are the local planning agent for PRESENT, a presentation-building engine.

Your job is to interpret an EXISTING PowerPoint deck plus an embedded Markdown presentation brief and produce a JSON slide plan for NEW slides that PRESENT will append to a copy of the existing deck.

HARD RULES:
1. Use only information supported by the existing deck or embedded brief. Do not invent facts, research, prices, statistics, claims, or citations.
2. The deck should teach the business methodology described in the brief. Pearl imagery is a teaching metaphor, not the subject.
3. Preserve the brief's intended narrative order and numbered slide blueprint.
4. One dominant idea per slide. Keep text concise and speaker-driven.
5. The current MVE builder supports ONLY these slide types: hero, comparison, content, statement.
6. Return JSON ONLY. No markdown fences and no explanation.
7. Prefer exactly one output slide for each numbered section in the brief's Slide Blueprint.
8. For content slides, use at most 8 short items. Do not paste paragraphs from the brief.
9. For comparison slides, use left_title, left, right_title, right.
10. For hero/statement slides, use headline and optional subheadline.
11. Preserve important wording from the brief when it is clearly intentional.
12. Include a short visual_direction string on each slide describing the desired future visual treatment. The current builder may ignore it, but PRESENT will retain it for later refinement.

Required JSON shape:
{{
  "deck": "...",
  "version": "0.1",
  "slides": [
    {{"type":"hero","headline":"...","subheadline":"...","visual_direction":"..."}},
    {{"type":"comparison","headline":"...","left_title":"...","left":["..."],"right_title":"...","right":["..."],"visual_direction":"..."}},
    {{"type":"content","headline":"...","subheadline":"...","items":["..."],"visual_direction":"..."}},
    {{"type":"statement","headline":"...","subheadline":"...","visual_direction":"..."}}
  ]
}}

EXISTING DECK INVENTORY:
{_deck_context(deck_info)}

EMBEDDED MARKDOWN BRIEF:
{markdown_text}
""".strip()

    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "format": "json",
        "options": {
            "temperature": 0.15,
        },
    }

    result = _request_json(f"{OLLAMA_BASE_URL}/api/generate", payload=payload, timeout=240)
    raw = result.get("response", "").strip()
    if not raw:
        raise ValueError("Gemma returned an empty response.")

    try:
        spec = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Gemma returned invalid JSON: {exc}") from exc

    return _validate_spec(spec)
