from pathlib import Path
from pptx import Presentation

from agent import LocalAgentUnavailable, plan_deck_with_gemma
from builder import add_slides
from ingest import find_embedded_markdown, inspect_deck, markdown_to_slide_spec


def amend_deck(source_pptx: str, output_pptx: str, use_agent: bool = True) -> dict:
    source = Path(source_pptx)
    output = Path(output_pptx)

    if not source.exists():
        raise FileNotFoundError(source)

    markdown_slide, markdown_text = find_embedded_markdown(str(source))
    planner = "deterministic parser"
    agent_error = None

    if use_agent:
        try:
            deck_info = inspect_deck(str(source))
            spec = plan_deck_with_gemma(deck_info, markdown_text)
            planner = "local Gemma via Ollama"
        except (LocalAgentUnavailable, ValueError) as exc:
            agent_error = str(exc)
            spec = markdown_to_slide_spec(markdown_text)
    else:
        spec = markdown_to_slide_spec(markdown_text)

    prs = Presentation(source)
    original_slide_count = len(prs.slides)
    add_slides(prs, spec)

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)

    return {
        "source": str(source),
        "output": str(output),
        "markdown_slide": markdown_slide,
        "original_slide_count": original_slide_count,
        "added_slide_count": len(spec.get("slides", [])),
        "final_slide_count": len(prs.slides),
        "planner": planner,
        "agent_error": agent_error,
    }
