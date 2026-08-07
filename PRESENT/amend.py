from pathlib import Path
from typing import Callable

from pptx import Presentation

from agent import LocalAgentUnavailable, plan_deck_with_gemma
from assets import extract_source_assets
from builder import add_slides
from ingest import find_embedded_markdown, inspect_deck, markdown_to_slide_spec


ProgressCallback = Callable[[str], None] | None


def amend_deck(
    source_pptx: str,
    output_pptx: str,
    use_agent: bool = True,
    progress_callback: ProgressCallback = None,
) -> dict:
    source = Path(source_pptx)
    output = Path(output_pptx)

    def report(message: str):
        if progress_callback:
            progress_callback(message)

    report("Validating source and output files...")
    if not source.exists():
        raise FileNotFoundError(source)
    if source.resolve() == output.resolve():
        raise ValueError("Output must be a different file so the source deck is not overwritten.")

    report("Locating the embedded Markdown presentation brief...")
    markdown_slide, markdown_text = find_embedded_markdown(str(source))

    report("Extracting and cataloging source-deck images...")
    asset_dir = output.parent / ".present_assets" / source.stem
    assets, contact_sheet = extract_source_assets(str(source), asset_dir)

    planner = "deterministic parser"
    agent_error = None

    if use_agent:
        try:
            report("Inspecting the existing deck structure and content...")
            deck_info = inspect_deck(str(source))

            report("Asking local Gemma to plan the presentation and visual treatments...")
            spec = plan_deck_with_gemma(
                deck_info,
                markdown_text,
                assets=assets,
                contact_sheet=contact_sheet,
            )
            planner = "local Gemma via Ollama + source-image vision"
        except (LocalAgentUnavailable, ValueError) as exc:
            agent_error = str(exc)
            report("Gemma planning was unavailable; using PRESENT's deterministic fallback...")
            spec = markdown_to_slide_spec(markdown_text)
    else:
        report("Parsing the embedded Markdown presentation blueprint...")
        spec = markdown_to_slide_spec(markdown_text)

    report("Opening the source PowerPoint and preserving existing slides...")
    prs = Presentation(source)
    original_slide_count = len(prs.slides)

    report("Composing the new visual slides...")
    add_slides(prs, spec, assets=assets)

    report("Saving the amended PowerPoint...")
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)

    report("Build complete.")
    return {
        "source": str(source),
        "output": str(output),
        "markdown_slide": markdown_slide,
        "original_slide_count": original_slide_count,
        "added_slide_count": len(spec.get("slides", [])),
        "final_slide_count": len(prs.slides),
        "planner": planner,
        "agent_error": agent_error,
        "asset_count": len(assets),
        "contact_sheet": contact_sheet,
    }
