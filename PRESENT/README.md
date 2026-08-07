# PRESENT

**PRESENT** is an Apex Shift Force-Multiplier Layered Engine (FMLE) for designing, building, amending, reviewing, and refining professional presentation decks.

## Initial Objective

Use an existing PowerPoint deck as the source of truth, preserve valuable assets and research, ingest presentation-specific design/content guidance, and produce a polished, editable `.pptx` through a repeatable generate → render → critique → refine workflow.

The first validation project is the **Apex Shift PEARL Methodology** deck.

## Core Principles

- **Existing-deck aware:** PRESENT should modify and extend an existing deck rather than blindly rebuilding it.
- **Content before decoration:** every slide must have a clear communication objective.
- **Visual storytelling:** diagrams, progression graphics, imagery, and whitespace should carry more of the message than paragraphs.
- **Speaker-friendly:** slides support the presenter rather than becoming teleprompters.
- **Design-system driven:** typography, spacing, colors, image treatment, and recurring components should be centralized and reusable.
- **Model-assisted, deterministic output:** an LLM can reason about narrative and design, while code handles repeatable PowerPoint construction.
- **Closed-loop refinement:** render slides, review them visually, revise, and repeat.

## PRESENT Pipeline

1. **Ingest** — open the source `.pptx`; extract slide text, images, shapes, notes, dimensions, and existing design cues.
2. **Interpret** — identify slide purpose, narrative role, reusable assets, and content constraints.
3. **Plan** — produce a structured slide specification for the target deck.
4. **Design** — select a layout archetype and visual treatment from the PRESENT design system.
5. **Build** — create or amend slides programmatically.
6. **Render** — export slides to images for review.
7. **Critique** — have a multimodal model evaluate hierarchy, clarity, balance, density, consistency, and narrative flow.
8. **Refine** — apply targeted revisions and rerender.
9. **Validate** — confirm text accuracy, visual consistency, editable PowerPoint structure, and absence of overflow/clipping.
10. **Package** — save the finished deck plus its slide specification and build report.

## First Test: PEARL

The PEARL deck should teach **PEARL**, not teach pearls. Pearl imagery is the visual metaphor used to explain the methodology.

**PEARL = Pilot Engagement and Roadmap Layering**

Capability progression:

1. **Pilot** — demonstrate one small, focused business capability through a no-cost, low-risk proof of value.
2. **MVP** — create the first usable layers of business value.
3. **Mature Application** — develop a complete, production-ready capability.
4. **Integration** — connect the mature application into the broader business system / ecosystem.

Metaphor mapping:

- Nucleus → Pilot
- First nacre layers → MVP
- Completed pearl → Mature Application
- Pearl added to necklace → Integrated Application / business ecosystem

## Initial Technology Direction

- Python 3.11+
- `python-pptx` for Open XML inspection and common slide editing
- `pywin32` for native PowerPoint automation on Windows where available
- Pillow for image preprocessing
- SVG generation for sharp diagrams and reusable visual components
- Optional multimodal model (Gemini/Gemma-compatible workflow) for planning and critique
- JSON/YAML slide specification as the contract between the reasoning layer and rendering layer

## Repository Structure

```text
PRESENT/
├── README.md
├── docs/
│   └── architecture.md
├── specs/
│   └── pearl-playbook.md
├── src/
│   └── present/
├── templates/
├── assets/
├── tests/
└── examples/
```

Git does not preserve empty directories, so folders are introduced as files are added.
