import json
from pathlib import Path
from builder import build_deck

SPEC_PATH = Path("specs/pearl.json")
OUTPUT_PATH = Path("output/PEARL_PRESENT_MVE.pptx")


def run():
    with SPEC_PATH.open("r", encoding="utf-8") as handle:
        spec = json.load(handle)
    build_deck(spec, str(OUTPUT_PATH))
    print(f"Built: {OUTPUT_PATH}")


if __name__ == "__main__":
    run()
