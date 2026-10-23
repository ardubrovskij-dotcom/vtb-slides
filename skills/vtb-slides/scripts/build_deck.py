"""Build a VTB deck from a JSON spec.

    python build_deck.py spec.json out.pptx

spec = {"font_mode": "vtb" | "safe", "title_case": "upper" | "as_is",
        "slides": [{"type": "cover", "title": "...", "subtitle": "...", "variant": 4}, ...]}
`type` is any Deck method (cover, section, agenda, cards, chart, ...); other keys are its arguments.
Relative image paths (photo, image) are resolved against the spec's folder.
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vtb import Deck  # noqa: E402

IMAGE_KEYS = ("photo", "image")


def _resolve(obj, base):
    if isinstance(obj, dict):
        return {k: (os.path.join(base, v) if k in IMAGE_KEYS and isinstance(v, str) and not os.path.isabs(v)
                    and os.path.exists(os.path.join(base, v)) else _resolve(v, base)) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_resolve(v, base) for v in obj]
    return obj


def build(spec, out, base="."):
    d = Deck(font_mode=spec.get("font_mode", "vtb"), title_case=spec.get("title_case", "upper"))
    for i, sl in enumerate(spec["slides"], 1):
        sl = _resolve(dict(sl), base)
        kind = sl.pop("type")
        fn = getattr(d, kind, None)
        if fn is None or kind.startswith("_") or kind in ("save", "text", "shape"):
            raise SystemExit(f"slide {i}: unknown type {kind!r}")
        try:
            fn(**sl)
        except TypeError as e:
            raise SystemExit(f"slide {i} ({kind}): {e}")
    d.save(out)
    return d


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    spec_path, out = sys.argv[1], sys.argv[2]
    with open(spec_path, encoding="utf8") as f:
        spec = json.load(f)
    d = build(spec, out, os.path.dirname(os.path.abspath(spec_path)))
    print(f"OK {out}: {len(d.prs.slides)} slides, {len(d.warnings)} warnings")
