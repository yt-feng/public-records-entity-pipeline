"""Read and write the append-only public relationship state.

The uncompressed JSON state eventually crosses GitHub's 100 MB per-file push
limit.  Keep the same public JSON schema inside a gzip member so Actions can
continue to persist the full relationship layer without dropping records.
"""

from __future__ import annotations

import gzip
import json
from pathlib import Path


ROOT = Path(__file__).parent
OUTPUT = ROOT / "data/relation_wave_records.json.gz"
LEGACY_OUTPUT = ROOT / "data/relation_wave_records.json"


def load_relation_state() -> dict:
    if OUTPUT.exists():
        with gzip.open(OUTPUT, "rt", encoding="utf-8") as handle:
            return json.load(handle)
    if LEGACY_OUTPUT.exists():
        return json.loads(LEGACY_OUTPUT.read_text(encoding="utf-8"))
    return {}


def write_relation_state(payload: dict) -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with gzip.open(OUTPUT, "wt", encoding="utf-8", compresslevel=6) as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    # The legacy file is removed only after the compressed replacement has been
    # written successfully, so a failed compression cannot erase the state.
    if LEGACY_OUTPUT.exists():
        LEGACY_OUTPUT.unlink()
