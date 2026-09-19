"""Harvest a bounded public structured-family candidate layer."""

from __future__ import annotations

import hashlib
import json
import subprocess
import time
from pathlib import Path


ROOT = Path(__file__).parent
OUTPUT = ROOT / "data/structured_candidate_records.json"
AS_OF = "2026-09-20"
FAMILY_PATTERNS = [
    "saud",
    "rashid",
    "thani",
    "sabah",
    "khalifa",
    "nahyan",
    "maktoum",
    "qasimi",
    "nuaimi",
    "sharqi",
    "mualla",
    "falasi",
    "said",
    "hashemite",
    "alaoui",
    "alawi",
]
EXACT_FAMILY_LABELS = [
    "House of Thani",
    "Al Thani",
    "House of Nahyan",
    "Al Nahyan",
]


def value(row: dict, key: str) -> str:
    return row.get(key, {}).get("value", "").strip()


def classify(label: str) -> tuple[str, str]:
    text = label.casefold()
    if "saud" in text or "rashid" in text:
        return "Saudi Arabia", label
    if "thani" in text:
        return "Qatar", label
    if "sabah" in text:
        return "Kuwait", label
    if "khalifa" in text:
        return "Bahrain", label
    if "nahyan" in text or "maktoum" in text or "qasimi" in text or "nuaimi" in text or "sharqi" in text or "mualla" in text or "falasi" in text:
        return "United Arab Emirates", label
    if "said" in text:
        return "Oman", label
    if "hashemite" in text:
        return "Jordan", label
    if "alaoui" in text or "alawi" in text:
        return "Morocco", label
    return "MENA / historical scope", label


def run_query(query: str) -> list[dict]:
    """Run one bounded Wikidata query with retries on transient responses."""
    last_error = "empty response"
    for attempt in range(4):
        try:
            result = subprocess.run(
                [
                    "curl", "-L", "--fail-with-body", "--max-time", "120", "--connect-timeout", "15", "-sS", "-G",
                    "https://query.wikidata.org/sparql", "--data-urlencode", f"query={query}",
                    "--data-urlencode", "format=json", "-A", "Public-Records-Research/0.2",
                ],
                check=True,
                capture_output=True,
                text=True,
            )
            payload = result.stdout.strip()
            if not payload:
                raise ValueError("empty response")
            return json.loads(payload)["results"]["bindings"]
        except (subprocess.CalledProcessError, ValueError, json.JSONDecodeError) as exc:
            last_error = str(exc)
            if attempt == 3:
                raise RuntimeError(last_error) from exc
            time.sleep(5 * (attempt + 1))
    raise RuntimeError(last_error)


def main() -> None:
    bindings = []
    successful_patterns = []
    failed_patterns = []
    for pattern in FAMILY_PATTERNS:
        query = f'''SELECT DISTINCT ?person ?personLabel ?family ?familyLabel WHERE {{
          ?person wdt:P31 wd:Q5 ; wdt:P53 ?family .
          ?family rdfs:label ?familyLabel .
          FILTER(LANG(?familyLabel)="en")
          FILTER(REGEX(LCASE(STR(?familyLabel)), "{pattern}"))
          SERVICE wikibase:label {{ bd:serviceParam wikibase:language "en". }}
        }} LIMIT 5000'''
        try:
            batch = run_query(query)
            bindings.extend(batch)
            successful_patterns.append(pattern)
            print(f"family_pattern={pattern} bindings={len(batch)}", flush=True)
        except RuntimeError as exc:
            failed_patterns.append(pattern)
            print(f"family_pattern={pattern} skipped after retries: {exc}", flush=True)

    # Exact family-label fallback for the two major Qatar/UAE houses whose
    # broad regex queries are intermittently slow at the public endpoint.
    labels = ", ".join(json.dumps(label) + "@en" for label in EXACT_FAMILY_LABELS)
    family_query = f'''SELECT DISTINCT ?family ?familyLabel WHERE {{
      VALUES ?wanted {{ {labels} }}
      ?family rdfs:label ?wanted .
      BIND(?wanted AS ?familyLabel)
    }}'''
    try:
        family_rows = run_query(family_query)
        family_qids = sorted({value(row, "family").rsplit("/", 1)[-1] for row in family_rows if value(row, "family")})
        print(f"exact_family_labels qids={family_qids}", flush=True)
    except RuntimeError as exc:
        family_qids = []
        print(f"exact_family_labels skipped after retries: {exc}", flush=True)
    for family_qid in family_qids:
        query = f'''SELECT DISTINCT ?person ?personLabel ?family ?familyLabel WHERE {{
          ?person wdt:P31 wd:Q5 ; wdt:P53 wd:{family_qid} .
          BIND(wd:{family_qid} AS ?family)
          ?family rdfs:label ?familyLabel .
          FILTER(LANG(?familyLabel)="en")
          SERVICE wikibase:label {{ bd:serviceParam wikibase:language "en". }}
        }} LIMIT 5000'''
        try:
            batch = run_query(query)
            bindings.extend(batch)
            print(f"family_qid={family_qid} bindings={len(batch)}", flush=True)
        except RuntimeError as exc:
            print(f"family_qid={family_qid} skipped after retries: {exc}", flush=True)
    records = []
    seen = set()
    for row in bindings:
        person_qid = value(row, "person").rsplit("/", 1)[-1]
        family_qid = value(row, "family").rsplit("/", 1)[-1]
        name = value(row, "personLabel")
        family = value(row, "familyLabel")
        if not person_qid or not family_qid or not name or not family or name.startswith("Q"):
            continue
        key = (person_qid, family_qid)
        if key in seen:
            continue
        seen.add(key)
        country, house = classify(family)
        record_id = "WDSC-" + hashlib.sha1("|".join(key).encode("utf-8")).hexdigest()[:16].upper()
        records.append({
            "record_id": record_id,
            "country_section": country,
            "house": house,
            "page_url": f"https://www.wikidata.org/wiki/{person_qid}",
            "name": name,
            "family_qid": family_qid,
            "source_text_short": f"Public structured data records explicit dynasty/family membership: {family}.",
            "record_status": "public structured family candidate; current role, residence and status not assumed",
            "evidence_level": "低/结构化候选",
            "intro": f"公开结构化数据将 {name} 记录为 {family} 成员候选；不据此推断当前职位、在世状态或资产。",
            "relationship": f"Explicit public family/dynasty membership field: {family}.",
            "relation_type": "候选/公开家族成员字段",
            "source_id": "SRC-069",
            "as_of": AS_OF,
        })
    OUTPUT.write_text(json.dumps({"records": records, "as_of": AS_OF, "family_patterns": FAMILY_PATTERNS, "successful_patterns": successful_patterns, "failed_patterns": failed_patterns}, ensure_ascii=False, indent=2), encoding="utf-8")
    if not successful_patterns:
        raise RuntimeError("all structured candidate patterns failed")
    print(json.dumps({"bindings": len(bindings), "records": len(records), "successful_patterns": successful_patterns, "failed_patterns": failed_patterns, "output": str(OUTPUT)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
