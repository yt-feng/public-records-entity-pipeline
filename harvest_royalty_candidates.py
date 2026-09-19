"""Harvest a bounded MENA royalty/family candidate layer from Wikidata.

This is deliberately a candidate layer: a Wikidata occupation or family claim
does not by itself prove present-day status, residence, wealth, or identity.
The query is split into regional batches so the public endpoint can be resumed
and a single slow country group does not discard successful results.
"""

from __future__ import annotations

import hashlib
import json
import subprocess
import time
from pathlib import Path


ROOT = Path(__file__).parent
OUTPUT = ROOT / "data/structured_candidate_records.json"
AS_OF = "2026-09-20"
SOURCE_ID = "SRC-070"

COUNTRY_BATCHES = {
    "gcc": {
        "Saudi Arabia": "Q851",
        "Qatar": "Q846",
        "United Arab Emirates": "Q878",
        "Kuwait": "Q817",
        "Bahrain": "Q398",
        "Oman": "Q842",
    },
    "levant": {
        "Jordan": "Q810",
        "Iraq": "Q796",
        "Syria": "Q858",
        "Lebanon": "Q822",
        "Palestine": "Q219",
        "Yemen": "Q805",
    },
    "north_africa": {
        "Egypt": "Q79",
        "Morocco": "Q1028",
        "Algeria": "Q262",
        "Tunisia": "Q948",
        "Libya": "Q1016",
        "Sudan": "Q1049",
    },
}


def value(row: dict, key: str) -> str:
    return row.get(key, {}).get("value", "").strip()


def run_query(query: str) -> list[dict]:
    last_error = "empty response"
    for attempt in range(3):
        try:
            result = subprocess.run(
                [
                    "curl", "-L", "--fail-with-body", "--max-time", "180", "--connect-timeout", "20", "-sS", "-G",
                    "https://query.wikidata.org/sparql", "--data-urlencode", f"query={query}",
                    "--data-urlencode", "format=json", "-A", "Public-Records-Research/0.3",
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
            if attempt == 2:
                raise RuntimeError(last_error) from exc
            time.sleep(10 * (attempt + 1))
    raise RuntimeError(last_error)


def qids(values: dict[str, str]) -> str:
    return " ".join(f"wd:{qid}" for qid in values.values())


def classify_country(label: str, fallback: str) -> str:
    return label or fallback


def make_family_record(row: dict, country_fallback: str) -> dict | None:
    person_qid = value(row, "person").rsplit("/", 1)[-1]
    name = value(row, "personLabel")
    country = classify_country(value(row, "countryLabel"), country_fallback)
    family_qid = value(row, "family").rsplit("/", 1)[-1]
    family = value(row, "familyLabel")
    if not person_qid or not name or name.startswith("Q") or not family_qid or not family:
        return None
    key = (person_qid, country, family_qid)
    record_id = "WDRY-" + hashlib.sha1("|".join(key).encode("utf-8")).hexdigest()[:16].upper()
    return {
        "record_id": record_id,
        "country_section": country,
        "house": family,
        "page_url": f"https://www.wikidata.org/wiki/{person_qid}",
        "name": name,
        "family_qid": family_qid,
        "source_text_short": f"Public structured data records {name} with family claim {family} and country signal {country}.",
        "record_status": "public structured MENA family candidate; current role, residence and status not assumed",
        "evidence_level": "低/结构化候选",
        "intro": f"Wikidata records {name} as associated with {family}; country signal={country}. This is a candidate for identity review.",
        "relationship": f"Explicit Wikidata family claim: {name} -> {family}; country signal={country}.",
        "relation_type": "候选/公开家族字段/MENA扩展",
        "source_id": SOURCE_ID,
        "as_of": AS_OF,
    }


def make_royalty_record(row: dict, country_fallback: str) -> dict | None:
    person_qid = value(row, "person").rsplit("/", 1)[-1]
    name = value(row, "personLabel")
    country = classify_country(value(row, "countryLabel"), country_fallback)
    if not person_qid or not name or name.startswith("Q"):
        return None
    key = (person_qid, country, "royalty")
    record_id = "WDRY-" + hashlib.sha1("|".join(key).encode("utf-8")).hexdigest()[:16].upper()
    return {
        "record_id": record_id,
        "country_section": country,
        "house": "Royalty occupation signal",
        "page_url": f"https://www.wikidata.org/wiki/{person_qid}",
        "name": name,
        "family_qid": "",
        "source_text_short": f"Public structured data places {name} in a royalty occupation hierarchy with country signal {country}.",
        "record_status": "public structured MENA royalty candidate; occupation and country signals require independent review",
        "evidence_level": "低/结构化候选",
        "intro": f"Wikidata links {name} to the royalty occupation hierarchy; country signal={country}. This does not assert current status.",
        "relationship": f"Structured royalty occupation signal for {name}; country signal={country}.",
        "relation_type": "候选/royalty occupation/MENA扩展",
        "source_id": SOURCE_ID,
        "as_of": AS_OF,
    }


def make_title_record(row: dict, country_fallback: str) -> dict | None:
    person_qid = value(row, "person").rsplit("/", 1)[-1]
    name = value(row, "personLabel")
    country = classify_country(value(row, "countryLabel"), country_fallback)
    title_qid = value(row, "title").rsplit("/", 1)[-1]
    title = value(row, "titleLabel")
    if not person_qid or not name or name.startswith("Q") or not title_qid or not title:
        return None
    key = (person_qid, country, title_qid)
    record_id = "WDRY-" + hashlib.sha1("|".join(key).encode("utf-8")).hexdigest()[:16].upper()
    return {
        "record_id": record_id,
        "country_section": country,
        "house": f"Noble/princely title: {title}",
        "page_url": f"https://www.wikidata.org/wiki/{person_qid}",
        "name": name,
        "family_qid": "",
        "source_text_short": f"Public structured data gives {name} the noble/princely title {title}; country signal={country}.",
        "record_status": "public structured MENA title candidate; title and family identity require independent review",
        "evidence_level": "低/结构化候选",
        "intro": f"Wikidata records {name} with title {title}; country signal={country}. This does not assert current status or family branch.",
        "relationship": f"Structured title signal: {name} -> {title}; country signal={country}.",
        "relation_type": "候选/noble-princely title/MENA扩展",
        "source_id": SOURCE_ID,
        "as_of": AS_OF,
    }


def main() -> None:
    records_by_id: dict[str, dict] = {}
    if OUTPUT.exists():
        existing = json.loads(OUTPUT.read_text(encoding="utf-8"))
        records_by_id = {record.get("record_id"): record for record in existing.get("records", []) if record.get("record_id")}

    successful: list[str] = []
    failed: list[str] = []
    for batch_name, countries in COUNTRY_BATCHES.items():
        country_values = qids(countries)
        family_query = f'''SELECT DISTINCT ?person ?personLabel ?country ?countryLabel ?family ?familyLabel WHERE {{
          VALUES ?country {{ {country_values} }}
          ?person wdt:P31 wd:Q5 ; wdt:P27 ?country ; wdt:P53 ?family .
          ?family rdfs:label ?familyLabel .
          FILTER(LANG(?familyLabel)="en")
          SERVICE wikibase:label {{ bd:serviceParam wikibase:language "en". }}
        }} LIMIT 50000'''
        royalty_query = f'''SELECT DISTINCT ?person ?personLabel ?country ?countryLabel WHERE {{
          VALUES ?country {{ {country_values} }}
          ?person wdt:P31 wd:Q5 ; wdt:P27 ?country ; wdt:P106/wdt:P279* wd:Q11573099 .
          SERVICE wikibase:label {{ bd:serviceParam wikibase:language "en". }}
        }} LIMIT 50000'''
        title_query = f'''SELECT DISTINCT ?person ?personLabel ?country ?countryLabel ?title ?titleLabel WHERE {{
          VALUES ?country {{ {country_values} }}
          ?person wdt:P31 wd:Q5 ; wdt:P27 ?country ; wdt:P97 ?title .
          ?title rdfs:label ?titleLabel .
          FILTER(LANG(?titleLabel)="en")
          FILTER(REGEX(LCASE(STR(?titleLabel)), "prince|princess|emir|emira|sheikh|sheikha|sultan|king|queen|sharif|caliph|sayyid|sayyida"))
          SERVICE wikibase:label {{ bd:serviceParam wikibase:language "en". }}
        }} LIMIT 50000'''
        try:
            family_rows = run_query(family_query)
            royalty_rows = run_query(royalty_query)
            title_rows = run_query(title_query)
            added = 0
            for row in family_rows:
                record = make_family_record(row, batch_name)
                if record and record["record_id"] not in records_by_id:
                    records_by_id[record["record_id"]] = record
                    added += 1
            for row in royalty_rows:
                record = make_royalty_record(row, batch_name)
                if record and record["record_id"] not in records_by_id:
                    records_by_id[record["record_id"]] = record
                    added += 1
            for row in title_rows:
                record = make_title_record(row, batch_name)
                if record and record["record_id"] not in records_by_id:
                    records_by_id[record["record_id"]] = record
                    added += 1
            successful.append(batch_name)
            print(f"batch={batch_name} family_bindings={len(family_rows)} royalty_bindings={len(royalty_rows)} title_bindings={len(title_rows)} new_records={added}", flush=True)
        except RuntimeError as exc:
            failed.append(batch_name)
            print(f"batch={batch_name} skipped after retries: {exc}", flush=True)

    records = list(records_by_id.values())
    OUTPUT.write_text(
        json.dumps({
            "records": records,
            "as_of": AS_OF,
            "source": "Wikidata family membership and royalty occupation hierarchy",
            "successful_batches": successful,
            "failed_batches": failed,
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    if not successful:
        raise RuntimeError("all MENA royalty candidate batches failed")
    print(json.dumps({"records": len(records), "successful_batches": successful, "failed_batches": failed, "output": str(OUTPUT)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
