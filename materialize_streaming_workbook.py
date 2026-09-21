"""Materialize the public workbook with a streaming XLSX writer.

The earlier materializer opened and saved the whole workbook in normal
``openpyxl`` mode several times.  That is safe for small files but becomes
unboundedly slow once People and Relationships contain hundreds of thousands
of observations.  This builder reads the committed workbook in read-only mode
and writes a fresh, reader-facing workbook with XlsxWriter's constant-memory
mode.  The public observation sheets are preserved, while the compact fast
entity index and append-only relation state supply the new unique rows.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
from copy import copy
from pathlib import Path

import xlsxwriter
from openpyxl import load_workbook

from build_entity_network_layer import (
    GCC,
    clipped,
    choose_longest,
    exclusion_reason,
    norm,
)
from relation_state import load_relation_state


ROOT = Path(__file__).parent
WORKBOOK = ROOT / "inputs/entity_network_master.xlsx"
INDEX = ROOT / "data/fast_entity_index.json"
TARGET_UNIQUE = 100_000
GCC_SOURCE_ID = "SRC-068"
GENERATED_SHEETS = {"Unique_People", "GCC_Unique_People", "Excluded_Candidates"}

UNIQUE_HEADERS = [
    "Entity_ID", "主姓名", "别名/观测姓名", "国家/酋长国", "统治家族/支系",
    "身份匹配依据", "实体键/QID", "匹配置信度", "主要记录状态", "人物简介",
    "关系摘要", "关系类型", "观察记录数", "来源数", "来源ID", "来源URL",
    "原始Observation_ID", "截至日期", "下一步核验", "去重说明",
]
EXCLUDED_HEADERS = [
    "Exclusion_ID", "主姓名", "别名/观测姓名", "国家/酋长国", "统治家族/支系",
    "排除原因", "观察记录数", "来源ID", "来源URL", "原始Observation_ID",
]


def parts(value: object) -> list[str]:
    return [item.strip() for item in re.split(r"[；;]", str(value or "")) if item.strip()]


def append_unique(values: list[str], value: object) -> None:
    text = str(value or "").strip()
    if text and text not in values:
        values.append(text)


def entity_id_for_qid(qid: str) -> str:
    return "ENT-" + hashlib.sha1(f"QID:{qid}".encode("utf-8")).hexdigest()[:16].upper()


def relationship_row(record: dict[str, object]) -> list[object]:
    return [
        f"WDC5-REL-{record['record_id']}",
        f"UNRESOLVED-MEMBER:{record['member_qid']}",
        record.get("member_name", ""),
        record["record_id"],
        record.get("related_name") or record.get("name", ""),
        record.get("relation_type", ""),
        record.get("relationship", ""),
        "补充 GCC Wikidata 显式父母、兄弟姐妹、配偶和子女 claims；不等于独立血缘证明或现任王室身份。",
        "P3",
        record.get("evidence_level", ""),
        GCC_SOURCE_ID,
        record.get("page_url", ""),
        "人工核验两个端点、关系方向、家族归属和跨来源重复",
    ]


def people_row(record: dict[str, object]) -> list[object]:
    country = str(record.get("country_section") or "")
    return [
        record["record_id"],
        country,
        record.get("house", ""),
        f"{country}/Wikidata entity claims wave 5/{record.get('edge', '')}",
        record.get("name") or record.get("related_name", ""),
        "",
        "Public structured family-network candidate; fifth-wave endpoint human check passed; current identity, role and living status not assumed",
        record.get("intro", ""),
        record.get("relationship", ""),
        record.get("relation_type", ""),
        record.get("record_status", ""),
        "Wikidata GCC entity-claim wave-5 observation",
        record.get("evidence_level", ""),
        GCC_SOURCE_ID,
        record.get("page_url", ""),
        record.get("as_of", ""),
        "Confirm both endpoint identities, edge direction, family affiliation and current status with an independent source",
        f"Source member QID={record['member_qid']}; related endpoint QID={record['related_qid']}; fifth-wave expansion",
        "",
    ]


def unique_meta(row: list[object]) -> dict[str, object]:
    return {
        "entity_id": str(row[0] or ""),
        "name": str(row[1] or ""),
        "names": parts(row[2] or row[1]),
        "countries": parts(row[3]),
        "houses": parts(row[4]),
        "basis": str(row[5] or ""),
        "keys": parts(row[6]),
        "confidence": str(row[7] or ""),
        "status": str(row[8] or ""),
        "intros": [str(row[9] or "").strip()] if str(row[9] or "").strip() else [],
        "relations": parts(row[10]),
        "relation_types": parts(row[11]),
        "observations": int(row[12] or 0),
        "sources": parts(row[14]),
        "urls": parts(row[15]),
        "observation_ids": parts(row[16]),
        "as_of": str(row[17] or ""),
        "review": str(row[18] or ""),
        "notes": str(row[19] or ""),
    }


def meta_to_row(meta: dict[str, object]) -> list[object]:
    names = list(meta["names"])
    countries = list(meta["countries"])
    houses = list(meta["houses"])
    keys = list(meta["keys"])
    observations = list(meta["observation_ids"])
    sources = list(meta["sources"])
    notes = str(meta["notes"] or "")
    if not notes:
        notes = f"由 {int(meta['observations']):,} 条 People 观察合并；原始观察保留在 People。"
    return [
        meta["entity_id"],
        meta["name"] or choose_longest(names),
        clipped(names, 10_000),
        "；".join(countries),
        "；".join(houses),
        meta["basis"] or "Wikidata QID",
        clipped(keys, 10_000),
        meta["confidence"] or "B2",
        meta["status"],
        choose_longest(list(meta["intros"])),
        clipped(list(meta["relations"]), 30_000),
        "；".join(meta["relation_types"]),
        int(meta["observations"]),
        len(sources),
        "；".join(sources),
        clipped(list(meta["urls"]), 30_000),
        clipped(observations, 30_000),
        meta["as_of"],
        meta["review"] or "Confirm identity, family affiliation, relationship direction and current status with independent sources",
        notes,
    ]


def apply_record(meta: dict[str, object], record: dict[str, object], key: str) -> None:
    name = str(record.get("name") or record.get("related_name") or "").strip()
    country = str(record.get("country_section") or "").strip()
    house = str(record.get("house") or "").strip()
    relation = str(record.get("relationship") or "").strip()
    relation_type = str(record.get("relation_type") or "").strip()
    record_id = str(record.get("record_id") or "").strip()
    source_id = GCC_SOURCE_ID
    url = str(record.get("page_url") or "").strip()
    as_of = str(record.get("as_of") or "").strip()
    if name:
        append_unique(meta["names"], name)
        if len(name) > len(str(meta["name"] or "")):
            meta["name"] = name
    append_unique(meta["countries"], country)
    append_unique(meta["houses"], house)
    append_unique(meta["keys"], key)
    append_unique(meta["sources"], source_id)
    append_unique(meta["urls"], url)
    append_unique(meta["observation_ids"], record_id)
    if relation:
        append_unique(meta["relations"], relation)
    append_unique(meta["relation_types"], relation_type)
    if record.get("record_status") and not meta["status"]:
        meta["status"] = str(record["record_status"])
    intro = str(record.get("intro") or "").strip()
    if intro:
        append_unique(meta["intros"], intro)
    if as_of and as_of > str(meta["as_of"] or ""):
        meta["as_of"] = as_of
    meta["observations"] = int(meta["observations"]) + 1


def new_meta(entity_id: str, entity: dict[str, object]) -> dict[str, object]:
    names = [str(v).strip() for v in entity.get("names", []) if str(v).strip()]
    keys = [str(v).strip() for v in entity.get("keys", []) if str(v).strip()]
    return {
        "entity_id": entity_id,
        "name": choose_longest(names),
        "names": names,
        "countries": [str(v).strip() for v in entity.get("countries", []) if str(v).strip()],
        "houses": [str(v).strip() for v in entity.get("houses", []) if str(v).strip()],
        "basis": "Wikidata QID" if any(v.startswith("QID:") for v in keys) else "Normalized name+country+house",
        "keys": keys,
        "confidence": "B2" if any(v.startswith("QID:") for v in keys) else "C",
        "status": "",
        "intros": [],
        "relations": [],
        "relation_types": [],
        "observations": 0,
        "sources": [],
        "urls": [],
        "observation_ids": [],
        "as_of": "",
        "review": "Confirm identity, family affiliation, relationship direction and current status with independent sources",
        "notes": "由紧凑公共实体索引恢复；关系/观察明细保留在 People 与 Relationships。",
    }


def source_snapshot() -> tuple[list[str], dict[str, list[object]], list[list[object]], list[str], set[str], set[str], list[list[object]], list[list[object]]]:
    """Read only the compact source metadata needed before streaming output."""
    wb = load_workbook(WORKBOOK, read_only=True, data_only=True)
    sheet_names = list(wb.sheetnames)
    unique_rows: dict[str, list[object]] = {}
    excluded_rows: list[list[object]] = []
    source_people_ids: set[str] = set()
    source_relationship_ids: set[str] = set()
    people_headers: list[object] = []
    relationship_headers: list[object] = []
    if "People" in wb.sheetnames:
        for row_num, row in enumerate(wb["People"].iter_rows(values_only=True)):
            if row_num == 0:
                people_headers = list(row)
            elif row and row[0]:
                source_people_ids.add(str(row[0]).strip())
    if "Relationships" in wb.sheetnames:
        for row_num, row in enumerate(wb["Relationships"].iter_rows(values_only=True)):
            if row_num == 0:
                relationship_headers = list(row)
            elif row and row[0]:
                source_relationship_ids.add(str(row[0]).strip())
    if "Unique_People" in wb.sheetnames:
        for row in wb["Unique_People"].iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                unique_rows[str(row[0]).strip()] = list(row[:20])
    if "Excluded_Candidates" in wb.sheetnames:
        for row in wb["Excluded_Candidates"].iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                excluded_rows.append(list(row[:10]))
    wb.close()
    return sheet_names, unique_rows, excluded_rows, people_headers, source_people_ids, source_relationship_ids, relationship_headers, []


def read_index() -> dict[str, object]:
    return json.loads(INDEX.read_text(encoding="utf-8"))


def add_source_metadata(rows: list[list[object]], sheet: str, existing_ids: set[str]) -> list[list[object]]:
    if sheet == "Sources":
        if GCC_SOURCE_ID not in existing_ids:
            rows.append([
                GCC_SOURCE_ID,
                "开放结构化数据",
                "Wikidata GCC entity claims wave 5",
                "https://www.wikidata.org/w/api.php",
                "从当前严格 GCC QID 集合继续读取 P22/P25/P3373/P26/P40 claims，排除既有关系键",
                "低/结构化关系扩展",
                "2026-09-19",
                "第五波关系边仍需独立核验；端点 P31=human 不等于现任王室身份",
            ])
    elif sheet == "Attachment_Readout":
        if GCC_SOURCE_ID not in existing_ids:
            rows.append([
                GCC_SOURCE_ID,
                "Wikidata GCC entity claims 第五波。",
                "GCC 结构化关系扩展层",
                "从当前严格 GCC QID 集合读取显式父母、兄弟姐妹、配偶和子女 claims，按 member/related/edge/context 排除重复",
                "否",
                "结构化 claims 是关系线索；需独立核验",
            ])
    return rows


def configure_sheet(ws, row_count: int, col_count: int, header_format: object) -> None:
    ws.freeze_panes(1, 0)
    ws.hide_gridlines(2)
    if row_count > 0 and col_count > 0:
        ws.autofilter(0, 0, row_count - 1, col_count - 1)
    ws.set_row(0, 30, header_format)
    ws.set_column(0, min(col_count - 1, 19), 20)


def main() -> None:
    sheet_names, source_unique_rows, source_excluded_rows, _, source_people_ids, source_relationship_ids, _, _ = source_snapshot()
    index = read_index()
    entities = index.get("entities", {})
    key_to_entity = index.get("key_to_entity", {})
    excluded_qids = set(index.get("excluded_qids", []))
    relation_state = load_relation_state()
    relation_records = relation_state.get("records", [])

    metas: dict[str, dict[str, object]] = {
        entity_id: unique_meta(row) for entity_id, row in source_unique_rows.items()
    }
    for entity_id, entity in entities.items():
        if entity_id not in metas:
            metas[entity_id] = new_meta(entity_id, entity)

    indexed_observations = {
        observation_id
        for meta in metas.values()
        for observation_id in meta["observation_ids"]
    }
    new_people_records: list[dict[str, object]] = []
    new_relationship_records: list[list[object]] = []
    new_excluded: list[list[object]] = []
    excluded_observation_ids = {
        observation_id for row in source_excluded_rows for observation_id in parts(row[9] if len(row) > 9 else "")
    }

    for record in relation_records:
        record_id = str(record.get("record_id") or "").strip()
        country = str(record.get("country_section") or "").strip()
        if not record_id or country not in GCC:
            continue
        qid = str(record.get("related_qid") or "").strip()
        name = str(record.get("name") or record.get("related_name") or "").strip()
        url = str(record.get("page_url") or "").strip()
        house = str(record.get("house") or "").strip()
        excluded = qid in excluded_qids or bool(exclusion_reason([name], [url], [country, house]))
        if excluded:
            if record_id not in excluded_observation_ids:
                exc_key = f"QID:{qid}" if qid else record_id
                new_excluded.append([
                    "EXC-" + hashlib.sha1(f"{exc_key}|{record_id}".encode("utf-8")).hexdigest()[:16].upper(),
                    name,
                    name,
                    country,
                    house,
                    exclusion_reason([name], [url], [country, house]) or "范围排除：已标记为排除候选",
                    1,
                    GCC_SOURCE_ID,
                    url,
                    record_id,
                ])
                excluded_observation_ids.add(record_id)
            continue

        key = f"QID:{qid}" if qid else ""
        entity_id = key_to_entity.get(key) if key else None
        if not entity_id:
            entity_id = entity_id_for_qid(qid) if qid else "ENT-" + hashlib.sha1(record_id.encode("utf-8")).hexdigest()[:16].upper()
            if key:
                key_to_entity[key] = entity_id
            if entity_id not in metas:
                metas[entity_id] = new_meta(entity_id, {"keys": [key] if key else [], "names": [], "countries": [], "houses": []})
        if entity_id not in metas:
            metas[entity_id] = new_meta(entity_id, entities.get(entity_id, {"keys": [key], "names": [], "countries": [], "houses": []}))
        if record_id not in indexed_observations:
            apply_record(metas[entity_id], record, key or record_id)
            indexed_observations.add(record_id)
        if record_id not in source_people_ids:
            new_people_records.append(record)
        relation_id = f"WDC5-REL-{record_id}"
        if relation_id not in source_relationship_ids:
            new_relationship_records.append(relationship_row(record))

    unique_rows = [meta_to_row(meta) for meta in metas.values()]
    unique_rows.sort(key=lambda row: (str(row[3]), str(row[4]), str(row[1]), str(row[0])))
    gcc_rows = [
        row for row in unique_rows
        if (countries := {value.strip() for value in str(row[3] or "").split("；") if value.strip()}) and countries <= GCC
    ]
    excluded_rows = source_excluded_rows + new_excluded

    temp_path = WORKBOOK.with_suffix(".streaming.xlsx")
    if temp_path.exists():
        temp_path.unlink()
    workbook = xlsxwriter.Workbook(str(temp_path), {"constant_memory": True, "strings_to_urls": False})
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "text_wrap": True, "valign": "vcenter"})
    cell_format = workbook.add_format({"valign": "top"})
    generated_rows = {
        "Unique_People": (UNIQUE_HEADERS, unique_rows),
        "GCC_Unique_People": (UNIQUE_HEADERS, gcc_rows),
        "Excluded_Candidates": (EXCLUDED_HEADERS, excluded_rows),
    }

    source_wb = load_workbook(WORKBOOK, read_only=True, data_only=True)
    source_people_count = max(source_wb["People"].max_row - 1, 0) if "People" in source_wb.sheetnames else 0
    source_relationship_count = max(source_wb["Relationships"].max_row - 1, 0) if "Relationships" in source_wb.sheetnames else 0
    people_total = source_people_count + len(new_people_records)
    relationship_total = source_relationship_count + len(new_relationship_records)
    dashboard_rows = list(source_wb["Dashboard"].iter_rows(values_only=True)) if "Dashboard" in source_wb.sheetnames else []
    dashboard_b27 = dashboard_rows[26][1] if len(dashboard_rows) > 26 and len(dashboard_rows[26]) > 1 else 0
    dashboard_b28 = dashboard_rows[27][1] if len(dashboard_rows) > 27 and len(dashboard_rows[27]) > 1 else 0
    source_count = max(source_wb["Sources"].max_row - 1, 0) if "Sources" in source_wb.sheetnames else 0
    if GCC_SOURCE_ID not in {
        str(row[0]).strip()
        for row in source_wb["Sources"].iter_rows(min_row=2, values_only=True)
        if row and row[0]
    }:
        source_count += 1
    people_count = 0
    relationship_count = 0
    country_counts: dict[str, int] = {}
    for sheet_name in sheet_names:
        ws = workbook.add_worksheet(sheet_name)
        if sheet_name in generated_rows:
            headers, rows = generated_rows[sheet_name]
            ws.write_row(0, 0, headers, header_format)
            for row_num, row in enumerate(rows, start=1):
                ws.write_row(row_num, 0, row, cell_format)
            configure_sheet(ws, len(rows) + 1, len(headers), header_format)
            continue

        row_count = 0
        col_count = 0
        source_ws = source_wb[sheet_name]
        existing_ids: set[str] = set()
        for row_num, values in enumerate(source_ws.iter_rows(values_only=True)):
            row = list(values)
            if row_num == 0:
                ws.write_row(0, 0, row, header_format)
                col_count = len(row)
                continue
            if row and row[0] and sheet_name in {"Sources", "Attachment_Readout"}:
                existing_ids.add(str(row[0]).strip())
            if sheet_name == "People":
                people_count += 1
                country = str(row[1] or "") if len(row) > 1 else ""
                country_counts[country] = country_counts.get(country, 0) + 1
            elif sheet_name == "Relationships":
                relationship_count += 1
            elif sheet_name == "Dashboard":
                row_index = row_num + 1
                if len(row) < 5:
                    row.extend([None] * (5 - len(row)))
                if row_index == 4:
                    row[1] = people_total
                elif row_index == 5:
                    row[1] = relationship_total
                elif row_index == 8:
                    row[1] = max(TARGET_UNIQUE - len(unique_rows), 0)
                elif row_index == 9:
                    row[1] = len(gcc_rows)
                elif row_index == 10:
                    row[1] = len(unique_rows) - len(gcc_rows)
                elif row_index == 29:
                    row[1] = len(gcc_rows)
                elif row_index == 30 and dashboard_b27:
                    row[1] = round(len(gcc_rows) / float(dashboard_b27), 4)
                elif row_index == 31 and dashboard_b28:
                    row[1] = round(len(gcc_rows) / float(dashboard_b28), 4)
            elif sheet_name == "Coverage":
                if len(row) > 4:
                    row[4] = country_counts.get(str(row[0] or ""), 0)
                if row_num == 3 and len(row) > 9:
                    row[9] = source_count
                if row_num == 4 and len(row) > 9:
                    row[9] = people_count
                if row_num == 5 and len(row) > 9:
                    row[9] = relationship_count
            elif sheet_name == "ReadMe":
                if row_num + 1 == 5 and len(row) > 1:
                    row[1] = (
                        f"People 表目前 {people_total:,} 条观察记录；Unique_People 当前 {len(unique_rows):,} 个保守去重实体，"
                        f"其中严格 GCC {len(gcc_rows):,} 个，距离 100,000 个唯一实体目标还差 {max(TARGET_UNIQUE - len(unique_rows), 0):,} 个。"
                    )
                elif row_num + 1 == 6 and len(row) > 1:
                    row[1] = (
                        f"GCC_Unique_People 为严格 GCC 六国视图，共 {len(gcc_rows):,} 个实体；公开结构化关系是线索，"
                        "不等于独立核验的现任王室身份或在世状态。"
                    )
            ws.write_row(row_num, 0, row, cell_format)
            row_count = row_num

        if sheet_name == "People":
            for record in new_people_records:
                row_num = row_count + 1
                row = people_row(record)
                ws.write_row(row_num, 0, row, cell_format)
                people_count += 1
                country = str(row[1] or "")
                country_counts[country] = country_counts.get(country, 0) + 1
                row_count = row_num
        elif sheet_name == "Relationships":
            for row in new_relationship_records:
                row_num = row_count + 1
                ws.write_row(row_num, 0, row, cell_format)
                relationship_count += 1
                row_count = row_num
        elif sheet_name in {"Sources", "Attachment_Readout"}:
            additions = add_source_metadata([], sheet_name, existing_ids)
            for row in additions:
                row_num = row_count + 1
                ws.write_row(row_num, 0, row, cell_format)
                row_count = row_num
        configure_sheet(ws, row_count + 1, max(col_count, 1), header_format)

    # Dashboard and ReadMe are copied above; patching their cells requires a
    # second lightweight output pass, so write dedicated summary sheets before
    # closing by replacing the values during the copy is handled below.
    source_wb.close()
    workbook.close()

    # XlsxWriter writes rows sequentially and cannot edit an already closed
    # sheet.  The summary values are therefore applied in the copy function on
    # the next run through the same builder; keep the generated metrics in the
    # progress file and use the explicit ReadMe sheet as the source of truth.
    progress = {
        "people_observations": people_count,
        "relationship_observations": relationship_count,
        "unique_entities": len(unique_rows),
        "strict_scope_entities": len(gcc_rows),
        "unique_gap_to_100k": max(TARGET_UNIQUE - len(unique_rows), 0),
        "new_people_records": len(new_people_records),
        "new_relationship_records": len(new_relationship_records),
        "excluded_candidates": len(excluded_rows),
        "output": str(WORKBOOK),
    }
    os.replace(temp_path, WORKBOOK)
    print(json.dumps(progress, ensure_ascii=False))


if __name__ == "__main__":
    main()
