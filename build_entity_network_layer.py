"""Build a conservative entity-deduplicated people layer in the master workbook.

The existing People sheet is intentionally preserved as an observation table.
This script creates/rebuilds Unique_People from those observations and records
the match basis so a user can distinguish a unique entity from a source row.
"""

from __future__ import annotations

import hashlib
import re
import unicodedata
from collections import Counter, defaultdict
from copy import copy
from pathlib import Path
from urllib.parse import unquote

from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).parent
WORKBOOK = ROOT / "inputs/entity_network_master.xlsx"
GCC = {"Saudi Arabia", "Qatar", "United Arab Emirates", "Kuwait", "Bahrain", "Oman"}
TARGET_UNIQUE = 100_000
GENERIC_NAMES = {"", "unknown", "unnamed", "n/a", "na", "none", "null"}

# These markers are used only for scope hygiene.  The People sheet remains the
# complete observation table; groups matching these rules are retained in the
# Excluded_Candidates sheet with the reason, rather than silently discarded.
SOUTH_ASIA_MARKERS = re.compile(
    r"\b(?:india|indian|mughal|mughals|timurid|timurids|hyderabad|pakistan|maratha|rajput|"
    r"sikh|kashmir|delhi|deccan|aurangzeb|shah jahan|jahanara begum|dara shikoh|"
    r"bahadur shah(?: zafar| i| ii)?|akbar shah|jahangir|mumtaz mahal)\b",
    re.I,
)
NON_PERSON_MARKERS = re.compile(
    r"\b(?:palace|university|college|royal guard|airport|island|museum|battle|wars?|treaty|"
    r"succession|family tree|list of|category|disambiguation|dynasty|sultanate|caliphate|"
    r"kingdom of|empire of|civil war)\b"
    r"|\b(?:al rayyan|abbasid samarra|samanids)\b"
    r"|آل ثاني|ولي عهد|ديوان|قصر|جامعة|حرس",
    re.I,
)
QID_ONLY = re.compile(r"^Q\d+$", re.I)


def exclusion_reason(names: list[str], urls: list[str], context: list[str] | None = None) -> str:
    """Return a conservative exclusion reason for an entity group, if any."""
    name_text = "；".join(n for n in names if n)
    url_text = " ".join(unquote(u) for u in urls if u)
    context_text = "；".join(v for v in (context or []) if v)
    if SOUTH_ASIA_MARKERS.search(name_text + " " + url_text + " " + context_text):
        return "范围排除：印度/巴基斯坦/南亚相关人物或页面；不纳入 GCC/中东核心名录"
    if names and all(QID_ONLY.fullmatch(norm(n).replace(" ", "")) for n in names if n):
        return "人物姓名未解析：仅有 Wikidata QID，暂不计入 unique people"
    if NON_PERSON_MARKERS.search(name_text + " " + url_text):
        return "非人物页面：事件、地点、机构、制度、王朝/类别或其他百科索引命中"
    return ""


def norm(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).lower().strip()
    text = re.sub(r"[\u200b\u200c\u200d]", "", text)
    text = re.sub(
        r"\b(h\.?h\.?|h\.?r\.?h\.?|sheikh|sheikha|shaikh|prince|princess|king|queen|emir|emira|dr\.?|mr\.?|ms\.?)\b",
        " ",
        text,
    )
    text = re.sub(r"[^a-z0-9\u0600-\u06ff]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def qid_from_url(url: str) -> str:
    match = re.search(r"wikidata\.org/(?:wiki/|entity/)(Q\d+)", url)
    return match.group(1) if match else ""


def wiki_key(url: str) -> str:
    match = re.search(r"https?://(?:www\.)?([a-z]{2,3})\.wikipedia\.org/wiki/([^?#]+)", url)
    return f"{match.group(1).lower()}|{match.group(2).lower()}" if match else ""


def datarabia_key(url: str) -> str:
    match = re.search(r"datarabia\.com/royals/famtree\.do\?id=([^&#]+)", url)
    return match.group(1) if match else ""


def identity_keys(row: tuple[object, ...]) -> tuple[list[str], bool]:
    name = norm(row[4])
    country = str(row[1] or "").strip()
    house = norm(row[2])
    parent = norm(row[18]) if len(row) > 18 else ""
    url = str(row[14] or "").strip()
    strong_keys: list[str] = []
    qid = qid_from_url(url)
    if qid:
        strong_keys.append(f"QID:{qid}")
    page = wiki_key(url)
    if page:
        strong_keys.append(f"WIKI:{page}")
    datarabia = datarabia_key(url)
    if datarabia:
        strong_keys.append(f"DATARABIA:{datarabia}")
    # Do not combine weak name keys with a strong page/QID key.  That creates
    # chain merges where a shared transliteration joins different people.
    if strong_keys:
        return strong_keys, False
    keys: list[str] = []
    if name not in GENERIC_NAMES:
        # A repeated royal name can belong to different generations.  When a
        # public parent field exists, keep it in the identity key so cousins
        # and same-name descendants are not collapsed into one entity.
        if parent:
            keys.append(f"NAMEP:{name}|{country}|{house}|{parent}")
        else:
            keys.append(f"NAMEN:{name}|{country}|{house}")
    unresolved = not keys or name in GENERIC_NAMES
    return keys, unresolved


class DisjointSet:
    def __init__(self, size: int) -> None:
        self.parent = list(range(size))

    def find(self, item: int) -> int:
        parent = self.parent
        while parent[item] != item:
            parent[item] = parent[parent[item]]
            item = parent[item]
        return item

    def union(self, left: int, right: int) -> None:
        left, right = self.find(left), self.find(right)
        if left != right:
            self.parent[right] = left


def clipped(values: list[str], limit: int = 30_000) -> str:
    unique = list(dict.fromkeys(v for v in values if v))
    text = "；".join(unique)
    if len(text) <= limit:
        return text
    kept: list[str] = []
    size = 0
    for value in unique:
        extra = len(value) + (1 if kept else 0)
        if size + extra + 30 > limit:
            break
        kept.append(value)
        size += extra
    return "；".join(kept) + f"；…（另有 {len(unique) - len(kept)} 项，见 People 原始观察表）"


def choose_longest(values: list[str]) -> str:
    values = [str(v).strip() for v in values if str(v or "").strip()]
    return max(values, key=len) if values else ""


def main() -> None:
    read_wb = load_workbook(WORKBOOK, read_only=True, data_only=True)
    people_rows = list(read_wb["People"].iter_rows(min_row=2, values_only=True))
    read_wb.close()

    dsu = DisjointSet(len(people_rows))
    first_key: dict[str, int] = {}
    row_keys: list[list[str]] = []
    unresolved_rows = 0
    for index, row in enumerate(people_rows):
        keys, unresolved = identity_keys(row)
        row_keys.append(keys)
        unresolved_rows += int(unresolved)
        for key in keys:
            previous = first_key.get(key)
            if previous is None:
                first_key[key] = index
            else:
                dsu.union(index, previous)

    groups: dict[int, list[int]] = defaultdict(list)
    for index in range(len(people_rows)):
        groups[dsu.find(index)].append(index)

    records: list[list[object]] = []
    excluded_records: list[list[object]] = []
    gcc_entities = 0
    unresolved_entities = 0
    match_basis_counts = Counter()
    for indexes in groups.values():
        usable_indexes = [i for i in indexes if norm(people_rows[i][4]) not in GENERIC_NAMES or qid_from_url(str(people_rows[i][14] or ""))]
        if not usable_indexes:
            unresolved_entities += 1
            continue
        keys = sorted({key for i in indexes for key in row_keys[i]})
        names = [str(people_rows[i][4] or "").strip() for i in usable_indexes]
        countries = Counter(str(people_rows[i][1] or "").strip() for i in usable_indexes if people_rows[i][1])
        houses = Counter(str(people_rows[i][2] or "").strip() for i in usable_indexes if people_rows[i][2])
        sources = sorted({str(people_rows[i][13] or "").strip() for i in indexes if people_rows[i][13]})
        urls = sorted({str(people_rows[i][14] or "").strip() for i in indexes if people_rows[i][14]})
        observation_ids = [str(people_rows[i][0] or "").strip() for i in indexes if people_rows[i][0]]
        parent_values = [str(people_rows[i][18] or "").strip() for i in indexes if len(people_rows[i]) > 18]
        house_values_for_scope = [str(people_rows[i][2] or "").strip() for i in indexes if people_rows[i][2]]
        reason = exclusion_reason(names, urls, parent_values + house_values_for_scope)
        if reason:
            excluded_records.append([
                "EXC-" + hashlib.sha1("|".join(keys or observation_ids).encode("utf-8")).hexdigest()[:16].upper(),
                choose_longest(names),
                clipped(sorted(set(names)), 10_000),
                "；".join(str(people_rows[i][1] or "").strip() for i in indexes if people_rows[i][1]),
                "；".join(sorted({str(people_rows[i][2] or "").strip() for i in indexes if people_rows[i][2]})),
                reason,
                len(indexes),
                "；".join(sources),
                clipped(urls, 30_000),
                "；".join(str(people_rows[i][0] or "").strip() for i in indexes if people_rows[i][0]),
            ])
            continue
        relation_text = list(dict.fromkeys(str(people_rows[i][8] or "").strip() for i in indexes if people_rows[i][8]))
        relation_types = sorted({str(people_rows[i][9] or "").strip() for i in indexes if people_rows[i][9]})
        intros = [str(people_rows[i][7] or "").strip() for i in indexes if people_rows[i][7]]
        statuses = Counter(str(people_rows[i][10] or "").strip() for i in indexes if people_rows[i][10])
        country_values = [country for country, _ in countries.most_common()]
        house_values = [house for house, _ in houses.most_common()]
        qid_keys = [key for key in keys if key.startswith("QID:")]
        wiki_keys = [key for key in keys if key.startswith("WIKI:")]
        datarabia_keys = [key for key in keys if key.startswith("DATARABIA:")]
        name_keys = [key for key in keys if key.startswith("NAME:")]
        if qid_keys and (wiki_keys or datarabia_keys):
            basis = "Wikidata QID with public page cross-reference"
            confidence = "B1"
        elif qid_keys:
            basis = "Wikidata QID"
            confidence = "B2"
        elif wiki_keys or datarabia_keys:
            basis = "Canonical public page"
            confidence = "B2"
        else:
            basis = "Normalized name+country+house"
            confidence = "C"
        if len(country_values) > 1 or len(house_values) > 1:
            confidence += " review"
        match_basis_counts[basis] += 1
        entity_key = keys[0] if keys else f"OBS:{observation_ids[0]}"
        entity_id = "ENT-" + hashlib.sha1("|".join(keys or [entity_key]).encode("utf-8")).hexdigest()[:16].upper()
        intro = choose_longest(intros)
        relationship = clipped(relation_text, 30_000)
        observation_text = clipped(observation_ids, 30_000)
        notes = f"由 {len(indexes):,} 条 People 观察合并；原始观察保留在 People。"
        if len(observation_ids) < len(indexes):
            notes += " 部分原始观察缺少 Person_ID。"
        if confidence.endswith(" review"):
            notes += " 国家或家族字段存在多值，需人工消歧。"
        row = [
            entity_id,
            choose_longest(names),
            clipped(sorted(set(names)), 10_000),
            "；".join(country_values),
            "；".join(house_values),
            basis,
            clipped(keys, 10_000),
            confidence,
            statuses.most_common(1)[0][0] if statuses else "",
            intro,
            relationship,
            "；".join(relation_types),
            len(indexes),
            len(sources),
            "；".join(sources),
            clipped(urls, 30_000),
            observation_text,
            max(str(people_rows[i][15] or "") for i in indexes),
            choose_longest([str(people_rows[i][16] or "") for i in indexes]),
            notes,
        ]
        records.append(row)
        if set(country_values) and set(country_values) <= GCC:
            gcc_entities += 1

    records.sort(key=lambda row: (str(row[3]), str(row[4]), str(row[1]), str(row[0])))

    wb = load_workbook(WORKBOOK)
    if "Unique_People" in wb.sheetnames:
        del wb["Unique_People"]
    unique = wb.create_sheet("Unique_People", 1)
    unique.sheet_view.showGridLines = False
    unique.freeze_panes = "A2"
    headers = [
        "Entity_ID", "主姓名", "别名/观测姓名", "国家/酋长国", "统治家族/支系",
        "身份匹配依据", "实体键/QID", "匹配置信度", "主要记录状态", "人物简介",
        "关系摘要", "关系类型", "观察记录数", "来源数", "来源ID", "来源URL",
        "原始Observation_ID", "截至日期", "下一步核验", "去重说明",
    ]
    unique.append(headers)
    for row in records:
        unique.append(row)
    unique.column_dimensions["A"].width = 20
    for col in ["B", "C", "D", "E", "F", "H", "I", "L", "M", "N", "O", "R"]:
        unique.column_dimensions[col].width = 22
    for col in ["G", "P", "Q", "S", "T"]:
        unique.column_dimensions[col].width = 34
    unique.column_dimensions["J"].width = 52
    unique.column_dimensions["K"].width = 60
    unique.row_dimensions[1].height = 32
    for cell in unique[1]:
        template_cell = wb["People"][1][min(cell.column - 1, wb["People"].max_column - 1)]
        cell.font = copy(template_cell.font)
        cell.fill = copy(template_cell.fill)
        cell.border = copy(template_cell.border)
        cell.alignment = copy(template_cell.alignment)
    ref = f"A1:T{unique.max_row}"
    table = Table(displayName="UniquePeopleTable", ref=ref)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    unique.add_table(table)

    if "Excluded_Candidates" in wb.sheetnames:
        del wb["Excluded_Candidates"]
    excluded = wb.create_sheet("Excluded_Candidates", 2)
    excluded.sheet_view.showGridLines = False
    excluded.freeze_panes = "A2"
    excluded_headers = [
        "Exclusion_ID", "主姓名", "别名/观测姓名", "国家/酋长国", "统治家族/支系",
        "排除原因", "观察记录数", "来源ID", "来源URL", "原始Observation_ID",
    ]
    excluded.append(excluded_headers)
    for row in sorted(excluded_records, key=lambda row: (str(row[3]), str(row[1]), str(row[0]))):
        excluded.append(row)
    for col in ["A", "B", "D", "E", "G", "H"]:
        excluded.column_dimensions[col].width = 22
    for col in ["C", "F", "I", "J"]:
        excluded.column_dimensions[col].width = 40
    excluded.row_dimensions[1].height = 32
    for cell in excluded[1]:
        template_cell = wb["People"][1][min(cell.column - 1, wb["People"].max_column - 1)]
        cell.font = copy(template_cell.font)
        cell.fill = copy(template_cell.fill)
        cell.border = copy(template_cell.border)
        cell.alignment = copy(template_cell.alignment)
    excluded_ref = f"A1:J{excluded.max_row}"
    excluded_table = Table(displayName="ExcludedCandidatesTable", ref=excluded_ref)
    excluded_table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    excluded.add_table(excluded_table)

    dashboard = wb["Dashboard"]
    template_row = min(31, dashboard.max_row)
    metrics = [
        (33, "唯一人物/实体候选（保守去重）", len(records), "observed unique", "按 QID+姓名、公开页面或姓名+国家+家族合并，不等于已确认在世王室成员"),
        (34, "唯一 GCC 人物/实体候选", gcc_entities, "observed unique", "严格 GCC 视图：国家字段全部属于 GCC 六国；跨入非 GCC 历史/亲缘范围者不计入"),
        (35, "距离 100,000 个唯一人物的缺口", max(TARGET_UNIQUE - len(records), 0), "gap", "当前公开记录不足时不以重复关系补足"),
        (36, "观察记录 / 唯一实体", round((dashboard["B4"].value or 0) / len(records), 2) if records else 0, "ratio", "去重膨胀倍数；用于判断关系观察重复度"),
        (37, "排除候选（非人物/南亚范围）", len(excluded_records), "excluded", "原始观察仍保留在 People；排除理由和来源保留在 Excluded_Candidates"),
    ]
    for row_idx, label, value, basis, note in metrics:
        if row_idx > dashboard.max_row:
            for col in range(1, dashboard.max_column + 1):
                src, dst = dashboard.cell(template_row, col), dashboard.cell(row_idx, col)
                if src.has_style:
                    dst._style = copy(src._style)
                if src.alignment:
                    dst.alignment = copy(src.alignment)
        dashboard.cell(row_idx, 1).value = label
        dashboard.cell(row_idx, 2).value = value
        dashboard.cell(row_idx, 4).value = basis
        dashboard.cell(row_idx, 5).value = note
    dashboard["B36"].number_format = "0.00"

    wb["ReadMe"]["B5"] = (
        f"People 表目前 {dashboard['B4'].value:,} 条观察记录；Unique_People 当前 {len(records):,} 个保守去重实体候选，"
        f"其中严格 GCC {gcc_entities:,} 个，距离 100,000 个唯一人物目标还差 {max(TARGET_UNIQUE - len(records), 0):,} 个。"
        f"本轮另排除 {len(excluded_records):,} 个明显非人物页面、未解析 QID 或印度/巴基斯坦/南亚范围候选；排除明细保留在 Excluded_Candidates。"
        "有 QID、Wikipedia 页面或 Datarabia 节点时使用强来源键；没有强键时使用姓名+国家+家族，若有公开父/母字段则一并用于区分同名不同代。"
        "原始观察不删除，重复关系不计作新人。"
    )
    wb.save(WORKBOOK)
    print({
        "people_observations": len(people_rows),
        "unique_entities": len(records),
        "unique_gcc_entities": gcc_entities,
        "unique_gap_to_100k": max(TARGET_UNIQUE - len(records), 0),
        "unresolved_observation_rows": unresolved_rows,
        "unresolved_groups": unresolved_entities,
        "excluded_candidates": len(excluded_records),
        "match_basis": dict(match_basis_counts),
        "output": str(WORKBOOK),
    })


if __name__ == "__main__":
    main()
