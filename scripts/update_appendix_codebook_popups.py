#!/usr/bin/env python3
"""Add exact thesis-codebook metadata to the canonical Appendix A3 node popups."""

from __future__ import annotations

import argparse
import html
import json
import re
from pathlib import Path

from docx import Document


PAGES = {
    "chapter-5-conceptualisation.html": "C4",
    "chapter-6-operationalisation.html": "C5",
    "chapter-7-experience.html": "C6",
    "chapter-8-cross-dataset-authority.html": "ALL",
}

EXPECTED_COUNTS = {"C4": 34, "C5": 33, "C6": 33}

PROPOSAL_DATASETS = {
    "C4": "chapter_5_conceptualisation",
    "C5": "chapter_6_operationalisation",
    "C6": "chapter_7_experience",
}

PUBLISHED_CHAPTER = {"C4": "C5", "C5": "C6", "C6": "C7"}

ANALYTIC_HOME_COUNT_FIELDS = {
    "C4": (
        ("Obedience", "count__C4__Obedience"),
        ("Internalisation", "count__C4__Internalisation"),
        ("Performativity", "count__C4__Performativity"),
    ),
    "C5": (
        ("Using control to influence ‘hearts and minds’", "count__C5__Hearts.and.minds"),
        ("Using control to govern the individual body", "count__C5__Individual.body"),
        ("Using control to tune the group", "count__C5__Group.order"),
    ),
    "C6": (
        ("Disconnecting from previous life", "count__C6__Breaking"),
        ("Reconnecting inside the programme", "count__C6__Belonging"),
        ("Developing within a precarious situation", "count__C6__Becoming"),
    ),
}

TOOLTIP_STYLE = (
    "position:fixed;visibility:hidden;padding:8px;white-space:normal;"
    "font-family:verdana;font-size:14px;color:#000000;background-color:#f5f4ed;"
    "border-radius:3px;border:1px solid #808074;"
    "box-shadow:3px 3px 10px rgba(0,0,0,0.2);"
    "max-width:calc(100vw - 24px);max-height:calc(100vh - 24px);"
    "overflow:auto;overflow-wrap:anywhere;box-sizing:border-box;z-index:1000;"
)

SCRIPT_RE = re.compile(
    r'(<script type="application/json"[^>]*>)(.*?)(</script>)', re.DOTALL
)
METADATA_DIV_RE = re.compile(
    r"(<div style='margin-bottom:(?:9|10)px;'>)(.*?)(</div>)", re.DOTALL
)
EXISTING_CODEBOOK_RE = re.compile(
    r"<b>Analytic home:</b>.*?<br>"
    r"<b>Definition:</b>.*?<br>"
    r"<b>Include:</b>.*?<br>"
    r"<b>Exclude:</b>.*?<br>",
    re.DOTALL,
)
DOMINANT_RE = re.compile(
    r"<b>Dominant (?:theme|operational route|chapter theme):</b>.*?<br>", re.DOTALL
)
POPUP_HEADING_RE = re.compile(
    r"(<div style='font-size:15px;font-weight:700;margin-bottom:(?:7|8)px;'>)"
    r"(.*?)"
    r"(</div>)",
    re.DOTALL,
)

LEGACY_EVIDENCE_FIELDS = (
    "Authority-recognition relation",
    "Recognition problem",
    "Relational stabiliser",
    "Source-tracing test",
)

AUTHORITY_HUB_CONNECTIONS = {
    "C4": (
        "hierarchy_authority",
        "authority_legitimacy_intent",
        "authority_maintenance",
    ),
    "C5": (
        "standard_setting",
        "accountability_orientation",
        "explanatory_reasoning",
        "relational_influence",
        "visibility_supervision",
        "group_standard_imposition",
    ),
    "C6": ("authority_disregard",),
}

POPUP_LABEL_REPLACEMENTS = {
    "<b>Total coded clusters:</b>": "<b>Coded clusters:</b>",
    "<b>Evidence links represented:</b>": "<b>Linked evidence extracts:</b>",
    "<b>Theme frequencies:</b>": "<b>Coded clusters by analytic home:</b>",
    "<b>Route frequencies:</b>": "<b>Coded clusters by analytic home:</b>",
    "<b>Bomb Night evidence links:</b>": "<b>Bomb Night-linked evidence extracts:</b>",
    "<b>Cluster:</b>": "<b>Cluster ID:</b>",
}


def clean_cell(value: str) -> str:
    return " ".join(value.split())


def strip_field_prefix(value: str, field: str) -> str:
    return re.sub(rf"^{re.escape(field)}\s*:?\s*", "", value, flags=re.IGNORECASE)


def extract_codebooks(docx_path: Path) -> dict[str, dict[str, dict[str, str]]]:
    document = Document(docx_path)
    found: list[dict[str, dict[str, str]]] = []

    for table in document.tables:
        if not table.rows or len(table.columns) != 5:
            continue
        header = [clean_cell(cell.text) for cell in table.rows[0].cells]
        if not (
            len(header) == 5
            and header[0] == "Code"
            and header[1] == "Analytic home"
            and header[2].startswith("Definition")
            and header[3].startswith("Include")
            and header[4].startswith("Exclude")
        ):
            continue

        codebook: dict[str, dict[str, str]] = {}
        for row in table.rows[1:]:
            cells = [clean_cell(cell.text) for cell in row.cells]
            if not cells[0]:
                continue
            code, analytic_home, definition, include, exclude = cells
            if code in codebook:
                raise ValueError(f"Duplicate codebook code: {code}")
            codebook[code] = {
                "analytic_home": analytic_home,
                "definition": definition,
                "include": strip_field_prefix(include, "Include"),
                "exclude": strip_field_prefix(exclude, "Exclude"),
            }
        found.append(codebook)

    if [len(book) for book in found] != [34, 33, 33]:
        raise ValueError(
            "Expected the three Appendix 3B codebooks with 34, 33 and 33 rows; "
            f"found {[len(book) for book in found]}"
        )

    return {"C4": found[0], "C5": found[1], "C6": found[2]}


def load_proposals(
    proposals_path: Path,
    codebooks: dict[str, dict[str, dict[str, str]]],
) -> dict[str, dict[str, dict[str, str]]]:
    raw = json.loads(proposals_path.read_text(encoding="utf-8"))
    mapped: dict[str, dict[str, dict[str, str]]] = {}

    for group, dataset_name in PROPOSAL_DATASETS.items():
        source = raw[dataset_name]
        mapped[group] = {}
        for legacy_id, proposal in source.items():
            current_id = proposal["id"]
            if current_id not in codebooks[group]:
                raise ValueError(
                    f"{group}: proposed identifier {current_id} is absent from the "
                    "authoritative codebook"
                )
            metadata = codebooks[group][current_id]
            for key in ("definition", "include", "exclude"):
                proposal_value = proposal[key]
                if key in ("include", "exclude"):
                    proposal_value = strip_field_prefix(proposal_value, key)
                if clean_cell(proposal_value) != clean_cell(metadata[key]):
                    raise ValueError(
                        f"{group}: {legacy_id} {key} differs between the approved "
                        "proposal record and the authoritative codebook"
                    )
            mapped[group][legacy_id] = {
                **metadata,
                "legacy_id": legacy_id,
                "current_id": current_id,
                "name": proposal["name"],
            }

        if len(mapped[group]) != EXPECTED_COUNTS[group]:
            raise ValueError(
                f"{group} contains {len(mapped[group])} proposal mappings, "
                f"expected {EXPECTED_COUNTS[group]}"
            )

    return mapped


def field_block(metadata: dict[str, str]) -> str:
    def esc(key: str) -> str:
        return html.escape(metadata[key], quote=False)

    return (
        f"<b>Analytic home:</b> {esc('analytic_home')}<br>"
        f"<b>Definition:</b> {esc('definition')}<br>"
        f"<b>Include:</b> {esc('include')}<br>"
        f"<b>Exclude:</b> {esc('exclude')}<br>"
    )


def update_title(title: str, metadata: dict[str, str]) -> str:
    match = METADATA_DIV_RE.search(title)
    if not match:
        raise ValueError("Could not locate the opening metadata block in a node popup")

    heading = POPUP_HEADING_RE.search(title)
    if heading:
        title = (
            title[: heading.start()]
            + heading.group(1)
            + html.escape(metadata["name"], quote=False)
            + heading.group(3)
            + title[heading.end() :]
        )
        match = METADATA_DIV_RE.search(title)
        if not match:
            raise ValueError("Popup metadata block disappeared after heading update")

    body = match.group(2)
    body = EXISTING_CODEBOOK_RE.sub("", body, count=1)
    body = DOMINANT_RE.sub("", body, count=1)
    replacement = match.group(1) + field_block(metadata) + body + match.group(3)
    revised = title[: match.start()] + replacement + title[match.end() :]

    # These four fields came from an earlier interpretive authority template,
    # not from the approved codebook or the empirical evidence inventory.
    # Removing them keeps each popup's hierarchy explicit: codebook boundary,
    # then participant evidence, without a second theory-laden classification.
    for field in LEGACY_EVIDENCE_FIELDS:
        revised = re.sub(
            rf"<br><b>{re.escape(field)}:</b>.*?(?=<br>|</div>)",
            "",
            revised,
            flags=re.DOTALL,
        )
    revised = re.sub(
        r"<br><b>Theme function:</b>.*?(?=<br>|</div>)",
        "",
        revised,
        flags=re.DOTALL,
    )
    for old_label, new_label in POPUP_LABEL_REPLACEMENTS.items():
        revised = revised.replace(old_label, new_label)
    return revised


def authority_hub_title(
    mappings: dict[str, dict[str, dict[str, str]]],
) -> str:
    items: list[str] = []
    for group, legacy_ids in AUTHORITY_HUB_CONNECTIONS.items():
        chapter = PUBLISHED_CHAPTER[group][1:]
        for legacy_id in legacy_ids:
            name = html.escape(mappings[group][legacy_id]["name"], quote=False)
            items.append(f"<li>Chapter {chapter}: {name}</li>")

    return (
        "<div style='max-width:620px;font-family:Arial, sans-serif;"
        "font-size:12px;line-height:1.40;'>"
        "<div style='font-size:15px;font-weight:700;margin-bottom:8px;'>"
        "Authority</div>"
        "<div style='margin-bottom:10px;'>"
        "<b>Status:</b> Interpretive hub, not a coded node.<br>"
        "<b>Function:</b> Connects authority-related patterns across "
        "conceptualisation, operationalisation and experience in Chapters 5, 6 and 7.<br>"
        "<b>Includes:</b> The authority-related code nodes listed below.<br>"
        "<b>Does not show:</b> A raw ENA code, independently coded evidence, "
        "frequency, prevalence or causal direction."
        "</div>"
        "<div style='font-weight:700;margin-bottom:6px;'>"
        "Connected authority-related nodes</div>"
        "<ul style='padding-left:18px;margin-top:0;'>"
        + "".join(items)
        + "</ul></div>"
    )


def add_cross_dataset_analytic_counts(
    title: str,
    nodes: dict,
    node_index: int,
    codebook_group: str,
) -> str:
    parts = [
        f"{label} = {nodes[field][node_index]}"
        for label, field in ANALYTIC_HOME_COUNT_FIELDS[codebook_group]
    ]
    field = "<b>Coded clusters by analytic home:</b> " + "; ".join(parts)
    if "<b>Coded clusters by analytic home:</b>" in title:
        return re.sub(
            r"<b>Coded clusters by analytic home:</b>.*?(?=<br>|</div>)",
            field,
            title,
            count=1,
            flags=re.DOTALL,
        )
    revised, count = re.subn(
        r"(<b>Linked evidence extracts:</b>[^<]*)(</div>)",
        rf"\1<br>{field}\2",
        title,
        count=1,
    )
    if count != 1:
        raise ValueError("Could not add cross-dataset analytic-home counts")
    return revised


def resolve_metadata(
    node_id: str,
    page_group: str,
    mappings: dict[str, dict[str, dict[str, str]]],
) -> dict[str, str]:
    if node_id == "AUTHORITY_HUB":
        return {
            "name": "Authority",
            "legacy_id": "AUTHORITY_HUB",
            "current_id": "AUTHORITY_HUB",
            "codebook_group": "ALL",
            "analytic_home": "Cross-dataset synthesis",
            "definition": (
                "Synthetic integrative device connecting conceptualisation, "
                "operationalisation and experience across Chapters 5, 6 and 7."
            ),
            "include": "The authority-related code nodes listed below.",
            "exclude": "Interpretation as a raw ENA code or independently coded evidence.",
        }

    if page_group == "ALL":
        match = re.fullmatch(r"(C[456])__(.+)", node_id)
        if not match:
            raise ValueError(f"Unrecognised Chapter 8 node identifier: {node_id}")
        codebook_group, legacy_id = match.groups()
    else:
        codebook_group, legacy_id = page_group, node_id

    try:
        return {**mappings[codebook_group][legacy_id], "codebook_group": codebook_group}
    except KeyError as error:
        raise ValueError(
            f"No {codebook_group} codebook row found for node {node_id}"
        ) from error


def wrap_label(value: str, width: int = 24) -> str:
    words = value.split()
    lines: list[str] = []
    line = ""
    for word in words:
        candidate = f"{line} {word}".strip()
        if line and len(candidate) > width:
            lines.append(line)
            line = word
        else:
            line = candidate
    if line:
        lines.append(line)
    return "\n".join(lines)


def node_label(metadata: dict[str, str], page_group: str) -> str:
    label = wrap_label(metadata["name"])
    if page_group == "ALL" and metadata["current_id"] != "AUTHORITY_HUB":
        return f"{PUBLISHED_CHAPTER[metadata['codebook_group']]}: {label}"
    return label


def update_edge_titles(
    payload: dict,
    page_group: str,
    mappings: dict[str, dict[str, dict[str, str]]],
) -> None:
    edges = payload["x"]["edges"]
    from_ids = edges["from"]
    to_ids = edges["to"]
    titles = edges.get("title", [])
    if len(from_ids) != len(to_ids) or len(from_ids) != len(titles):
        raise ValueError("Edge endpoint/title count mismatch")

    revised_titles: list[str] = []
    for from_id, to_id, title in zip(from_ids, to_ids, titles):
        from_name = resolve_metadata(from_id, page_group, mappings)["name"]
        to_name = resolve_metadata(to_id, page_group, mappings)["name"]
        heading = (
            f"<b>{html.escape(from_name, quote=False)} ↔ "
            f"{html.escape(to_name, quote=False)}</b>"
        )
        revised, count = re.subn(r"<b>.*?</b>", heading, title, count=1, flags=re.DOTALL)
        if count != 1:
            raise ValueError(f"Could not update edge title {from_id} ↔ {to_id}")
        revised_titles.append(revised)
    edges["title"] = revised_titles


def update_page(
    page_path: Path,
    page_group: str,
    mappings: dict[str, dict[str, dict[str, str]]],
) -> int:
    source = page_path.read_text(encoding="utf-8")
    script_match = SCRIPT_RE.search(source)
    if not script_match:
        raise ValueError(f"No htmlwidget JSON payload found in {page_path}")

    payload = json.loads(script_match.group(2))
    nodes = payload["x"]["nodes"]
    node_ids = nodes["id"]
    titles = nodes["title"]
    labels = nodes["label"]
    if len(node_ids) != len(titles) or len(node_ids) != len(labels):
        raise ValueError(f"Node/title count mismatch in {page_path}")

    revised_titles = []
    revised_labels = []
    for node_index, (node_id, title) in enumerate(zip(node_ids, titles)):
        metadata = resolve_metadata(node_id, page_group, mappings)
        revised = (
            authority_hub_title(mappings)
            if node_id == "AUTHORITY_HUB"
            else update_title(title, metadata)
        )
        if page_group == "ALL" and node_id != "AUTHORITY_HUB":
            revised = add_cross_dataset_analytic_counts(
                revised,
                nodes,
                node_index,
                metadata["codebook_group"],
            )
        if node_id == "AUTHORITY_HUB":
            revised_titles.append(revised)
            revised_labels.append(node_label(metadata, page_group))
            continue
        for label in ("Analytic home", "Definition", "Include", "Exclude"):
            if revised.count(f"<b>{label}:</b>") != 1:
                raise ValueError(
                    f"{page_path.name}: {node_id} does not contain exactly one {label} field"
                )
        if DOMINANT_RE.search(revised):
            raise ValueError(f"{page_path.name}: dominant-theme wording remains in {node_id}")
        for legacy_field in LEGACY_EVIDENCE_FIELDS:
            if f"<b>{legacy_field}:</b>" in revised:
                raise ValueError(
                    f"{page_path.name}: legacy field {legacy_field} remains in {node_id}"
                )
        if "<b>Theme function:</b>" in revised:
            raise ValueError(f"{page_path.name}: Theme function remains in {node_id}")
        for old_label in POPUP_LABEL_REPLACEMENTS:
            if old_label in revised:
                raise ValueError(
                    f"{page_path.name}: non-standard popup label remains in {node_id}: "
                    f"{old_label}"
                )
        revised_titles.append(revised)
        revised_labels.append(node_label(metadata, page_group))

    nodes["title"] = revised_titles
    nodes["label"] = revised_labels
    update_edge_titles(payload, page_group, mappings)
    payload["x"]["tooltipStyle"] = TOOLTIP_STYLE
    # Keep every published key independently zoomable. This is essential on
    # narrower screens, where the longer Chapter 6 and Chapter 7 theme labels
    # otherwise appear clipped inside the legend region.
    payload["x"]["legend"]["zoom"] = True
    revised_payload = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    revised_source = (
        source[: script_match.start(2)]
        + revised_payload
        + source[script_match.end(2) :]
    )
    page_path.write_text(revised_source, encoding="utf-8")
    return len(node_ids)


def write_metadata_audit(
    output_path: Path, mappings: dict[str, dict[str, dict[str, str]]]
) -> None:
    audit = {
        "source": "Appendix 3B codebook tables in the authoritative thesis DOCX",
        "datasets": {
            "chapter_5_conceptualisation": mappings["C4"],
            "chapter_6_operationalisation": mappings["C5"],
            "chapter_7_experience": mappings["C6"],
        },
        "chapter_8_non_code_node": {
            "AUTHORITY_HUB": resolve_metadata("AUTHORITY_HUB", "ALL", mappings)
        },
    }
    output_path.write_text(
        json.dumps(audit, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--proposals", required=True, type=Path)
    parser.add_argument("--repo-root", required=True, type=Path)
    args = parser.parse_args()

    codebooks = extract_codebooks(args.docx)
    for group, expected in EXPECTED_COUNTS.items():
        if len(codebooks[group]) != expected:
            raise ValueError(f"{group} contains {len(codebooks[group])} rows, expected {expected}")

    mappings = load_proposals(args.proposals, codebooks)

    appendix_dir = args.repo_root / "appendix-networks"
    updated: dict[str, int] = {}
    for filename, page_group in PAGES.items():
        updated[filename] = update_page(
            appendix_dir / filename, page_group, mappings
        )

    write_metadata_audit(appendix_dir / "codebook-popup-metadata.json", mappings)
    print(json.dumps({"updated_nodes": updated}, indent=2))


if __name__ == "__main__":
    main()
