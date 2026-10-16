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


def clean_cell(value: str) -> str:
    return " ".join(value.split())


def strip_field_prefix(value: str, field: str) -> str:
    return re.sub(rf"^{re.escape(field)}\s*:\s*", "", value, flags=re.IGNORECASE)


def extract_codebooks(docx_path: Path) -> dict[str, dict[str, dict[str, str]]]:
    document = Document(docx_path)
    found: list[dict[str, dict[str, str]]] = []

    for table in document.tables:
        if not table.rows or len(table.columns) != 5:
            continue
        header = [clean_cell(cell.text) for cell in table.rows[0].cells]
        if header != ["Code", "Analytic home", "Definition", "Include", "Exclude"]:
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
            "Expected the three Appendix 4 codebooks with 34, 33 and 33 rows; "
            f"found {[len(book) for book in found]}"
        )

    return {"C4": found[0], "C5": found[1], "C6": found[2]}


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

    body = match.group(2)
    body = EXISTING_CODEBOOK_RE.sub("", body, count=1)
    body = DOMINANT_RE.sub("", body, count=1)
    replacement = match.group(1) + field_block(metadata) + body + match.group(3)
    return title[: match.start()] + replacement + title[match.end() :]


def resolve_metadata(
    node_id: str,
    page_group: str,
    codebooks: dict[str, dict[str, dict[str, str]]],
) -> dict[str, str]:
    if node_id == "AUTHORITY_HUB":
        return {
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
        codebook_group, code = match.groups()
    else:
        codebook_group, code = page_group, node_id

    try:
        return codebooks[codebook_group][code]
    except KeyError as error:
        raise ValueError(
            f"No {codebook_group} codebook row found for node {node_id}"
        ) from error


def update_page(
    page_path: Path,
    page_group: str,
    codebooks: dict[str, dict[str, dict[str, str]]],
) -> int:
    source = page_path.read_text(encoding="utf-8")
    script_match = SCRIPT_RE.search(source)
    if not script_match:
        raise ValueError(f"No htmlwidget JSON payload found in {page_path}")

    payload = json.loads(script_match.group(2))
    nodes = payload["x"]["nodes"]
    node_ids = nodes["id"]
    titles = nodes["title"]
    if len(node_ids) != len(titles):
        raise ValueError(f"Node/title count mismatch in {page_path}")

    revised_titles = []
    for node_id, title in zip(node_ids, titles):
        metadata = resolve_metadata(node_id, page_group, codebooks)
        revised = update_title(title, metadata)
        for label in ("Analytic home", "Definition", "Include", "Exclude"):
            if revised.count(f"<b>{label}:</b>") != 1:
                raise ValueError(
                    f"{page_path.name}: {node_id} does not contain exactly one {label} field"
                )
        if DOMINANT_RE.search(revised):
            raise ValueError(f"{page_path.name}: dominant-theme wording remains in {node_id}")
        revised_titles.append(revised)

    nodes["title"] = revised_titles
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
    output_path: Path, codebooks: dict[str, dict[str, dict[str, str]]]
) -> None:
    audit = {
        "source": "Appendix 4 codebook tables in the authoritative thesis DOCX",
        "datasets": {
            "chapter_5_conceptualisation": codebooks["C4"],
            "chapter_6_operationalisation": codebooks["C5"],
            "chapter_7_experience": codebooks["C6"],
        },
        "chapter_8_non_code_node": {
            "AUTHORITY_HUB": resolve_metadata("AUTHORITY_HUB", "ALL", codebooks)
        },
    }
    output_path.write_text(
        json.dumps(audit, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--repo-root", required=True, type=Path)
    args = parser.parse_args()

    codebooks = extract_codebooks(args.docx)
    for group, expected in EXPECTED_COUNTS.items():
        if len(codebooks[group]) != expected:
            raise ValueError(f"{group} contains {len(codebooks[group])} rows, expected {expected}")

    appendix_dir = args.repo_root / "appendix-networks"
    updated: dict[str, int] = {}
    for filename, page_group in PAGES.items():
        updated[filename] = update_page(
            appendix_dir / filename, page_group, codebooks
        )

    write_metadata_audit(appendix_dir / "codebook-popup-metadata.json", codebooks)
    print(json.dumps({"updated_nodes": updated}, indent=2))


if __name__ == "__main__":
    main()
