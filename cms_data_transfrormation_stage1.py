import ast
import difflib
import json
import re
import uuid

from datetime import date, datetime
from copy import copy
from pathlib import Path
from typing import List, Tuple, Union

import pandas as pd

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

if __package__:
    from .cms_content_utils import normalize_train_url
    from .train_output_english_french import (
        TRAIN_PARSED_COLUMNS,
        build_train_output,
        format_train_output_sheet,
    )
else:
    from cms_content_utils import normalize_train_url
    from train_output_english_french import (
        TRAIN_PARSED_COLUMNS,
        build_train_output,
        format_train_output_sheet,
    )


# ============================================================
# REGEX
# ============================================================

MARKDOWN_LINK_PATTERN = re.compile(
    r"\[([^\]]+)\]\((https?://[^)\s]+)\)"
)

BARE_URL_PATTERN = re.compile(
    r"https?://[^\s)\]>]+"
)


# ============================================================
# URL / CONTENT HELPERS
# ============================================================

def _clean_url_name_content(value: str) -> str:

    link_only_pattern = re.compile(
        r"^(?:\s*[-*+]\s*)?"
        r"(?:\[[^\]]+?\]\(\s*https?://[^)\s]+"
        r"(?:\s+[^)]*)?\s*\)\s*)+$"
    )

    if link_only_pattern.fullmatch(value):
        return "Null"

    def replace_markdown_link(match: re.Match[str]) -> str:
        return re.sub(
            r"\*+",
            "",
            match.group(1),
        ).strip()

    value = re.sub(
        r"\[([^\]]+?)\]\(\s*https?://[^)\s]+"
        r"(?:\s+[^)]*)?\s*\)",
        replace_markdown_link,
        value,
    )

    value = BARE_URL_PATTERN.sub(
        "",
        value,
    )

    value = re.sub(
        r"\*+",
        "",
        value,
    )

    value = re.sub(
        r"(?m)^\s*[-*+]\s+",
        "",
        value,
    )

    value = re.sub(
        r"\s+",
        " ",
        value,
    )

    return value.strip() or "Null"


def _context_for_url(
    content: str,
    matches: list[re.Match[str]],
    target_index: int,
) -> str:

    target = matches[target_index]

    next_start = (
        matches[target_index + 1].start()
        if target_index + 1 < len(matches)
        else len(content)
    )

    after_text = content[
        target.end():
        next_start
    ]

    after_context = _clean_url_name_content(
        after_text
    )

    if after_context != "Null":
        return after_context

    previous_end = (
        matches[target_index - 1].end()
        if target_index > 0
        else 0
    )

    before_context = _clean_url_name_content(
        content[
            previous_end:
            target.start()
        ]
    )

    return before_context


# ============================================================
# CONTENT SUB-HEADING HELPERS
# ============================================================

MARKDOWN_SUBHEADING_PATTERN = re.compile(r"(?m)^\s*#{2,6}\s+(.+?)\s*$")


def _clean_sub_heading(value: object) -> str:
    """Return display text for a Markdown sub-heading."""
    if value is None:
        return "Null"
    text = re.sub(r"\*+", "", str(value)).strip()
    text = re.sub(r"\s+#+\s*$", "", text).strip()
    # Keep punctuation such as ':' because it can be meaningful in the CMS text,
    # but remove a single trailing colon for cleaner translator metadata.
    text = text[:-1].rstrip() if text.endswith(":") else text
    return text or "Null"


def _subheading_for_position(content_text: str, position: int) -> str:
    """Find the nearest Markdown sub-heading preceding a URL/text position."""
    current = "Null"
    for match in MARKDOWN_SUBHEADING_PATTERN.finditer(content_text):
        if match.start() > position:
            break
        current = _clean_sub_heading(match.group(1))
    return current


def _split_content_by_subheading(raw_content: object) -> list[dict]:
    """Split one Strapi content item into ordered Markdown sub-heading sections."""
    text = "" if raw_content is None else str(raw_content)
    matches = list(MARKDOWN_SUBHEADING_PATTERN.finditer(text))
    if not matches:
        return [{"sub_heading": "Null", "raw_content": text,
                 "clean_content": _content_without_urls(text)}]
    sections = []
    # Text before the first Markdown sub-heading belongs to the parent heading.
    if text[:matches[0].start()].strip():
        prefix = text[:matches[0].start()]
        sections.append({"sub_heading": "Null", "raw_content": prefix,
                         "clean_content": _content_without_urls(prefix)})
    for i, match in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        body = text[match.end():end]
        sections.append({
            "sub_heading": _clean_sub_heading(match.group(1)),
            "raw_content": body,
            "clean_content": _content_without_urls(body),
        })
    return sections


def _section_diff_lookup(train_value: object, prod_value: object) -> dict:
    """Return {(parent heading, sub-heading): changed-content details}."""
    lookup = {}
    for train_item, prod_item in _pair_content_items_for_diff(train_value, prod_value):
        train_sections = _split_content_by_subheading(train_item.get("raw_content", ""))
        prod_sections = _split_content_by_subheading(prod_item.get("raw_content", ""))
        remaining = list(prod_sections)
        for pos, train_section in enumerate(train_sections):
            sub = train_section["sub_heading"]
            mi = next((i for i, x in enumerate(remaining)
                       if x["sub_heading"] == sub), None)
            if mi is None and pos < len(remaining):
                mi = pos
            prod_section = remaining.pop(mi) if mi is not None else {
                "sub_heading": "Null", "clean_content": "Null"
            }
            details = _changed_content_details(
                train_section.get("clean_content", "Null"),
                prod_section.get("clean_content", "Null"),
            )
            lookup[(train_item.get("heading", "Null"), sub)] = details
    return lookup

# ============================================================
# EXTRACT CONTENT LINKS
# ============================================================

def _extract_link_details(
    value: object,
) -> list[tuple[str, str, str, str]]:

    if pd.isna(value):
        return [
            (
                "Null",
                "Null",
                "Null",
                "Null",
                "Null",
            )
        ]

    text = str(value).strip()

    if text in {
        "",
        "Null",
        "[]",
    }:
        return [
            (
                "Null",
                "Null",
                "Null",
                "Null",
                "Null",
            )
        ]

    parsed = None

    try:
        parsed = ast.literal_eval(
            text
        )
    except (
        ValueError,
        SyntaxError,
    ):
        parsed = None

    items = (
        parsed
        if isinstance(parsed, list)
        else []
    )

    if (
        not items
        and isinstance(value, list)
    ):
        items = value

    rows = []

    for item in items:

        if not isinstance(
            item,
            dict,
        ):
            continue

        heading = (
            str(
                item.get("heading")
                or "Null"
            )
            .strip()
            or "Null"
        )

        content = (
            item.get("content")
            or ""
        )

        content_text = str(
            content
        )

        url_matches = list(
            re.finditer(
                r"\[([^\]]+?)\]\(\s*"
                r"(https?://[^)\s]+)\s*\)"
                r"|(?P<bare>https?://[^\s)\]>]+)",
                content_text,
            )
        )

        for (
            match_index,
            match,
        ) in enumerate(
            url_matches
        ):

            url = (
                match.group(2)
                or match.group("bare")
            )

            url = url.rstrip(
                ".,;:"
            )

            visible_link_text = (
                re.sub(
                    r"\*+",
                    "",
                    match.group(1),
                ).strip()
                if match.group(1)
                else ""
            )

            url_name_content = (
                _context_for_url(
                    content_text,
                    url_matches,
                    match_index,
                )
            )

            sub_heading = _subheading_for_position(content_text, match.start())

            rows.append(
                (
                    heading,
                    sub_heading,
                    visible_link_text or "Null",
                    url,
                    url_name_content,
                )
            )

        if not url_matches:

            # A content item without URLs can still contain one or more
            # Markdown sub-headings. Emit one row per sub-heading so the
            # translator output remains scoped to that section.
            sections = _split_content_by_subheading(content_text)
            for section in sections:
                rows.append(
                    (
                        heading,
                        section.get("sub_heading", "Null"),
                        "Null",
                        "Null",
                        "Null",
                    )
                )

    if not rows:

        for match in (
            MARKDOWN_LINK_PATTERN
            .finditer(text)
        ):
            rows.append(
                (
                    "Null",
                    "Null",
                    match.group(1),
                    match.group(2),
                    "Null",
                )
            )

    return (
        rows
        if rows
        else [
            (
                "Null",
                "Null",
                "Null",
                "Null",
                "Null",
            )
        ]
    )



def _extract_content_items_for_diff(value: object) -> list[dict]:
    """Parse serialized contents and return one normalized record per content item."""
    if pd.isna(value):
        return []
    if isinstance(value, list):
        items = value
    else:
        text = str(value).strip()
        if text in {"", "Null", "[]"}:
            return []
        try:
            items = ast.literal_eval(text)
        except (ValueError, SyntaxError):
            return []
    if not isinstance(items, list):
        return []
    out = []
    for idx, item in enumerate(items):
        if not isinstance(item, dict):
            continue
        heading = str(item.get("heading") or "Null").strip() or "Null"
        raw_content = str(item.get("content") or "")
        clean_content = _content_without_urls(raw_content)
        out.append({
            "index": idx,
            "heading": heading,
            "raw_content": raw_content,
            "clean_content": clean_content,
        })
    return out


def _pair_content_items_for_diff(train_value: object, prod_value: object) -> list[tuple[dict, dict]]:
    """Pair content items primarily by heading, then by source position."""
    train_items = _extract_content_items_for_diff(train_value)
    prod_items = _extract_content_items_for_diff(prod_value)
    remaining = list(prod_items)
    pairs = []
    for train_item in train_items:
        match_index = next((i for i, x in enumerate(remaining) if x["heading"] == train_item["heading"]), None)
        if match_index is None:
            match_index = next((i for i, x in enumerate(remaining) if x["index"] == train_item["index"]), None)
        prod_item = remaining.pop(match_index) if match_index is not None else {"index": -1, "heading": "Null", "raw_content": "", "clean_content": "Null"}
        pairs.append((train_item, prod_item))
    return pairs


def _remove_exact_metadata_lines(changed_text: object, *metadata_values: object) -> str:
    """Remove only whole lines exactly equal to structural metadata; preserve prose."""
    text = "" if pd.isna(changed_text) else str(changed_text)
    flattened = []
    for value in metadata_values:
        if isinstance(value, (list, tuple, set)):
            flattened.extend(value)
        else:
            flattened.append(value)
    targets = {
        re.sub(r"\s+", " ", str(v).strip()).casefold()
        for v in flattened
        if not pd.isna(v) and str(v).strip() not in {"", "Null"}
    }
    kept = []
    for line in text.splitlines():
        key = re.sub(r"\s+", " ", line.strip()).casefold()
        if key and key in targets:
            continue
        kept.append(line)
    return "\n".join(kept).strip()


def _extract_url_specific_segments(raw_content: object) -> list[dict]:
    """
    Extract URL-specific prose for patterns such as:

        - [Link A](url): Description A
        - [Link B](url): Description B

    A segment is URL-specific only when the text immediately following the
    Markdown link begins with a colon.  This deliberately does NOT classify
    shared-prose patterns where prose appears before a group of links.
    """
    text = "" if raw_content is None else str(raw_content)
    matches = list(MARKDOWN_LINK_PATTERN.finditer(text))
    rows = []
    for index, match in enumerate(matches):
        next_start = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        after = text[match.end():next_start]
        # Ignore whitespace after the link, but require ':' as the first
        # meaningful character to prove that the prose belongs to this URL.
        stripped = after.lstrip()
        if not stripped.startswith(":"):
            continue
        description = stripped[1:]
        # Remove list syntax/blank lines that belong to the next presentation
        # boundary while preserving the actual sentence text.
        description = re.sub(r"(?m)^\s*[-*+]\s*$", "", description)
        description = description.strip()
        clean_description = _content_without_urls(description)
        rows.append({
            "index": index,
            "label": re.sub(r"\*+", "", match.group(1)).strip() or "Null",
            "url": match.group(2).rstrip(".,;:"),
            "sub_heading": _subheading_for_position(text, match.start()),
            "clean_content": clean_description,
        })
    return rows


def _url_specific_diff_lookup(train_value: object, prod_value: object) -> dict:
    """
    Return URL-specific change details keyed by
    (parent heading, sub-heading, train URL label).

    Matching prefers the same normalized URL label and falls back to URL
    position within the same content item.
    """
    lookup = {}
    for train_item, prod_item in _pair_content_items_for_diff(train_value, prod_value):
        train_segments = _extract_url_specific_segments(train_item.get("raw_content", ""))
        prod_segments = _extract_url_specific_segments(prod_item.get("raw_content", ""))
        remaining = list(prod_segments)
        for train_seg in train_segments:
            label_key = re.sub(r"\s+", " ", train_seg["label"].strip()).casefold()
            mi = next((
                i for i, x in enumerate(remaining)
                if re.sub(r"\s+", " ", x["label"].strip()).casefold() == label_key
                and x.get("sub_heading", "Null") == train_seg.get("sub_heading", "Null")
            ), None)
            if mi is None:
                mi = next((i for i, x in enumerate(remaining)
                           if x.get("index") == train_seg.get("index")), None)
            prod_seg = remaining.pop(mi) if mi is not None else {"clean_content": "Null"}
            details = _changed_content_details(
                train_seg.get("clean_content", "Null"),
                prod_seg.get("clean_content", "Null"),
            )
            lookup[(
                train_item.get("heading", "Null"),
                train_seg.get("sub_heading", "Null"),
                train_seg.get("label", "Null"),
            )] = details
    return lookup


def _section_url_names_lookup(value: object) -> dict:
    """Return {(parent heading, sub-heading): [all Markdown link labels in that section]}."""
    lookup = {}
    for item in _extract_content_items_for_diff(value):
        heading = item.get("heading", "Null")
        for section in _split_content_by_subheading(item.get("raw_content", "")):
            sub_heading = section.get("sub_heading", "Null")
            raw = str(section.get("raw_content", "") or "")
            labels = []
            for match in MARKDOWN_LINK_PATTERN.finditer(raw):
                label = re.sub(r"\*+", "", match.group(1)).strip()
                if label and label not in labels:
                    labels.append(label)
            lookup[(heading, sub_heading)] = labels
    return lookup

# ============================================================
# PAIR TRAIN / PROD LINKS
# ============================================================

def _pair_link_details(
    train_links,
    prod_links,
):

    remaining_prod = list(
        prod_links
    )

    pairs = []

    for train_link in train_links:

        match_index = next(
            (
                index
                for (
                    index,
                    prod_link,
                )
                in enumerate(
                    remaining_prod
                )
                if (
                    prod_link[:5]
                    ==
                    train_link[:5]
                )
            ),
            None,
        )

        if match_index is None:

            match_index = next(
                (
                    index
                    for (
                        index,
                        prod_link,
                    )
                    in enumerate(
                        remaining_prod
                    )
                    if (
                        prod_link[0]
                        ==
                        train_link[0]
                        and prod_link[1] == train_link[1]
                    )
                ),
                None,
            )

        if match_index is not None:

            pairs.append(
                (
                    train_link,
                    remaining_prod.pop(
                        match_index
                    ),
                )
            )

        else:

            pairs.append(
                (
                    train_link,
                    (
                        "Null",
                        "Null",
                        "Null",
                        "Null",
                        "Null",
                    ),
                )
            )

    pairs.extend(
        (
            (
                "Null",
                "Null",
                "Null",
                "Null",
                "Null",
            ),
            prod_link,
        )
        for prod_link
        in remaining_prod
    )

    return pairs


# ============================================================
# CHANGED WORDS
# ============================================================

def _changed_words(
    train_value: object,
    prod_value: object,
) -> str:

    train_text = (
        ""
        if pd.isna(train_value)
        else str(train_value)
    )

    prod_text = (
        ""
        if pd.isna(prod_value)
        else str(prod_value)
    )

    train_words = train_text.split()
    prod_words = prod_text.split()

    changed = []

    for (
        tag,
        train_start,
        train_end,
        _,
        _,
    ) in (
        difflib.SequenceMatcher(
            None,
            train_words,
            prod_words,
        )
        .get_opcodes()
    ):

        if tag in {
            "delete",
            "replace",
        }:

            changed.extend(
                train_words[
                    train_start:
                    train_end
                ]
            )

    return " ".join(
        changed
    )


# ============================================================
# ROW KEY
# ============================================================

def _row_key_value() -> str:

    return (
        uuid.uuid4()
        .hex[:8]
        .upper()
    )


# ============================================================
# NULL NORMALIZATION
# ============================================================

def _normalize_null_like(
    value: object,
) -> object:

    if pd.isna(value):
        return "Null"

    if isinstance(
        value,
        str,
    ):

        stripped = value.strip()

        if stripped in {
            "",
            "[]",
            "Null",
        }:
            return "Null"

    return value


# ============================================================
# CONTENT WITHOUT URLS
# ============================================================

def _content_without_urls(
    value: object,
) -> object:

    if pd.isna(value):
        return "Null"

    text = str(value)

    if text.strip() in {
        "",
        "[]",
        "Null",
    }:
        return "Null"

    try:
        parsed = ast.literal_eval(
            text
        )
    except (
        ValueError,
        SyntaxError,
    ):
        parsed = None

    items = (
        parsed
        if isinstance(parsed, list)
        else [text]
    )

    cleaned_items = []

    for item in items:

        if not isinstance(
            item,
            dict,
        ):
            values = [
                str(item)
            ]

        else:
            values = [
                str(item_value)
                for item_value
                in item.values()
                if (
                    item_value is not None
                    and
                    str(item_value).strip()
                )
            ]

        for item_value in values:

            item_value = re.sub(
                r"\[([^\[\]]+?)\]\("
                r"\s*https?://[^)\s]+"
                r"(?:\s+[^)]*)?\s*\)",
                r"\1\n",
                item_value,
            )

            item_value = (
                BARE_URL_PATTERN.sub(
                    "\n",
                    item_value,
                )
            )

            item_value = re.sub(
                r"\*+",
                "",
                item_value,
            )

            item_value = re.sub(
                r"(?m)^\s*[-*+]\s+",
                "",
                item_value,
            )

            lines = [
                re.sub(
                    r"\s+",
                    " ",
                    line,
                ).strip()
                for line
                in item_value.splitlines()
                if line.strip()
            ]

            cleaned_items.extend(
                line
                for line
                in lines
                if line
            )

    return (
        "\n".join(
            cleaned_items
        )
        or "Null"
    )


# ============================================================
# CONTENT CHANGE DETAILS
# ============================================================

def _changed_content_details(
    train_value: object,
    prod_value: object,
) -> dict:
    """
    Compare Train vs Prod content line-by-line.

    IMPORTANT:

    Content Text Changes contains ONLY the actual English
    changed content.

    It does NOT contain wording such as:

        ----- CHANGE BLOCK 1 -----
        ----- CHANGE BLOCK 2 -----
        ----- CHANGE BLOCK 3 -----

    Technical information is stored separately in:

        Content Text Context Before
        Content Text Context After
        Content Text Change Block Count

    Those technical fields are kept in the Comparison sheet
    for Stage 2 and are NOT shown on the French Content
    Translation sheet.
    """

    train_text = (
        ""
        if pd.isna(train_value)
        else str(train_value)
    )

    prod_text = (
        ""
        if pd.isna(prod_value)
        else str(prod_value)
    )

    train_lines = (
        train_text.splitlines()
    )

    prod_lines = (
        prod_text.splitlines()
    )

    matcher = difflib.SequenceMatcher(
        None,
        train_lines,
        prod_lines,
        autojunk=False,
    )

    blocks = []

    for (
        tag,
        train_start,
        train_end,
        prod_start,
        prod_end,
    ) in matcher.get_opcodes():

        if tag not in {
            "delete",
            "replace",
        }:
            continue

        if (
            train_start
            >=
            train_end
        ):
            continue

        changed_lines = (
            train_lines[
                train_start:
                train_end
            ]
        )

        changed_text = (
            "\n".join(
                changed_lines
            )
            .strip()
        )

        if not changed_text:
            continue

        # ----------------------------------------------------
        # ONE LINE BEFORE
        # ----------------------------------------------------

        before = ""

        if train_start > 0:

            before = (
                train_lines[
                    train_start - 1
                ]
                .strip()
            )

        # ----------------------------------------------------
        # ONE LINE AFTER
        # ----------------------------------------------------

        after = ""

        if (
            train_end
            <
            len(train_lines)
        ):

            after = (
                train_lines[
                    train_end
                ]
                .strip()
            )

        blocks.append(
            {
                "before":
                    before,

                "changed":
                    changed_text,

                "after":
                    after,
            }
        )

    # ========================================================
    # NO CHANGES
    # ========================================================

    if not blocks:

        return {
            "before": "",
            "changed": "",
            "after": "",
            "block_count": 0,
        }

    # ========================================================
    # ONE CHANGE BLOCK
    # ========================================================

    if len(blocks) == 1:

        return {
            "before":
                blocks[0][
                    "before"
                ],

            "changed":
                blocks[0][
                    "changed"
                ],

            "after":
                blocks[0][
                    "after"
                ],

            "block_count":
                1,
        }

    # ========================================================
    # MULTIPLE CHANGE BLOCKS
    #
    # NO "CHANGE BLOCK" LABELS ARE INSERTED.
    #
    # Blocks are separated only by a blank line.
    # ========================================================

    before_parts = []

    changed_parts = []

    after_parts = []

    for block in blocks:

        if block[
            "before"
        ]:

            before_parts.append(
                block[
                    "before"
                ]
            )

        if block[
            "changed"
        ]:

            changed_parts.append(
                block[
                    "changed"
                ]
            )

        if block[
            "after"
        ]:

            after_parts.append(
                block[
                    "after"
                ]
            )

    return {
        "before":
            "\n\n".join(
                before_parts
            ),

        "changed":
            "\n\n".join(
                changed_parts
            ),

        "after":
            "\n\n".join(
                after_parts
            ),

        "block_count":
            len(blocks),
    }


# ============================================================
# BACKWARD COMPATIBILITY
# ============================================================

def _changed_content_text(
    train_value: object,
    prod_value: object,
) -> str:

    return (
        _changed_content_details(
            train_value,
            prod_value,
        )[
            "changed"
        ]
    )


# ============================================================
# NORMALIZE ONE LOCALE
# ============================================================

def process_file(
    input_path: str,
    output_csv: str,
    locale: str,
) -> pd.DataFrame:

    input_path = Path(
        input_path
    )

    output_csv = Path(
        output_csv
    )

    if not input_path.exists():

        raise FileNotFoundError(
            "\nSource JSON not found:\n"
            f"{input_path}"
        )

    with open(
        input_path,
        "r",
        encoding="utf-8",
    ) as fp:

        jsondata = json.load(
            fp
        )

    filtered = [
        obj
        for obj in jsondata
        if (
            obj.get(
                "publishedAt"
            )
            is not None
            and
            obj.get(
                "locale"
            )
            ==
            locale
        )
    ]

    df_raw = pd.DataFrame(
        filtered
    )

    if (
        not df_raw.empty
        and
        "details"
        in df_raw.columns
    ):

        details_df = (
            pd.json_normalize(
                df_raw[
                    "details"
                ]
            )
            .reset_index(
                drop=True
            )
        )

        other_df = (
            df_raw
            .drop(
                columns=[
                    "details"
                ]
            )
            .reset_index(
                drop=True
            )
        )

        df = pd.concat(
            [
                other_df,
                details_df,
            ],
            axis=1,
        )

    else:

        df = (
            df_raw.copy()
        )

    if (
        "data_export_date"
        in df.columns
    ):

        df = (
            df.drop(
                columns=[
                    "data_export_date"
                ]
            )
        )

    df.insert(
        0,
        "data_export_date",
        date.today().isoformat(),
    )

    output_csv.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    df.to_csv(
        output_csv,
        encoding="utf-8-sig",
        index=False,
    )

    print()

    print(
        f"Locale : {locale}"
    )

    print(
        f"Records: {len(df)}"
    )

    print(
        f"Saved  : {output_csv}"
    )

    return df


# ============================================================
# JOIN TRAIN / PROD
# ============================================================

def _is_empty_train_content(values: list[object]) -> bool:
    return all(
        pd.isna(value)
        or str(value).strip().lower() in {"", "null", "none", "nan"}
        for value in values
    )


def _populate_train_output_content(
    train_output_df: pd.DataFrame,
    train_df: pd.DataFrame,
) -> None:
    if "contents" not in train_df.columns:
        raise ValueError("Train Input is missing the required 'contents' column.")

    contents_by_uuid = train_df.set_index("__uuid")["contents"]
    for train_uuid, output_rows in train_output_df.groupby(
        "__uuid",
        sort=False,
    ):
        parsed_rows = [
            values
            for values in _extract_link_details(contents_by_uuid.loc[train_uuid])
            if not _is_empty_train_content(list(values))
        ]
        output_indexes = [
            index
            for index, row in output_rows.iterrows()
            if not _is_empty_train_content(
                row[TRAIN_PARSED_COLUMNS].tolist()
            )
        ]
        if len(parsed_rows) != len(output_indexes):
            raise ValueError(
                "Unable to align Train Input contents with Comparison rows "
                f"for UUID {train_uuid}: {len(parsed_rows)} parsed content "
                f"rows, {len(output_indexes)} Comparison content rows."
            )

        for index, values in zip(output_indexes, parsed_rows):
            parsed_values = [
                values[0],
                values[1],
                values[2],
                values[4],
                normalize_train_url(values[3]) or values[3],
            ]
            for column, value in zip(
                TRAIN_PARSED_COLUMNS,
                parsed_values,
            ):
                train_output_df.at[index, column] = value


def write_train_output_sheet(
    writer,
    train_df: pd.DataFrame,
    comparison_df: pd.DataFrame,
    french_urls_df: pd.DataFrame,
    french_content_df: pd.DataFrame,
) -> pd.DataFrame:
    train_output_df = build_train_output(
        train_df,
        comparison_df,
        french_urls_df,
        french_content_df,
    )
    _populate_train_output_content(train_output_df, train_df)
    train_output_df.to_excel(
        writer,
        sheet_name="Train output",
        index=False,
    )
    format_train_output_sheet(writer.book["Train output"])
    return train_output_df


def join_export_files(
    train_csv: str,
    prod_csv: str,
    output_csv: str,
    join_columns: Union[
        List[str],
        Tuple[str, ...],
    ] = (
        "Local Community Name",
        "title",
    ),
) -> pd.DataFrame:

    # ========================================================
    # READ NORMALIZED FILES
    # ========================================================

    train_df = pd.read_csv(
        train_csv,
        dtype=object,
    )

    prod_df = pd.read_csv(
        prod_csv,
        dtype=object,
    )

    # ========================================================
    # VALIDATE JOIN COLUMNS
    # ========================================================

    missing_cols = [
        column
        for column
        in join_columns
        if (
            column
            not in train_df.columns
            or
            column
            not in prod_df.columns
        )
    ]

    if missing_cols:

        raise ValueError(
            "Missing join columns: "
            f"{missing_cols}"
        )

    # ========================================================
    # PREFIX OVERLAPPING COLUMNS
    # ========================================================

    overlapping_columns = [
        column
        for column
        in train_df.columns.intersection(
            prod_df.columns
        )
        if (
            column
            not in join_columns
        )
    ]

    train_renamed = (
        train_df.rename(
            columns={
                column:
                    f"train_{column}"
                for column
                in overlapping_columns
            }
        )
    )

    prod_renamed = (
        prod_df.rename(
            columns={
                column:
                    f"prod_{column}"
                for column
                in overlapping_columns
            }
        )
    )

    train_renamed = (
        train_renamed.rename(
            columns={
                "train_data_export_date":
                    "data_export_date"
            }
        )
    )

    if (
        "prod_data_export_date"
        in prod_renamed.columns
    ):

        prod_renamed = (
            prod_renamed.drop(
                columns=[
                    "prod_data_export_date"
                ]
            )
        )

    # ========================================================
    # DROP UNNECESSARY COMPARISON COLUMNS
    # ========================================================

    columns_to_drop = {
        "prod___uuid",

        "train_locale",
        "prod_locale",

        "train_createdAt",
        "prod_createdAt",

        "train_updatedAt",
        "prod_updatedAt",

        "train_publishedAt",
        "prod_publishedAt",

        "train_buttonText",
        "prod_buttonText",

        "train_description",
        "prod_description",

        "train_buttonAriaLabel",
        "prod_buttonAriaLabel",

        "train_buttonVariant",
        "prod_buttonVariant",

        "train_image.name",
        "prod_image.name",

        "train_localizations",
        "prod_localizations",
    }

    train_renamed = (
        train_renamed.drop(
            columns=[
                column
                for column
                in columns_to_drop
                if (
                    column
                    in train_renamed.columns
                )
            ],
            errors="ignore",
        )
    )

    prod_renamed = (
        prod_renamed.drop(
            columns=[
                column
                for column
                in columns_to_drop
                if (
                    column
                    in prod_renamed.columns
                )
            ],
            errors="ignore",
        )
    )

    # ========================================================
    # FULL OUTER JOIN
    # ========================================================

    merged_df = (
        train_renamed.merge(
            prod_renamed,
            left_on=list(
                join_columns
            ),
            right_on=list(
                join_columns
            ),
            how="outer",
            indicator=True,
        )
    )

    # ========================================================
    # NORMALIZE CONTENTS
    # ========================================================

    for column in [
        column
        for column
        in merged_df.columns
        if (
            column.endswith(
                "contents"
            )
        )
    ]:

        merged_df[
            column
        ] = (
            merged_df[
                column
            ]
            .apply(
                _normalize_null_like
            )
        )

        merged_df[
            column
        ] = (
            merged_df[
                column
            ]
            .apply(
                lambda value:
                    value.replace(
                        "train.",
                        "",
                    )
                    if isinstance(
                        value,
                        str,
                    )
                    else value
            )
        )

        merged_df[
            column
        ] = (
            merged_df[
                column
            ]
            .apply(
                _normalize_null_like
            )
        )

    # ========================================================
    # TILE / SUBPAGE
    # ========================================================

    for (
        content_column,
        result_column,
    ) in (
        (
            "train_contents",
            "train_subpage_or_tile",
        ),
        (
            "prod_contents",
            "prod_subpage_or_tile",
        ),
    ):

        if (
            content_column
            not in merged_df.columns
        ):

            continue

        merged_df[
            result_column
        ] = (
            merged_df[
                content_column
            ]
            .apply(
                lambda value:
                    "Tile"
                    if (
                        isinstance(
                            value,
                            str,
                        )
                        and
                        value.strip()
                        ==
                        "Null"
                    )
                    else (
                        ""
                        if (
                            pd.isna(
                                value
                            )
                            or
                            (
                                isinstance(
                                    value,
                                    str,
                                )
                                and
                                value.strip()
                                ==
                                ""
                            )
                        )
                        else
                        "Subpage"
                    )
            )
        )

    if (
        "train_subpage_or_tile"
        in merged_df.columns
        and
        "prod_subpage_or_tile"
        in merged_df.columns
    ):

        merged_df[
            "Tile v Subpage Match"
        ] = (
            (
                merged_df[
                    "train_subpage_or_tile"
                ]
                .fillna(
                    "Null"
                )
                .astype(str)
            )
            ==
            (
                merged_df[
                    "prod_subpage_or_tile"
                ]
                .fillna(
                    "Null"
                )
                .astype(str)
            )
        ).map(
            {
                True:
                    "Yes",
                False:
                    "No",
            }
        )

    # ========================================================
    # CONTENT COMPARISON
    # ========================================================

    if (
        "train_contents"
        in merged_df.columns
        and
        "prod_contents"
        in merged_df.columns
    ):

        merged_df[
            "train_contents_without_urls"
        ] = (
            merged_df[
                "train_contents"
            ]
            .apply(
                _content_without_urls
            )
        )

        merged_df[
            "prod_contents_without_urls"
        ] = (
            merged_df[
                "prod_contents"
            ]
            .apply(
                _content_without_urls
            )
        )

        merged_df[
            "Content Without URLs Match"
        ] = (
            (
                merged_df[
                    "train_contents_without_urls"
                ]
                .fillna(
                    "Null"
                )
                .astype(str)
            )
            ==
            (
                merged_df[
                    "prod_contents_without_urls"
                ]
                .fillna(
                    "Null"
                )
                .astype(str)
            )
        ).map(
            {
                True:
                    "Yes",
                False:
                    "No",
            }
        )

        # ====================================================
        # CONTENT CHANGE DETAILS
        # ====================================================

        content_change_details = (
            merged_df.apply(
                lambda row:
                    _changed_content_details(
                        row[
                            "train_contents_without_urls"
                        ],
                        row[
                            "prod_contents_without_urls"
                        ],
                    ),
                axis=1,
            )
        )

        # ----------------------------------------------------
        # TECHNICAL - COMPARISON ONLY
        # ----------------------------------------------------

        merged_df[
            "Content Text Context Before"
        ] = (
            content_change_details.apply(
                lambda details:
                    details[
                        "before"
                    ]
            )
        )

        # ----------------------------------------------------
        # ACTUAL TRANSLATION TEXT
        #
        # NO CHANGE BLOCK LABELS.
        # ----------------------------------------------------

        merged_df[
            "Content Text Changes"
        ] = (
            content_change_details.apply(
                lambda details:
                    details[
                        "changed"
                    ]
            )
        )

        # ----------------------------------------------------
        # TECHNICAL - COMPARISON ONLY
        # ----------------------------------------------------

        merged_df[
            "Content Text Context After"
        ] = (
            content_change_details.apply(
                lambda details:
                    details[
                        "after"
                    ]
            )
        )

        # ----------------------------------------------------
        # TECHNICAL - COMPARISON ONLY
        # ----------------------------------------------------

        merged_df[
            "Content Text Change Block Count"
        ] = (
            content_change_details.apply(
                lambda details:
                    details[
                        "block_count"
                    ]
            )
        )

        merged_df[
            "Content Match"
        ] = (
            merged_df[
                "Content Without URLs Match"
            ]
        )

    # ========================================================
    # BUTTON LINK COMPARISON
    # ========================================================

    if (
        "train_buttonLink"
        in merged_df.columns
        and
        "prod_buttonLink"
        in merged_df.columns
    ):

        merged_df[
            "Button Link Match"
        ] = (
            (
                merged_df[
                    "train_buttonLink"
                ]
                .fillna(
                    "Null"
                )
                .astype(str)
            )
            ==
            (
                merged_df[
                    "prod_buttonLink"
                ]
                .fillna(
                    "Null"
                )
                .astype(str)
            )
        ).map(
            {
                True:
                    "Yes",
                False:
                    "No",
            }
        )

    # ========================================================
    # EXPAND CONTENT URLS + CALCULATE DIFF PER HEADING/CONTENT ITEM
    # ========================================================

    expanded_rows = []

    for _, row in merged_df.iterrows():
        train_links = _extract_link_details(row.get("train_contents"))
        prod_links = _extract_link_details(row.get("prod_contents"))
        link_pairs = _pair_link_details(train_links, prod_links)

        # Calculate changes at parent-heading + Markdown-sub-heading level.
        # This prevents a change under one #### section from being repeated
        # for every URL under the same parent train_heading.
        section_diffs = _section_diff_lookup(
            row.get("train_contents"),
            row.get("prod_contents"),
        )
        # All URL labels are structural metadata.  Remove every link label
        # belonging to the current heading/sub-heading section from the prose
        # sent for translation, not only the label for the current expanded row.
        section_url_names = _section_url_names_lookup(row.get("train_contents"))
        # For patterns like [URL](...): description, calculate the diff at
        # the individual URL level so descriptions from neighboring URLs are
        # not repeated in every expanded row.
        url_specific_diffs = _url_specific_diff_lookup(
            row.get("train_contents"),
            row.get("prod_contents"),
        )

        for train_link, prod_link in link_pairs:
            expanded_row = row.to_dict()
            train_heading = train_link[0]
            train_heading_sub_heading = train_link[1]
            train_url_name = train_link[2]

            # Prefer URL-specific prose when the source proves a
            # [URL](...): description relationship. Otherwise retain the
            # shared heading/sub-heading behavior used for prose that belongs
            # to a group of links.
            details = url_specific_diffs.get(
                (train_heading, train_heading_sub_heading, train_url_name)
            )
            if details is None:
                details = section_diffs.get(
                    (train_heading, train_heading_sub_heading),
                    {"changed": "", "before": "", "after": "", "block_count": 0},
                )

            all_section_url_names = section_url_names.get(
                (train_heading, train_heading_sub_heading),
                [],
            )
            per_heading_change = _remove_exact_metadata_lines(
                details.get("changed", ""),
                train_heading,
                train_heading_sub_heading,
                all_section_url_names,
            )

            expanded_row.update({
                "train_heading": train_heading,
                "train_heading_sub_heading": train_heading_sub_heading,
                "train_url_name": train_url_name,
                "train_url": train_link[3],
                "train_url_name_content": train_link[4],
                "prod_heading": prod_link[0],
                "prod_heading_sub_heading": prod_link[1],
                "prod_url_name": prod_link[2],
                "prod_url": prod_link[3],
                "prod_url_name_content": prod_link[4],
                "Content Text Changes": per_heading_change,
                "Content Text Context Before": details.get("before", ""),
                "Content Text Context After": details.get("after", ""),
                "Content Text Change Block Count": details.get("block_count", 0),
            })
            expanded_rows.append(expanded_row)

    merged_df = pd.DataFrame(expanded_rows)

    # ========================================================
    # IDS
    # ========================================================

    merged_df.insert(
        0,
        "Comparison ID",
        range(
            1,
            len(
                merged_df
            )
            + 1,
        ),
    )

    merged_df.insert(
        1,
        "Row Key",
        [
            _row_key_value()
            for _
            in range(
                len(
                    merged_df
                )
            )
        ],
    )

    # ========================================================
    # MATCH STATUS
    # ========================================================

    def _match_status(
        left_value,
        right_value,
    ):

        left_is_null = (
            pd.isna(
                left_value
            )
            or
            (
                isinstance(
                    left_value,
                    str,
                )
                and
                left_value.strip()
                in {
                    "",
                    "Null",
                }
            )
        )

        right_is_null = (
            pd.isna(
                right_value
            )
            or
            (
                isinstance(
                    right_value,
                    str,
                )
                and
                right_value.strip()
                in {
                    "",
                    "Null",
                }
            )
        )

        if (
            left_is_null
            and
            not right_is_null
        ):

            return (
                "Only in prod"
            )

        if (
            left_value
            ==
            right_value
        ):

            return "Yes"

        return "No"

    for (
        left_column,
        right_column,
        result_name,
    ) in (
        (
            "train_heading",
            "prod_heading",
            "Heading Match",
        ),
        (
            "train_url_name",
            "prod_url_name",
            "URL Name Match",
        ),
        (
            "train_url",
            "prod_url",
            "URL Match",
        ),
    ):

        if (
            left_column
            in merged_df.columns
            and
            right_column
            in merged_df.columns
        ):

            merged_df[
                result_name
            ] = (
                merged_df.apply(
                    lambda row:
                        _match_status(
                            row.get(
                                left_column
                            ),
                            row.get(
                                right_column
                            ),
                        ),
                    axis=1,
                )
            )

    for column in [
        "Tile v Subpage Match",
        "Content Match",
        "Button Link Match",
        "Heading Match",
        "URL Name Match",
        "URL Match",
    ]:

        if (
            column
            not in merged_df.columns
        ):

            merged_df[
                column
            ] = "Yes"

    # ========================================================
    # FRENCH URL SHEET
    # ========================================================

    button_url_df = (
        merged_df[
            merged_df[
                "Button Link Match"
            ]
            .eq(
                "No"
            )
            |
            merged_df[
                "URL Match"
            ]
            .eq(
                "No"
            )
        ]
        .copy()
    )

    button_is_not_match = (
        button_url_df[
            "Button Link Match"
        ]
        .eq(
            "No"
        )
    )

    url_is_not_match = (
        button_url_df[
            "URL Match"
        ]
        .eq(
            "No"
        )
    )

    button_url_df[
        "Button and URL Difference"
    ] = (
        "Only Content URL"
    )

    button_url_df.loc[
        (
            button_is_not_match
            &
            ~url_is_not_match
        ),
        "Button and URL Difference",
    ] = (
        "Only Button Link"
    )

    button_url_df.loc[
        (
            button_is_not_match
            &
            url_is_not_match
        ),
        "Button and URL Difference",
    ] = (
        "Both Button and URL in content"
    )

    button_url_df = (
        button_url_df.drop(
            columns=[
                "_merge"
            ],
            errors="ignore",
        )
    )

    french_urls_columns = [
        "Row Key",
        "Local Community Name",
        "title",

        "train_buttonLink",
        "prod_buttonLink",

        "Content Without URLs Match",
        "Content Match",
        "Button Link Match",

        "train_heading",
        "train_heading_sub_heading",
        "train_url_name",
        "train_url_name_content",
        "train_url",

        "prod_heading",
        "prod_heading_sub_heading",
        "prod_url_name",
        "prod_url_name_content",
        "prod_url",

        "URL Match",

        "Button and URL Difference",
    ]

    button_url_df = (
        button_url_df[
            [
                column
                for column
                in french_urls_columns
                if (
                    column
                    in button_url_df.columns
                )
            ]
        ]
        .copy()
    )

    button_url_df[
        "French Button Link corresponding to column D"
    ] = ""

    button_url_df[
        "French Content URL corresponding to column J"
    ] = ""

    # ========================================================
    # FRENCH CONTENT TRANSLATION
    # ========================================================

    def _is_blank_or_null(value: object) -> bool:
        """Return True for blank/Null-like values used in comparison output."""
        if value is None:
            return True
        try:
            if pd.isna(value):
                return True
        except (TypeError, ValueError):
            pass
        text = str(value).strip()
        return text == "" or text.lower() in {"null", "none", "nan"}

    content_heading_url_df = (
        merged_df[
            merged_df[
                "Content Match"
            ]
            .eq(
                "No"
            )
            &
            (
                merged_df[
                    "Heading Match"
                ]
                .eq(
                    "No"
                )
                |
                merged_df[
                    "URL Name Match"
                ]
                .eq(
                    "No"
                )
            )
        ]
        .copy()
    )

    content_heading_url_df = (
        content_heading_url_df.drop(
            columns=[
                "_merge"
            ],
            errors="ignore",
        )
    )

    french_content_columns = [
        "Row Key",
        "Local Community Name",
        "title",

        "train_contents",
        "prod_contents",

        "train_contents_without_urls",
        "prod_contents_without_urls",

        "Content Without URLs Match",

        # ----------------------------------------------------
        # CLEAN CONTENT ONLY
        # ----------------------------------------------------

        "Content Text Changes",

        "train_heading",
        "train_heading_sub_heading",
        "train_url_name",

        "Heading Match",
        "URL Name Match",
    ]

    content_heading_url_df = (
        content_heading_url_df[
            [
                column
                for column
                in french_content_columns
                if (
                    column
                    in content_heading_url_df.columns
                )
            ]
        ]
        .copy()
    )

    # ========================================================
    # TRANSLATOR INPUT COLUMNS - O:R
    # ========================================================
    #
    # Business rules:
    #   O = French translation for Content Text Changes (column I).
    #       No translation is requested when column I is blank/Null.
    #
    #   P = French translation for train_heading (column J).
    #       If Heading Match (column M) is Yes, write No Translation Required so the
    #       translator does not modify an already matching heading.
    #
    #   Q = French translation for train_heading_sub_heading (column K).
    #       If column K is blank/Null, no translation is requested.
    #
    #   R = French translation for train_url_name (column L).
    #       If URL Name Match (column N) is Yes, write No Translation Required so the
    #       translator does not modify an already matching URL name.
    #
    # Only these four columns are translator input columns.
    # ========================================================

    # Column O: Content Text Changes French - Column I
    content_heading_url_df[
        "Content Text Changes French - Column I"
    ] = content_heading_url_df[
        "Content Text Changes"
    ].apply(
        lambda value: "No Translation Required" if _is_blank_or_null(value) else ""
    )

    # Column P: Train_heading French Column J
    content_heading_url_df[
        "Train_heading French Column J"
    ] = content_heading_url_df.apply(
        lambda row: (
            "No Translation Required"
            if str(row.get("Heading Match", "")).strip().lower() == "yes"
            else ""
        ),
        axis=1,
    )

    # Column Q: train_heading_sub_heading French Column K
    content_heading_url_df[
        "train_heading_sub_heading French Column K"
    ] = content_heading_url_df[
        "train_heading_sub_heading"
    ].apply(
        lambda value: "No Translation Required" if _is_blank_or_null(value) else ""
    )

    # Column R: train_url_name French Column L
    content_heading_url_df[
        "train_url_name French Column L"
    ] = content_heading_url_df.apply(
        lambda row: (
            "No Translation Required"
            if str(row.get("URL Name Match", "")).strip().lower() == "yes"
            else ""
        ),
        axis=1,
    )

    # ========================================================
    # SAVE CSV
    # ========================================================

    output_csv = Path(
        output_csv
    )

    output_csv.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    merged_df.to_csv(
        output_csv,
        index=False,
        encoding="utf-8-sig",
    )

    # ========================================================
    # EXCEL
    # ========================================================

    excel_output = (
        output_csv.with_suffix(
            ".xlsx"
        )
    )

    with pd.ExcelWriter(
        excel_output,
        engine="openpyxl",
    ) as writer:

        # ====================================================
        # INSTRUCTIONS
        # ====================================================

        instructions = [
            "Instructions to Ontario Health",

            'Go to the "French URLs" tab',

            "To provide French Button Links equivalents to English Button Links",

            '    Filter "Button and URL Difference" to select only '
            '"Only Button Link" and "Both Button Link and Content URL"',

            '    Filter "train_buttonLink" to exclude/deselect "#" and "Blanks"',

            '    Review "train_buttonLink" and enter the corresponding '
            'French Button Link in '
            '"French Button Link corresponding to column D"',

            "    Note: English Button Links are duplicated to accommodate "
            "for all the content in this template; copy and paste "
            "corresponding French Button Link across duplicated rows",

            "To provide French Content URLs equivalent to English Content URLs",

            '    Filter "Button and URL Difference" to select only '
            '"Only Content URL" and "Both Button Link and Content URL"',

            '    Review "train_url" and enter the corresponding French '
            'Content URL in '
            '"French Content URL corresponding to column J"',

            "",

            "Instructions to Translator",

            'Go to the "French Content Translation" tab',

            'For Content Text Changes: review "Content Text Changes" (Column I). '
            'If Column I is blank or Null, the script pre-populates '
            '"Content Text Changes French - Column I" with "No Translation Required". Otherwise enter the French translation there.',

            'For train_heading: if "Heading Match" (Column M) = Yes, the script '
            'pre-populates "Train_heading French Column J" with "No Translation Required" and it must not be modified. '
            'If "Heading Match" = No, enter the French translation in that column.',

            'For train_heading_sub_heading: if Column K is blank or Null, no French translation '
            'is required in "train_heading_sub_heading French Column K". Otherwise enter the French translation there.',

            'For train_url_name: if "URL Name Match" (Column N) = Yes, the script '
            'pre-populates "train_url_name French Column L" with "No Translation Required" and it must not be modified. '
            'If "URL Name Match" = No, enter the French translation in that column.',

            'Do not translate or replace values marked "No Translation Required"; they indicate that no French translation or update is required.',

            "",

            "Notes",

            "Do not edit source/reference columns.",

            "Only enter values in the designated French translation columns.",

            "Do not add or remove rows and columns.",

            'The "Comparison" sheet contains technical information used '
            "by Stage 2 and should not be edited.",

            "",

            "Verification input",

            'Use the "Train Input" and "Prod Input" tabs to verify '
            "comparison results against the normalized source exports.",

            "These tabs are read-only reference data and should not be edited.",
        ]

        instructions_sheet = (
            writer.book.create_sheet(
                "Instructions",
                0,
            )
        )

        section_headers = {
            "Instructions to Ontario Health",
            "Instructions to Translator",
            "Notes",
            "Verification input",
        }

        for (
            row_number,
            instruction,
        ) in enumerate(
            instructions,
            start=1,
        ):

            cell = (
                instructions_sheet.cell(
                    row=row_number,
                    column=1,
                    value=instruction,
                )
            )

            cell.alignment = (
                Alignment(
                    vertical="top",
                    wrap_text=False,
                )
            )

            if (
                instruction
                in section_headers
            ):

                cell.font = (
                    Font(
                        bold=True
                    )
                )

        # ====================================================
        # DATA SHEETS
        # ====================================================

        train_df.to_excel(
            writer,
            sheet_name="Train Input",
            index=False,
        )

        prod_df.to_excel(
            writer,
            sheet_name="Prod Input",
            index=False,
        )

        merged_df.to_excel(
            writer,
            sheet_name="Comparison",
            index=False,
        )

        button_url_df.to_excel(
            writer,
            sheet_name="French URLs",
            index=False,
        )

        # Keep Column K consistent with the other display columns:
        # null-like sub-headings are shown as blank. Column Q still uses
        # "No Translation Required" for these rows.
        if "train_heading_sub_heading" in content_heading_url_df.columns:
            content_heading_url_df["train_heading_sub_heading"] = (
                content_heading_url_df["train_heading_sub_heading"]
                .apply(
                    lambda value: ""
                    if pd.isna(value)
                    or str(value).strip().lower() in {"null", "none", "nan"}
                    else value
                )
            )

        content_heading_url_df.to_excel(
            writer,
            sheet_name="French Content Translation",
            index=False,
        )

        write_train_output_sheet(
            writer,
            train_df,
            merged_df,
            button_url_df,
            content_heading_url_df,
        )

        write_translator_note(
            writer,
        )

        # ====================================================
        # TRANSLATION WORKLOAD SUMMARY
        # ====================================================

        write_translation_summary(
            writer,
            content_heading_url_df,
        )

        # ====================================================
        # FORMATTING
        # ====================================================

        for worksheet in (
            writer.book.worksheets
        ):

            # These sheets have custom layouts and are not data tables.
            if worksheet.title in {
                "Summary",
                "Note to Translator",
            }:
                continue

            if (
                worksheet.title
                !=
                "Instructions"
            ):

                worksheet.freeze_panes = (
                    "A2"
                )

            # =================================================
            # HEADERS
            # =================================================

            for header_cell in (
                worksheet[
                    1
                ]
            ):

                header_cell.font = (
                    Font(
                        bold=True
                    )
                )

                header_cell.alignment = (
                    Alignment(
                        horizontal=(
                            header_cell
                            .alignment
                            .horizontal
                        ),
                        vertical=(
                            header_cell
                            .alignment
                            .vertical
                        ),
                        wrap_text=True,
                    )
                )

                # =============================================
                # FRENCH URL COLORS
                # =============================================

                if (
                    worksheet.title
                    ==
                    "French URLs"
                ):

                    if (
                        header_cell.value
                        in {
                            "train_buttonLink",
                            "French Button Link corresponding to column D",
                        }
                    ):

                        header_cell.fill = (
                            PatternFill(
                                fill_type="solid",
                                fgColor="FFFFA500",
                            )
                        )

                    elif (
                        header_cell.value
                        in {
                            "train_url",
                            "French Content URL corresponding to column J",
                        }
                    ):

                        header_cell.fill = (
                            PatternFill(
                                fill_type="solid",
                                fgColor="FF00B050",
                            )
                        )

                # =============================================
                # FRENCH CONTENT COLORS
                # =============================================

                elif (
                    worksheet.title
                    ==
                    "French Content Translation"
                ):

                    if (
                        header_cell.value
                        in {
                            "train_heading",
                            "Train_heading French Column J",
                        }
                    ):

                        header_cell.fill = (
                            PatternFill(
                                fill_type="solid",
                                fgColor="FFFFA500",
                            )
                        )

                    elif (
                        header_cell.value
                        in {
                            "train_heading_sub_heading",
                            "train_heading_sub_heading French Column K",
                        }
                    ):

                        header_cell.fill = (
                            PatternFill(
                                fill_type="solid",
                                fgColor="FF70AD47",
                            )
                        )

                    elif (
                        header_cell.value
                        in {
                            "train_url_name",
                            "train_url_name French Column L",
                        }
                    ):

                        header_cell.fill = (
                            PatternFill(
                                fill_type="solid",
                                fgColor="FF00B050",
                            )
                        )

                    elif (
                        header_cell.value
                        in {
                            "Heading Match",
                            "URL Name Match",
                        }
                    ):

                        header_cell.fill = (
                            PatternFill(
                                fill_type="solid",
                                fgColor="FF4F81BD",
                            )
                        )

                    elif (
                        header_cell.value
                        in {
                            "Content Text Changes",
                            "Content Text Changes French - Column I",
                        }
                    ):

                        header_cell.fill = (
                            PatternFill(
                                fill_type="solid",
                                fgColor="FFFF0000",
                            )
                        )

                elif (
                    worksheet.title
                    ==
                    "Train output"
                ):

                    train_output_header_colors = {
                        "train_buttonLink": "FF87CEEB",
                        "French Button Link corresponding to column D": "FFFFA500",
                        "train_url": "FF87CEEB",
                        "French Content URL corresponding to column J": "FF00B050",
                        "Content Text Changes": "FF87CEEB",
                        "Content Text Changes French - Column I": "FFFF0000",
                        "train_heading": "FF87CEEB",
                        "Train_heading French Column J": "FFFFFF00",
                        "train_heading_sub_heading": "FF87CEEB",
                        "train_heading_sub_heading French Column K": "FF7030A0",
                        "train_url_name": "FF87CEEB",
                        "train_url_name French Column L": "FF92D050",
                        "Heading Match": "FF4F81BD",
                        "URL Name Match": "FF4F81BD",
                        "Row Key": "FF4F81BD",
                    }
                    fill_color = train_output_header_colors.get(
                        header_cell.value
                    )
                    if fill_color:
                        header_cell.fill = PatternFill(
                            fill_type="solid",
                            fgColor=fill_color,
                        )

            # =================================================
            # COLUMN WIDTH
            # =================================================

            for column_cells in (
                worksheet.iter_cols()
            ):

                max_content_width = max(
                    (
                        len(
                            str(
                                cell.value
                            )
                        )
                        if (
                            cell.value
                            is not None
                        )
                        else 0
                    )
                    for cell
                    in column_cells
                )

                column_letter = (
                    column_cells[
                        0
                    ].column_letter
                )

                worksheet.column_dimensions[
                    column_letter
                ].width = min(
                    max_content_width
                    + 2,
                    20,
                )

            if (
                worksheet.title
                ==
                "Instructions"
            ):

                worksheet.column_dimensions[
                    "A"
                ].width = 150

            # =================================================
            # TABLES
            # =================================================

            if (
                worksheet.title
                !=
                "Instructions"
                and
                worksheet.max_column
                >
                0
                and
                worksheet.max_row
                >
                0
            ):

                last_cell = (
                    worksheet.cell(
                        row=(
                            worksheet.max_row
                        ),
                        column=(
                            worksheet.max_column
                        ),
                    )
                    .coordinate
                )

                table_ref = (
                    f"A1:{last_cell}"
                )

                safe_table_name = (
                    re.sub(
                        r"[^A-Za-z0-9_]",
                        "",
                        worksheet.title,
                    )
                )

                table = (
                    Table(
                        displayName=(
                            f"{safe_table_name}Table"
                        ),
                        ref=(
                            table_ref
                        ),
                    )
                )

                table.tableStyleInfo = (
                    TableStyleInfo(
                        name=(
                            "TableStyleMedium2"
                        ),
                        showFirstColumn=False,
                        showLastColumn=False,
                        showRowStripes=True,
                        showColumnStripes=False,
                    )
                )

                worksheet.add_table(
                    table
                )

        write_translator_legend(
            writer,
        )

    return merged_df


# ============================================================
# TRANSLATION SUMMARY HELPERS
# ============================================================

def _summary_clean(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).replace("\r\n", "\n").replace("\r", "\n").strip()
    if text.casefold() in {"", "null", "none", "nan"}:
        return ""
    return text


def _summary_word_count(value):
    """Count human-readable words; URLs are not split into punctuation fragments."""
    text = _summary_clean(value)
    if not text:
        return 0
    # Treat words/numbers containing apostrophes or hyphens as one word.
    return len(re.findall(r"\b[\wÀ-ÖØ-öø-ÿ]+(?:['’\-][\wÀ-ÖØ-öø-ÿ]+)*\b", text, flags=re.UNICODE))


def _summary_character_count(value):
    """Count characters in the cleaned value, including spaces and punctuation."""
    text = _summary_clean(value)
    return len(text) if text else 0


def build_translation_summary(content_df):
    """
    Return overview + distinct-value detail for translator workload.

    Final aggregation rule:
      * Detail table is the source of truth.
      * "Sum of Word Counts" = sum of detail "Word Count", grouped by "Column".
      * "Sum of Character Counts" = sum of detail "Character Count", grouped by "Column".
      * Occurrences are NOT multiplied into either total.
    """
    fields = [
        "Content Text Changes",
        "train_heading",
        "train_heading_sub_heading",
        "train_url_name",
    ]

    overview_rows = []
    detail_rows = []

    for field in fields:
        if field not in content_df.columns:
            overview_rows.append({
                "Column": field,
                "Total Rows": len(content_df),
                "Nonblank Rows": 0,
                "Blank Rows": len(content_df),
                "Distinct Nonblank Rows": 0,
                "Words - All Nonblank Rows": 0,
            })
            continue

        cleaned = content_df[field].map(_summary_clean)
        nonblank = cleaned[cleaned.ne("")]

        # Preserve first-seen order; distinct means exact unique cell value.
        distinct_values = list(dict.fromkeys(nonblank.tolist()))
        counts = nonblank.value_counts(sort=False)

        overview_rows.append({
            "Column": field,
            "Total Rows": len(content_df),
            "Nonblank Rows": len(nonblank),
            "Blank Rows": len(content_df) - len(nonblank),
            "Distinct Nonblank Rows": len(distinct_values),
            "Words - All Nonblank Rows": int(
                sum(_summary_word_count(v) for v in nonblank)
            ),
        })

        for distinct_number, value in enumerate(distinct_values, start=1):
            detail_rows.append({
                "Column": field,
                "Distinct #": distinct_number,
                "Distinct Value": value,
                "Occurrences": int(counts.get(value, 0)),
                "Word Count": int(_summary_word_count(value)),
                "Character Count": int(_summary_character_count(value)),
            })

    overview_df = pd.DataFrame(overview_rows)
    detail_df = pd.DataFrame(
        detail_rows,
        columns=[
            "Column",
            "Distinct #",
            "Distinct Value",
            "Occurrences",
            "Word Count",
            "Character Count",
        ],
    )

    # IMPORTANT: Build the two requested Summary totals directly from the
    # detail table. This is equivalent to grouping detail Column A and summing
    # Word Count (E) and Character Count (F).
    if detail_df.empty:
        totals_df = pd.DataFrame({
            "Column": fields,
            "Sum of Word Counts": [0] * len(fields),
            "Sum of Character Counts": [0] * len(fields),
        })
    else:
        totals_df = (
            detail_df
            .groupby("Column", sort=False, as_index=False)
            .agg(
                **{
                    "Sum of Word Counts": ("Word Count", "sum"),
                    "Sum of Character Counts": ("Character Count", "sum"),
                }
            )
        )

    overview_df = overview_df.merge(
        totals_df,
        on="Column",
        how="left",
    )

    for metric in ("Sum of Word Counts", "Sum of Character Counts"):
        overview_df[metric] = overview_df[metric].fillna(0).astype(int)

    overview_df = overview_df[[
        "Column",
        "Total Rows",
        "Nonblank Rows",
        "Blank Rows",
        "Distinct Nonblank Rows",
        "Words - All Nonblank Rows",
        "Sum of Word Counts",
        "Sum of Character Counts",
    ]]

    return overview_df, detail_df


def write_translator_note(writer):
    worksheet = writer.book.create_sheet(
        "Note to Translator",
    )

    note_rows = {
        2: ("Note to Translator", True, 14),
        3: ("Tab - French Content Translation", True, 14),
        4: (
            "Review the text in columns I, J, K, and L, and translate it as needed.",
            False,
            11,
        ),
        5: (
            "Enter each French translation in the corresponding colour-coded column:",
            False,
            11,
        ),
        6: ("Column I", True, 11),
        7: ("Column J", True, 11),
        8: ("Column K", True, 11),
        9: ("Column L", True, 11),
        11: (
            "Use columns M and N to sort or filter the content.",
            False,
            11,
        ),
        12: (
            "Translate repeated rows as well, and enter each translation in its corresponding cell.",
            False,
            11,
        ),
    }

    for row_number, (value, bold, font_size) in note_rows.items():
        cell = worksheet.cell(
            row=row_number,
            column=2,
            value=value,
        )
        cell.font = Font(
            name="Aptos",
            size=font_size,
            bold=bold,
        )
        cell.alignment = Alignment(
            vertical="top",
            wrap_text=True,
        )

    for row_number in (2, 3, 4, 5, 11, 12):
        worksheet.merge_cells(
            start_row=row_number,
            start_column=2,
            end_row=row_number,
            end_column=11,
        )

    for row_number, target_column in enumerate(
        ("O", "P", "Q", "R"),
        start=6,
    ):
        cell = worksheet.cell(
            row=row_number,
            column=3,
            value=f"→ Column {target_column}",
        )
        cell.font = Font(
            name="Aptos",
            size=11,
            bold=True,
        )
        cell.alignment = Alignment(
            horizontal="left",
        )

    for row_number in range(2, 13):
        worksheet.row_dimensions[row_number].height = 20
    worksheet.row_dimensions[5].height = 32
    worksheet.row_dimensions[12].height = 41


def write_translator_legend(writer):
    source = writer.book["French Content Translation"]
    target = writer.book["Note to Translator"]

    for column_number in range(9, 19):
        source_column = get_column_letter(column_number)
        target_column = column_number - 7
        source_cell = source.cell(
            row=1,
            column=column_number,
        )

        column_label = target.cell(
            row=14,
            column=target_column,
            value=source_column,
        )
        column_label.font = Font(
            name="Aptos",
            size=10,
            bold=True,
        )
        column_label.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

        legend_cell = target.cell(
            row=15,
            column=target_column,
            value=source_cell.value,
        )
        legend_cell.font = copy(source_cell.font)
        legend_cell.fill = copy(source_cell.fill)
        legend_cell.border = copy(source_cell.border)
        legend_cell.alignment = copy(source_cell.alignment)
        legend_cell.number_format = source_cell.number_format
        legend_cell.protection = copy(source_cell.protection)

        source_width = source.column_dimensions[source_column].width
        if source_width is not None:
            target_letter = get_column_letter(target_column)
            target.column_dimensions[target_letter].width = source_width

    target.row_dimensions[14].height = 20
    target.row_dimensions[15].height = 80


def write_translation_summary(writer, content_df):
    overview, details = build_translation_summary(content_df)
    sheet = "Summary"

    # Layout:
    #   row 3      overview header
    #   rows 4:7   overview data
    #   row 8      filter-aware subtotal
    #   row 9      fixed grand total
    #   row 12     detail header
    #   rows 13+   detail data
    #   next row   filter-aware subtotal
    #   next row   fixed grand total
    overview_header_row = 3
    overview_data_start = 4
    overview_data_end = overview_data_start + len(overview) - 1
    overview_subtotal_row = overview_data_end + 1
    overview_grand_row = overview_data_end + 2

    detail_header_row = overview_grand_row + 3
    detail_data_start = detail_header_row + 1
    detail_data_end = detail_data_start + len(details) - 1
    detail_subtotal_row = detail_data_end + 1
    detail_grand_row = detail_data_end + 2

    overview.to_excel(writer, sheet_name=sheet, index=False, startrow=overview_header_row - 1)
    details.to_excel(writer, sheet_name=sheet, index=False, startrow=detail_header_row - 1)

    ws = writer.book[sheet]
    ws["A1"] = "French Content Translation - Distinct Row, Word & Character Count Summary"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (
        "Distinct counts exclude blank/Null values. The detail table contains one row per distinct value. "
        "Sum of Word Counts is the grouped sum of detail Word Count; Sum of Character Counts is the grouped "
        "sum of detail Character Count. Filtered Subtotal changes with Excel filters; Grand Total is the full table total."
    )
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)

    # Overview totals. B:H are numeric columns.
    ws.cell(overview_subtotal_row, 1, "Filtered Subtotal")
    ws.cell(overview_grand_row, 1, "Grand Total")
    for col_idx in range(2, 9):
        col_letter = get_column_letter(col_idx)
        ws.cell(
            overview_subtotal_row,
            col_idx,
            f"=SUBTOTAL(109,{col_letter}{overview_data_start}:{col_letter}{overview_data_end})",
        )
        ws.cell(
            overview_grand_row,
            col_idx,
            f"=SUM({col_letter}{overview_data_start}:{col_letter}{overview_data_end})",
        )

    # Detail totals. Group/filter is Column A; D/E/F are additive.
    ws.cell(detail_subtotal_row, 1, "Filtered Subtotal")
    ws.cell(detail_grand_row, 1, "Grand Total")

    # Distinct # subtotal is visible detail-row count; grand total is all detail rows.
    ws.cell(
        detail_subtotal_row,
        2,
        f"=SUBTOTAL(103,C{detail_data_start}:C{detail_data_end})",
    )
    ws.cell(detail_grand_row, 2, len(details))

    # Occurrences, Word Count, Character Count.
    for col_idx in (4, 5, 6):
        col_letter = get_column_letter(col_idx)
        ws.cell(
            detail_subtotal_row,
            col_idx,
            f"=SUBTOTAL(109,{col_letter}{detail_data_start}:{col_letter}{detail_data_end})",
        )
        ws.cell(
            detail_grand_row,
            col_idx,
            f"=SUM({col_letter}{detail_data_start}:{col_letter}{detail_data_end})",
        )

    header_fill = PatternFill(fill_type="solid", fgColor="FF4472C4")
    subtotal_fill = PatternFill(fill_type="solid", fgColor="FFFFE699")
    grand_fill = PatternFill(fill_type="solid", fgColor="FFC6E0B4")

    for row_no in (overview_header_row, detail_header_row):
        for cell in ws[row_no]:
            if cell.value is not None:
                cell.font = Font(bold=True, color="FFFFFFFF")
                cell.fill = header_fill
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    for row_no, fill in (
        (overview_subtotal_row, subtotal_fill),
        (overview_grand_row, grand_fill),
        (detail_subtotal_row, subtotal_fill),
        (detail_grand_row, grand_fill),
    ):
        max_col = 8 if row_no <= overview_grand_row else 6
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row_no, col_idx)
            cell.font = Font(bold=True)
            cell.fill = fill

    ws.freeze_panes = f"A{detail_header_row}"
    widths = {"A": 34, "B": 16, "C": 70, "D": 14, "E": 16, "F": 25, "G": 28, "H": 32}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # Keep subtotal/grand-total rows OUTSIDE the Excel tables. This allows
    # SUBTOTAL formulas to remain visible and react to table filters.
    if len(overview):
        t1 = Table(
            displayName="TranslationSummaryTable",
            ref=f"A{overview_header_row}:H{overview_data_end}",
        )
        t1.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(t1)

    if len(details):
        t2 = Table(
            displayName="TranslationDistinctValuesTable",
            ref=f"A{detail_header_row}:F{detail_data_end}",
        )
        t2.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(t2)

    # Console verification values (Python-side, independent of Excel formulas).
    print()
    print("=" * 70)
    print("SUMMARY TOTALS")
    print("=" * 70)
    print("Grand Total - Total Rows              :", int(overview["Total Rows"].sum()))
    print("Grand Total - Nonblank Rows           :", int(overview["Nonblank Rows"].sum()))
    print("Grand Total - Blank Rows              :", int(overview["Blank Rows"].sum()))
    print("Grand Total - Distinct Nonblank Rows  :", int(overview["Distinct Nonblank Rows"].sum()))
    print("Grand Total - Words All Nonblank Rows :", int(overview["Words - All Nonblank Rows"].sum()))
    print("Grand Total - Sum of Word Counts      :", int(details["Word Count"].sum()))
    print("Grand Total - Sum of Character Counts :", int(details["Character Count"].sum()))

# ============================================================
# FILE HELPERS
# ============================================================

def remove_generated_file(
    path: Path,
):

    path = Path(
        path
    )

    if path.exists():

        path.unlink()

        print(
            f"Removed: {path}"
        )


def verify_generated_file(
    path: Path,
):

    path = Path(
        path
    )

    if not path.exists():

        print(
            f"MISSING: {path}"
        )

        return False

    modified = (
        datetime.fromtimestamp(
            path.stat().st_mtime
        )
        .strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )

    print(
        f"OK: {path}"
    )

    print(
        f"    Size: "
        f"{path.stat().st_size:,} bytes"
    )

    print(
        f"    Modified: "
        f"{modified}"
    )

    return True


# ============================================================
# MAIN
# ============================================================

def main():

    print()

    print(
        "=" * 70
    )

    print(
        "CMS DATA EXPORT / NORMALIZATION / TRANSFORMATION"
    )

    print(
        "=" * 70
    )

    # ========================================================
    # ROOT
    # ========================================================

    data_root = (
        Path(
            __file__
        )
        .resolve()
        .parents[
            1
        ]
    )

    # ========================================================
    # SOURCE JSON
    # ========================================================

    train_json = (
        data_root
        / "train_content"
        / "entities"
        / "localised_contents.json"
    )

    prod_json = (
        data_root
        / "prod_content"
        / "entities"
        / "localised_contents.json"
    )

    # ========================================================
    # VERIFY SOURCE
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "SOURCE FILES"
    )

    print(
        "=" * 70
    )

    for source_file in [
        train_json,
        prod_json,
    ]:

        if not source_file.exists():

            raise FileNotFoundError(
                "\nSource JSON not found:\n"
                f"{source_file}"
            )

        print()

        print(
            source_file
        )

        print(
            "Modified:",
            datetime.fromtimestamp(
                source_file.stat().st_mtime
            ).strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
        )

        print(
            "Size:",
            f"{source_file.stat().st_size:,}",
            "bytes",
        )

    # ========================================================
    # OUTPUT DIRECTORIES
    # ========================================================

    train_output_dir = (
        data_root
        / "train_output"
    )

    prod_output_dir = (
        data_root
        / "prod_output"
    )

    transformations_dir = (
        data_root
        / "transformations"
    )

    for directory in [
        train_output_dir,
        prod_output_dir,
        transformations_dir,
    ]:

        directory.mkdir(
            parents=True,
            exist_ok=True,
        )

    # ========================================================
    # OUTPUT FILES
    # ========================================================

    train_english_output = (
        train_output_dir
        / "train_export_details_normalised.csv"
    )

    train_french_output = (
        train_output_dir
        / "train_export_details_french_normalised.csv"
    )

    prod_english_output = (
        prod_output_dir
        / "prod_export_details_normalised.csv"
    )

    prod_french_output = (
        prod_output_dir
        / "prod_export_details_french_normalised.csv"
    )

    comparison_csv = (
        transformations_dir
        / "cms_data_comparing_train_prod.csv"
    )

    comparison_xlsx = (
        transformations_dir
        / "cms_data_comparing_train_prod.xlsx"
    )

    generated_files = [
        train_english_output,
        train_french_output,
        prod_english_output,
        prod_french_output,
        comparison_csv,
        comparison_xlsx,
    ]

    # ========================================================
    # REMOVE OLD OUTPUTS
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "REMOVING PREVIOUS GENERATED OUTPUTS"
    )

    print(
        "=" * 70
    )

    for output_file in (
        generated_files
    ):

        remove_generated_file(
            output_file
        )

    # ========================================================
    # TRAIN ENGLISH
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "NORMALIZING TRAIN - ENGLISH"
    )

    print(
        "=" * 70
    )

    train_english_df = (
        process_file(
            train_json,
            train_english_output,
            locale=(
                "en-CA"
            ),
        )
    )

    # ========================================================
    # TRAIN FRENCH
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "NORMALIZING TRAIN - FRENCH"
    )

    print(
        "=" * 70
    )

    train_french_df = (
        process_file(
            train_json,
            train_french_output,
            locale=(
                "fr-CA"
            ),
        )
    )

    # ========================================================
    # PROD ENGLISH
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "NORMALIZING PROD - ENGLISH"
    )

    print(
        "=" * 70
    )

    prod_english_df = (
        process_file(
            prod_json,
            prod_english_output,
            locale=(
                "en-CA"
            ),
        )
    )

    # ========================================================
    # PROD FRENCH
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "NORMALIZING PROD - FRENCH"
    )

    print(
        "=" * 70
    )

    prod_french_df = (
        process_file(
            prod_json,
            prod_french_output,
            locale=(
                "fr-CA"
            ),
        )
    )

    # ========================================================
    # SUMMARY
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "NORMALIZATION SUMMARY"
    )

    print(
        "=" * 70
    )

    print(
        f"Train English : "
        f"{len(train_english_df)}"
    )

    print(
        f"Train French  : "
        f"{len(train_french_df)}"
    )

    print(
        f"Prod English  : "
        f"{len(prod_english_df)}"
    )

    print(
        f"Prod French   : "
        f"{len(prod_french_df)}"
    )

    if (
        train_english_df.empty
    ):

        raise RuntimeError(
            "Train English contains 0 records."
        )

    if (
        prod_english_df.empty
    ):

        raise RuntimeError(
            "Prod English contains 0 records."
        )

    # ========================================================
    # TRANSFORMATION
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "CREATING TRAIN VS PROD TRANSFORMATION"
    )

    print(
        "=" * 70
    )

    print()

    print(
        "Train source:"
    )

    print(
        train_english_output
    )

    print()

    print(
        "Prod source:"
    )

    print(
        prod_english_output
    )

    print()

    print(
        "Transformation output:"
    )

    print(
        comparison_csv
    )

    merged_df = (
        join_export_files(
            train_english_output,
            prod_english_output,
            comparison_csv,
        )
    )

    print()

    print(
        "Comparison rows:",
        len(
            merged_df
        ),
    )

    # ========================================================
    # VERIFY OUTPUTS
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "VERIFYING GENERATED FILES"
    )

    print(
        "=" * 70
    )

    missing_outputs = []

    for output_file in (
        generated_files
    ):

        print()

        if not verify_generated_file(
            output_file
        ):

            missing_outputs.append(
                output_file
            )

    if missing_outputs:

        raise RuntimeError(
            "\nExpected output files "
            "were not generated:\n\n"
            +
            "\n".join(
                str(path)
                for path
                in missing_outputs
            )
        )

    # ========================================================
    # VERIFY ROW COUNTS
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "VERIFYING NORMALIZED ROW COUNTS"
    )

    print(
        "=" * 70
    )

    verification_files = [
        (
            "Train English",
            train_english_output,
            len(
                train_english_df
            ),
        ),
        (
            "Train French",
            train_french_output,
            len(
                train_french_df
            ),
        ),
        (
            "Prod English",
            prod_english_output,
            len(
                prod_english_df
            ),
        ),
        (
            "Prod French",
            prod_french_output,
            len(
                prod_french_df
            ),
        ),
    ]

    for (
        name,
        file_path,
        expected_rows,
    ) in verification_files:

        check_df = (
            pd.read_csv(
                file_path,
                dtype=object,
            )
        )

        actual_rows = (
            len(
                check_df
            )
        )

        if (
            actual_rows
            !=
            expected_rows
        ):

            raise RuntimeError(
                f"{name}: expected "
                f"{expected_rows} rows, "
                f"found {actual_rows}."
            )

        print(
            f"{name:<15}: "
            f"PASS "
            f"({actual_rows} rows)"
        )

    # ========================================================
    # COMPLETE
    # ========================================================

    print()

    print(
        "=" * 70
    )

    print(
        "STAGE 1 COMPLETE"
    )

    print(
        "=" * 70
    )

    print()

    print(
        "Train English:"
    )

    print(
        train_english_output
    )

    print()

    print(
        "Train French:"
    )

    print(
        train_french_output
    )

    print()

    print(
        "Prod English:"
    )

    print(
        prod_english_output
    )

    print()

    print(
        "Prod French:"
    )

    print(
        prod_french_output
    )

    print()

    print(
        "Transformation CSV:"
    )

    print(
        comparison_csv
    )

    print()

    print(
        "Transformation Excel:"
    )

    print(
        comparison_xlsx
    )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()
