import ast
import json
import re
from copy import deepcopy
from datetime import datetime
from pathlib import Path

import pandas as pd

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


# ============================================================
# CONFIGURATION
# ============================================================

DATA_ROOT = (
    Path(__file__)
    .resolve()
    .parents[1]
)

ENGLISH_TRAIN_FILE = (
    DATA_ROOT
    / "train_output"
    / "train_export_details_normalised.csv"
)

FRENCH_TRAIN_FILE = (
    DATA_ROOT
    / "train_output"
    / "train_export_details_french_normalised.csv"
)

TRANSLATION_FILE = (
    DATA_ROOT
    / "French Translation File"
    / "cms_data_comparing_train_prod_French_Translation.xlsx"
)

FINAL_OUTPUT_DIR = (
    DATA_ROOT
    / "final_output"
)

FINAL_CSV = (
    FINAL_OUTPUT_DIR
    / "train_export_details_french_updated.csv"
)

FINAL_XLSX = (
    FINAL_OUTPUT_DIR
    / "train_export_details_french_updated.xlsx"
)

AUDIT_XLSX = (
    FINAL_OUTPUT_DIR
    / "french_translation_update_audit.xlsx"
)


# ============================================================
# COLUMN NAMES
# ============================================================

COL_UUID = "__uuid"
COL_LOCALIZATIONS = "localizations"
COL_LOCALE = "locale"

COL_ROW_KEY = "Row Key"
COL_COMMUNITY = "Local Community Name"
COL_TITLE = "title"

COL_CONTENTS = "contents"
COL_BUTTON_LINK = "buttonLink"

COL_TRAIN_HEADING = "train_heading"
COL_TRAIN_URL_NAME = "train_url_name"
COL_TRAIN_URL = "train_url"

COL_CONTENT_CHANGES = "Content Text Changes"
COL_CONTEXT_BEFORE = "Content Text Context Before"
COL_CONTEXT_AFTER = "Content Text Context After"
COL_BLOCK_COUNT = "Content Text Change Block Count"

COL_FRENCH_HEADING = "French translation of column B"
COL_FRENCH_URL_NAME = "French translation of column C"
COL_FRENCH_CONTENT = "Content Text Changes French"

COL_FRENCH_BUTTON_LINK = (
    "French Button Link corresponding to column D"
)

COL_FRENCH_CONTENT_URL = (
    "French Content URL corresponding to column J"
)


# ============================================================
# REGEX
# ============================================================

MARKDOWN_LINK_RE = re.compile(
    r"\[([^\]]+)\]\((https?://[^)\s]+)\)"
)

BARE_URL_RE = re.compile(
    r"https?://[^\s)\]>]+"
)


# ============================================================
# BASIC HELPERS
# ============================================================

def banner(text):

    print()
    print("=" * 70)
    print(text)
    print("=" * 70)


def clean_value(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = str(value).strip()

    if text.lower() in {
        "",
        "nan",
        "none",
        "null",
    }:
        return ""

    return text


def normalize_text(value):

    text = clean_value(value)

    text = (
        text
        .replace("\r\n", "\n")
        .replace("\r", "\n")
    )

    text = re.sub(
        r"[ \t]+",
        " ",
        text,
    )

    return text.strip()


def normalized_key(value):

    text = normalize_text(value)

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return (
        text
        .casefold()
        .strip()
    )


def values_equal(
    left,
    right,
):

    return (
        normalize_text(left)
        ==
        normalize_text(right)
    )


# ============================================================
# LOCALIZATION UUID PARSING
# ============================================================

def extract_localization_uuids(value):
    """
    Parse:

        [{'__uuid': 'abc'}]

    and return:

        ['abc']
    """

    if isinstance(
        value,
        list,
    ):

        parsed = value

    else:

        text = clean_value(value)

        if not text:
            return []

        try:

            parsed = ast.literal_eval(
                text
            )

        except (
            ValueError,
            SyntaxError,
        ):

            try:

                parsed = json.loads(
                    text
                )

            except (
                json.JSONDecodeError,
                TypeError,
            ):

                return []

    if not isinstance(
        parsed,
        list,
    ):
        return []

    uuids = []

    for item in parsed:

        if not isinstance(
            item,
            dict,
        ):
            continue

        uuid_value = clean_value(
            item.get(
                "__uuid"
            )
        )

        if uuid_value:

            uuids.append(
                uuid_value
            )

    return uuids


# ============================================================
# CONTENT PARSING
# ============================================================

def parse_contents(value):

    if isinstance(
        value,
        list,
    ):

        return deepcopy(
            value
        )

    text = clean_value(value)

    if text in {
        "",
        "[]",
        "Null",
    }:

        return []

    try:

        parsed = ast.literal_eval(
            text
        )

        if isinstance(
            parsed,
            list,
        ):

            return deepcopy(
                parsed
            )

    except (
        ValueError,
        SyntaxError,
    ):
        pass

    try:

        parsed = json.loads(
            text
        )

        if isinstance(
            parsed,
            list,
        ):

            return deepcopy(
                parsed
            )

    except (
        json.JSONDecodeError,
        TypeError,
    ):
        pass

    return []


def serialize_contents(contents):

    if not contents:
        return "[]"

    return repr(
        contents
    )


# ============================================================
# STAGE 1 PLAIN-TEXT LOGIC
# ============================================================

def plain_line(line):

    line = (
        ""
        if line is None
        else str(line)
    )

    line = MARKDOWN_LINK_RE.sub(
        lambda match:
            match.group(1),
        line,
    )

    line = BARE_URL_RE.sub(
        "",
        line,
    )

    line = re.sub(
        r"\*+",
        "",
        line,
    )

    line = re.sub(
        r"^\s*[-*+]\s+",
        "",
        line,
    )

    line = re.sub(
        r"\s+",
        " ",
        line,
    )

    return line.strip()


def content_plain_lines(content):

    raw_lines = (
        ""
        if content is None
        else str(content)
    ).splitlines()

    output = []

    for (
        raw_index,
        raw_line,
    ) in enumerate(
        raw_lines
    ):

        plain = plain_line(
            raw_line
        )

        if not plain:
            continue

        output.append(
            {
                "raw_index":
                    raw_index,

                "raw":
                    raw_line,

                "plain":
                    plain,
            }
        )

    return output


# ============================================================
# FIND TEXT SEQUENCE
# ============================================================

def find_sequence(
    haystack,
    needle,
):

    if not needle:
        return []

    if (
        len(needle)
        >
        len(haystack)
    ):

        return []

    haystack_keys = [
        normalized_key(value)
        for value
        in haystack
    ]

    needle_keys = [
        normalized_key(value)
        for value
        in needle
    ]

    size = len(
        needle_keys
    )

    matches = []

    for start in range(
        0,
        len(haystack_keys)
        -
        size
        +
        1,
    ):

        if (
            haystack_keys[
                start:
                start + size
            ]
            ==
            needle_keys
        ):

            matches.append(
                start
            )

    return matches


# ============================================================
# SPLIT MULTIPLE CHANGE BLOCKS
# ============================================================

def split_blocks(
    value,
    expected_count=0,
):

    text = clean_value(value)

    if not text:
        return []

    if expected_count <= 1:

        return [
            text
        ]

    blocks = [
        block.strip()
        for block
        in re.split(
            r"\n\s*\n+",
            text,
        )
        if block.strip()
    ]

    return blocks


# ============================================================
# BUILD UUID MAPS
# ============================================================

def build_uuid_maps(
    english_df,
    french_df,
):
    """
    Main Stage 2 mapping:

    English __uuid
            ↓
    French localizations
            ↓
    French dataframe index
    """

    english_uuid_to_index = {}

    for (
        index,
        row,
    ) in english_df.iterrows():

        uuid_value = clean_value(
            row.get(
                COL_UUID
            )
        )

        if uuid_value:

            english_uuid_to_index[
                uuid_value
            ] = index

    french_uuid_to_index = {}

    english_uuid_to_french_indexes = {}

    for (
        index,
        row,
    ) in french_df.iterrows():

        french_uuid = clean_value(
            row.get(
                COL_UUID
            )
        )

        if french_uuid:

            french_uuid_to_index[
                french_uuid
            ] = index

        localization_uuids = (
            extract_localization_uuids(
                row.get(
                    COL_LOCALIZATIONS
                )
            )
        )

        for english_uuid in (
            localization_uuids
        ):

            english_uuid_to_french_indexes.setdefault(
                english_uuid,
                []
            ).append(
                index
            )

    return (
        english_uuid_to_index,
        french_uuid_to_index,
        english_uuid_to_french_indexes,
    )


# ============================================================
# FIND ENGLISH RECORD
# ============================================================

def find_english_record(
    english_df,
    comparison_row,
):
    """
    The Comparison row was created from the English Train
    normalized record.

    Prefer train___uuid when available.

    Fallback:
        English Local Community Name + title
    """

    possible_uuid_columns = [
        "train___uuid",
        "__uuid",
        "train_uuid",
    ]

    for column in possible_uuid_columns:

        if (
            column
            in comparison_row
        ):

            uuid_value = clean_value(
                comparison_row.get(
                    column
                )
            )

            if uuid_value:

                matches = (
                    english_df[
                        COL_UUID
                    ]
                    .fillna("")
                    .astype(str)
                    .eq(
                        uuid_value
                    )
                )

                indexes = list(
                    english_df.index[
                        matches
                    ]
                )

                if len(indexes) == 1:

                    return (
                        indexes[0],
                        "ENGLISH_UUID",
                    )

    community = clean_value(
        comparison_row.get(
            COL_COMMUNITY
        )
    )

    title = clean_value(
        comparison_row.get(
            COL_TITLE
        )
    )

    if (
        COL_COMMUNITY
        in english_df.columns
        and
        COL_TITLE
        in english_df.columns
    ):

        community_key = (
            normalized_key(
                community
            )
        )

        title_key = (
            normalized_key(
                title
            )
        )

        matches = (
            english_df[
                COL_COMMUNITY
            ]
            .fillna("")
            .astype(str)
            .map(
                normalized_key
            )
            .eq(
                community_key
            )
            &
            english_df[
                COL_TITLE
            ]
            .fillna("")
            .astype(str)
            .map(
                normalized_key
            )
            .eq(
                title_key
            )
        )

        indexes = list(
            english_df.index[
                matches
            ]
        )

        if len(indexes) == 1:

            return (
                indexes[0],
                "ENGLISH_COMMUNITY_TITLE",
            )

        if len(indexes) > 1:

            return (
                None,
                "AMBIGUOUS_ENGLISH_RECORD",
            )

    return (
        None,
        "ENGLISH_RECORD_NOT_FOUND",
    )


# ============================================================
# RESOLVE EXACT FRENCH RECORD
# ============================================================

def resolve_french_record(
    english_df,
    french_df,
    comparison_row,
    english_uuid_to_french_indexes,
):
    """
    Resolve:

        Comparison row
            ↓
        English record
            ↓
        English __uuid
            ↓
        French localizations
            ↓
        exact French record
    """

    (
        english_index,
        english_method,
    ) = find_english_record(
        english_df,
        comparison_row,
    )

    if english_index is None:

        return {
            "english_index":
                None,

            "english_uuid":
                "",

            "french_index":
                None,

            "status":
                english_method,

            "method":
                english_method,
        }

    english_uuid = clean_value(
        english_df.at[
            english_index,
            COL_UUID
        ]
    )

    if not english_uuid:

        return {
            "english_index":
                english_index,

            "english_uuid":
                "",

            "french_index":
                None,

            "status":
                "ENGLISH_UUID_MISSING",

            "method":
                english_method,
        }

    candidates = (
        english_uuid_to_french_indexes.get(
            english_uuid,
            []
        )
    )

    if len(candidates) == 1:

        return {
            "english_index":
                english_index,

            "english_uuid":
                english_uuid,

            "french_index":
                candidates[0],

            "status":
                "FOUND",

            "method":
                (
                    english_method
                    +
                    " -> LOCALIZATION_UUID"
                ),
        }

    if len(candidates) > 1:

        return {
            "english_index":
                english_index,

            "english_uuid":
                english_uuid,

            "french_index":
                None,

            "status":
                "AMBIGUOUS_FRENCH_LOCALIZATION",

            "method":
                (
                    english_method
                    +
                    " -> LOCALIZATION_UUID"
                ),
        }

    # ========================================================
    # REVERSE CHECK
    #
    # Some exports may expose localization relation only
    # from English -> French.
    # ========================================================

    english_localizations = (
        extract_localization_uuids(
            english_df.at[
                english_index,
                COL_LOCALIZATIONS
            ]
        )
        if (
            COL_LOCALIZATIONS
            in english_df.columns
        )
        else []
    )

    reverse_candidates = []

    for french_uuid in (
        english_localizations
    ):

        matches = (
            french_df[
                COL_UUID
            ]
            .fillna("")
            .astype(str)
            .eq(
                french_uuid
            )
        )

        indexes = list(
            french_df.index[
                matches
            ]
        )

        reverse_candidates.extend(
            indexes
        )

    reverse_candidates = sorted(
        set(
            reverse_candidates
        )
    )

    if len(reverse_candidates) == 1:

        return {
            "english_index":
                english_index,

            "english_uuid":
                english_uuid,

            "french_index":
                reverse_candidates[0],

            "status":
                "FOUND",

            "method":
                (
                    english_method
                    +
                    " -> ENGLISH_LOCALIZATION_UUID"
                ),
        }

    if len(reverse_candidates) > 1:

        return {
            "english_index":
                english_index,

            "english_uuid":
                english_uuid,

            "french_index":
                None,

            "status":
                "AMBIGUOUS_REVERSE_LOCALIZATION",

            "method":
                (
                    english_method
                    +
                    " -> ENGLISH_LOCALIZATION_UUID"
                ),
        }

    return {
        "english_index":
            english_index,

        "english_uuid":
            english_uuid,

        "french_index":
            None,

        "status":
            "FRENCH_LOCALIZATION_NOT_FOUND",

        "method":
            (
                english_method
                +
                " -> LOCALIZATION_UUID"
            ),
    }


# ============================================================
# CONTENT ITEM MATCHING
# ============================================================

def find_content_item_index(
    contents,
    heading="",
    url_name="",
    url="",
):

    heading_key = normalized_key(
        heading
    )

    url_name_key = normalized_key(
        url_name
    )

    url = clean_value(
        url
    )

    # ========================================================
    # URL
    # ========================================================

    if url:

        candidates = []

        for (
            index,
            item,
        ) in enumerate(
            contents
        ):

            if not isinstance(
                item,
                dict,
            ):
                continue

            content = str(
                item.get(
                    "content"
                )
                or ""
            )

            if url in content:

                candidates.append(
                    index
                )

        candidates = sorted(
            set(
                candidates
            )
        )

        if len(candidates) == 1:

            return (
                candidates[0],
                "URL",
            )

    # ========================================================
    # URL NAME
    # ========================================================

    if url_name_key:

        candidates = []

        for (
            index,
            item,
        ) in enumerate(
            contents
        ):

            if not isinstance(
                item,
                dict,
            ):
                continue

            content = str(
                item.get(
                    "content"
                )
                or ""
            )

            for match in (
                MARKDOWN_LINK_RE.finditer(
                    content
                )
            ):

                if (
                    normalized_key(
                        match.group(1)
                    )
                    ==
                    url_name_key
                ):

                    candidates.append(
                        index
                    )

                    break

        candidates = sorted(
            set(
                candidates
            )
        )

        if len(candidates) == 1:

            return (
                candidates[0],
                "URL_NAME",
            )

    # ========================================================
    # HEADING
    # ========================================================

    if heading_key:

        candidates = []

        for (
            index,
            item,
        ) in enumerate(
            contents
        ):

            if not isinstance(
                item,
                dict,
            ):
                continue

            if (
                normalized_key(
                    item.get(
                        "heading"
                    )
                )
                ==
                heading_key
            ):

                candidates.append(
                    index
                )

        if len(candidates) == 1:

            return (
                candidates[0],
                "HEADING",
            )

    # ========================================================
    # SINGLE ITEM
    # ========================================================

    dict_items = [
        index
        for (
            index,
            item,
        ) in enumerate(
            contents
        )
        if isinstance(
            item,
            dict,
        )
    ]

    if len(dict_items) == 1:

        return (
            dict_items[0],
            "SINGLE_ITEM",
        )

    return (
        None,
        "CONTENT_ITEM_NOT_FOUND",
    )


# ============================================================
# LOCATE ENGLISH CHANGE BLOCK
# ============================================================

def locate_change_block(
    english_content,
    changed_text,
    context_before="",
    context_after="",
):

    logical = content_plain_lines(
        english_content
    )

    plain_values = [
        item[
            "plain"
        ]
        for item
        in logical
    ]

    changed_lines = [
        line.strip()
        for line
        in normalize_text(
            changed_text
        ).splitlines()
        if line.strip()
    ]

    if not changed_lines:

        return {
            "status":
                "EMPTY_CHANGE"
        }

    starts = find_sequence(
        plain_values,
        changed_lines,
    )

    before_key = normalized_key(
        context_before
    )

    after_key = normalized_key(
        context_after
    )

    if (
        len(starts) > 1
        and
        (
            before_key
            or
            after_key
        )
    ):

        filtered = []

        for start in starts:

            end = (
                start
                +
                len(
                    changed_lines
                )
            )

            before_ok = True
            after_ok = True

            if before_key:

                before_ok = (
                    start > 0
                    and
                    normalized_key(
                        plain_values[
                            start - 1
                        ]
                    )
                    ==
                    before_key
                )

            if after_key:

                after_ok = (
                    end
                    <
                    len(
                        plain_values
                    )
                    and
                    normalized_key(
                        plain_values[
                            end
                        ]
                    )
                    ==
                    after_key
                )

            if (
                before_ok
                and
                after_ok
            ):

                filtered.append(
                    start
                )

        if filtered:

            starts = filtered

    if not starts:

        return {
            "status":
                "NOT_FOUND"
        }

    if len(starts) > 1:

        return {
            "status":
                "AMBIGUOUS"
        }

    logical_start = (
        starts[0]
    )

    logical_end = (
        logical_start
        +
        len(
            changed_lines
        )
    )

    return {
        "status":
            "FOUND",

        "logical_start":
            logical_start,

        "logical_end":
            logical_end,

        "raw_start":
            logical[
                logical_start
            ][
                "raw_index"
            ],

        "raw_end":
            (
                logical[
                    logical_end - 1
                ][
                    "raw_index"
                ]
                +
                1
            ),
    }


# ============================================================
# PRESERVE MARKDOWN STRUCTURE
# ============================================================

def apply_translation_to_raw_lines(
    old_lines,
    translated_text,
):

    translated_lines = [
        line
        for line
        in (
            translated_text
            .replace(
                "\r\n",
                "\n",
            )
            .replace(
                "\r",
                "\n",
            )
            .splitlines()
        )
        if line.strip()
    ]

    if not translated_lines:

        return (
            old_lines,
            False,
            "EMPTY_TRANSLATION",
        )

    # ========================================================
    # SAME NUMBER OF LINES
    #
    # Preserve Markdown URLs and prefixes.
    # ========================================================

    if (
        len(old_lines)
        ==
        len(translated_lines)
    ):

        replacement = []

        for (
            old_line,
            translated_line,
        ) in zip(
            old_lines,
            translated_lines,
        ):

            old_markdown = (
                MARKDOWN_LINK_RE.search(
                    old_line
                )
            )

            if old_markdown:

                translated_markdown = (
                    MARKDOWN_LINK_RE.search(
                        translated_line
                    )
                )

                if translated_markdown:

                    new_label = (
                        translated_markdown.group(
                            1
                        )
                    )

                else:

                    new_label = (
                        translated_line.strip()
                    )

                    new_label = re.sub(
                        r"^\s*[-*+]\s+",
                        "",
                        new_label,
                    )

                prefix = (
                    old_line[
                        :old_markdown.start()
                    ]
                )

                suffix = (
                    old_line[
                        old_markdown.end():
                    ]
                )

                replacement.append(
                    prefix
                    +
                    f"[{new_label}]"
                    f"({old_markdown.group(2)})"
                    +
                    suffix
                )

                continue

            prefix_match = re.match(
                r"^(\s*(?:#{1,6}\s*|[-*+]\s+)?)",
                old_line,
            )

            prefix = (
                prefix_match.group(1)
                if prefix_match
                else ""
            )

            translated_clean = (
                translated_line.strip()
            )

            if (
                prefix
                and
                translated_clean.startswith(
                    prefix.strip()
                )
            ):

                replacement.append(
                    translated_clean
                )

            else:

                replacement.append(
                    prefix
                    +
                    translated_clean
                )

        return (
            replacement,
            True,
            "UPDATED_PRESERVED_STRUCTURE",
        )

    # ========================================================
    # DIFFERENT LINE COUNT
    #
    # Use translator text directly.
    # ========================================================

    return (
        translated_lines,
        True,
        "UPDATED_TRANSLATION_STRUCTURE",
    )


# ============================================================
# UPDATE EXISTING FRENCH CONTENT
# ============================================================

def update_existing_french_content(
    english_content,
    french_content,
    changed_text,
    french_translation,
    context_before="",
    context_after="",
):

    location = locate_change_block(
        english_content,
        changed_text,
        context_before,
        context_after,
    )

    if (
        location[
            "status"
        ]
        !=
        "FOUND"
    ):

        return (
            french_content,
            False,
            (
                "ENGLISH_BLOCK_"
                +
                location[
                    "status"
                ]
            ),
        )

    english_logical = (
        content_plain_lines(
            english_content
        )
    )

    french_logical = (
        content_plain_lines(
            french_content
        )
    )

    logical_start = (
        location[
            "logical_start"
        ]
    )

    logical_end = (
        location[
            "logical_end"
        ]
    )

    if (
        logical_end
        >
        len(
            french_logical
        )
    ):

        return (
            french_content,
            False,
            "FRENCH_STRUCTURE_TOO_SHORT",
        )

    french_raw_start = (
        french_logical[
            logical_start
        ][
            "raw_index"
        ]
    )

    french_raw_end = (
        french_logical[
            logical_end - 1
        ][
            "raw_index"
        ]
        +
        1
    )

    raw_lines = (
        french_content.splitlines()
    )

    old_lines = (
        raw_lines[
            french_raw_start:
            french_raw_end
        ]
    )

    (
        replacement_lines,
        success,
        status,
    ) = apply_translation_to_raw_lines(
        old_lines,
        french_translation,
    )

    if not success:

        return (
            french_content,
            False,
            status,
        )

    raw_lines[
        french_raw_start:
        french_raw_end
    ] = (
        replacement_lines
    )

    return (
        "\n".join(
            raw_lines
        ),
        True,
        status,
    )


# ============================================================
# EMPTY FRENCH CONTENT
# ============================================================

def build_content_from_english_template(
    english_content,
    changed_text,
    french_translation,
    context_before="",
    context_after="",
):
    """
    French contents is [].

    Use English content only as the structural template.

    IMPORTANT:
    Only the translated changed block is replaced.

    Other English text remains in the template. This is
    explicitly flagged as PARTIAL_TEMPLATE_BUILD in the audit
    so it can be reviewed before CMS import.
    """

    location = locate_change_block(
        english_content,
        changed_text,
        context_before,
        context_after,
    )

    if (
        location[
            "status"
        ]
        !=
        "FOUND"
    ):

        return (
            "",
            False,
            (
                "ENGLISH_BLOCK_"
                +
                location[
                    "status"
                ]
            ),
        )

    raw_lines = (
        english_content.splitlines()
    )

    raw_start = (
        location[
            "raw_start"
        ]
    )

    raw_end = (
        location[
            "raw_end"
        ]
    )

    old_lines = (
        raw_lines[
            raw_start:
            raw_end
        ]
    )

    (
        replacement_lines,
        success,
        status,
    ) = apply_translation_to_raw_lines(
        old_lines,
        french_translation,
    )

    if not success:

        return (
            "",
            False,
            status,
        )

    raw_lines[
        raw_start:
        raw_end
    ] = (
        replacement_lines
    )

    return (
        "\n".join(
            raw_lines
        ),
        True,
        (
            "PARTIAL_TEMPLATE_BUILD / "
            +
            status
        ),
    )


# ============================================================
# MARKDOWN LABEL UPDATE
# ============================================================

def replace_markdown_label(
    content,
    old_label,
    new_label,
    preferred_url="",
):

    old_key = normalized_key(
        old_label
    )

    new_label = clean_value(
        new_label
    )

    preferred_url = clean_value(
        preferred_url
    )

    if not new_label:

        return (
            content,
            False,
            "EMPTY_TRANSLATION",
        )

    matches = []

    for match in (
        MARKDOWN_LINK_RE.finditer(
            content
        )
    ):

        if (
            normalized_key(
                match.group(1)
            )
            !=
            old_key
        ):

            continue

        if (
            preferred_url
            and
            match.group(2)
            !=
            preferred_url
        ):

            continue

        matches.append(
            match
        )

    if (
        not matches
        and
        preferred_url
    ):

        for match in (
            MARKDOWN_LINK_RE.finditer(
                content
            )
        ):

            if (
                normalized_key(
                    match.group(1)
                )
                ==
                old_key
            ):

                matches.append(
                    match
                )

    if not matches:

        return (
            content,
            False,
            "LABEL_NOT_FOUND",
        )

    if len(matches) > 1:

        return (
            content,
            False,
            "AMBIGUOUS_LABEL",
        )

    match = (
        matches[0]
    )

    replacement = (
        f"[{new_label}]"
        f"({match.group(2)})"
    )

    return (
        content[
            :match.start()
        ]
        +
        replacement
        +
        content[
            match.end():
        ],
        True,
        "UPDATED",
    )


# ============================================================
# MARKDOWN URL UPDATE
# ============================================================

def replace_markdown_url(
    content,
    old_url,
    new_url,
):

    old_url = clean_value(
        old_url
    )

    new_url = clean_value(
        new_url
    )

    if not new_url:

        return (
            content,
            False,
            "EMPTY_NEW_URL",
        )

    matches = [
        match
        for match
        in MARKDOWN_LINK_RE.finditer(
            content
        )
        if (
            match.group(2)
            ==
            old_url
        )
    ]

    if not matches:

        return (
            content,
            False,
            "URL_NOT_FOUND",
        )

    if len(matches) > 1:

        return (
            content,
            False,
            "AMBIGUOUS_URL",
        )

    match = (
        matches[0]
    )

    replacement = (
        f"[{match.group(1)}]"
        f"({new_url})"
    )

    return (
        content[
            :match.start()
        ]
        +
        replacement
        +
        content[
            match.end():
        ],
        True,
        "UPDATED",
    )


# ============================================================
# COMPARISON LOOKUP
# ============================================================

def build_comparison_lookup(
    comparison_df,
):

    if (
        COL_ROW_KEY
        not in comparison_df.columns
    ):

        raise ValueError(
            'Comparison sheet is missing "Row Key".'
        )

    lookup = {}

    for (
        _,
        row,
    ) in comparison_df.iterrows():

        row_key = clean_value(
            row.get(
                COL_ROW_KEY
            )
        )

        if row_key:

            lookup[
                row_key
            ] = (
                row.to_dict()
            )

    return lookup


# ============================================================
# AUDIT
# ============================================================

AUDIT_COLUMNS = [
    "Timestamp",
    "Source Sheet",
    "Row Key",
    "English UUID",
    "French UUID",
    "English Index",
    "French Index",
    "Local Community Name",
    "title",
    "Match Method",
    "Update Type",
    "Source Value",
    "Customer Value",
    "Old French Value",
    "New French Value",
    "Status",
    "Details",
]


def add_audit(
    audit_rows,
    *,
    source_sheet,
    row_key,
    english_uuid="",
    french_uuid="",
    english_index=None,
    french_index=None,
    community="",
    title="",
    match_method="",
    update_type="",
    source_value="",
    customer_value="",
    old_value="",
    new_value="",
    status="",
    details="",
):

    audit_rows.append(
        {
            "Timestamp":
                datetime.now().isoformat(
                    timespec="seconds"
                ),

            "Source Sheet":
                source_sheet,

            "Row Key":
                clean_value(
                    row_key
                ),

            "English UUID":
                clean_value(
                    english_uuid
                ),

            "French UUID":
                clean_value(
                    french_uuid
                ),

            "English Index":
                (
                    ""
                    if english_index is None
                    else english_index
                ),

            "French Index":
                (
                    ""
                    if french_index is None
                    else french_index
                ),

            "Local Community Name":
                clean_value(
                    community
                ),

            "title":
                clean_value(
                    title
                ),

            "Match Method":
                match_method,

            "Update Type":
                update_type,

            "Source Value":
                clean_value(
                    source_value
                ),

            "Customer Value":
                clean_value(
                    customer_value
                ),

            "Old French Value":
                clean_value(
                    old_value
                ),

            "New French Value":
                clean_value(
                    new_value
                ),

            "Status":
                status,

            "Details":
                details,
        }
    )


# ============================================================
# LOAD INPUTS
# ============================================================

def load_inputs():

    required_files = [
        ENGLISH_TRAIN_FILE,
        FRENCH_TRAIN_FILE,
        TRANSLATION_FILE,
    ]

    missing = [
        path
        for path
        in required_files
        if not path.exists()
    ]

    if missing:

        raise FileNotFoundError(
            "\nRequired Stage 2 files missing:\n\n"
            +
            "\n".join(
                str(path)
                for path
                in missing
            )
        )

    english_df = pd.read_csv(
        ENGLISH_TRAIN_FILE,
        dtype=object,
    )

    french_df = pd.read_csv(
        FRENCH_TRAIN_FILE,
        dtype=object,
    )

    workbook = pd.ExcelFile(
        TRANSLATION_FILE
    )

    required_sheets = {
        "Comparison",
        "French URLs",
        "French Content Translation",
    }

    missing_sheets = (
        required_sheets
        -
        set(
            workbook.sheet_names
        )
    )

    if missing_sheets:

        raise ValueError(
            "\nTranslation workbook missing sheets:\n"
            +
            "\n".join(
                sorted(
                    missing_sheets
                )
            )
        )

    comparison_df = pd.read_excel(
        TRANSLATION_FILE,
        sheet_name="Comparison",
        dtype=object,
    )

    french_urls_df = pd.read_excel(
        TRANSLATION_FILE,
        sheet_name="French URLs",
        dtype=object,
    )

    french_content_df = pd.read_excel(
        TRANSLATION_FILE,
        sheet_name="French Content Translation",
        dtype=object,
    )

    return (
        english_df,
        french_df,
        comparison_df,
        french_urls_df,
        french_content_df,
    )


# ============================================================
# INPUT COUNTER
# ============================================================

def count_inputs(
    dataframe,
    column,
):

    if column not in dataframe.columns:
        return 0

    return sum(
        bool(
            clean_value(
                value
            )
        )
        for value
        in dataframe[
            column
        ]
    )


# ============================================================
# RESOLVE TRANSLATION ROW
# ============================================================

def resolve_translation_row(
    row,
    comparison_lookup,
    english_df,
    french_df,
    english_uuid_to_french_indexes,
):

    row_key = clean_value(
        row.get(
            COL_ROW_KEY
        )
    )

    comparison = (
        comparison_lookup.get(
            row_key
        )
    )

    if not comparison:

        return (
            None,
            {
                "status":
                    "COMPARISON_ROW_NOT_FOUND",

                "method":
                    "",
            },
        )

    resolution = (
        resolve_french_record(
            english_df,
            french_df,
            comparison,
            english_uuid_to_french_indexes,
        )
    )

    return (
        comparison,
        resolution,
    )


# ============================================================
# CONTENT TEXT TRANSLATIONS
# ============================================================

def apply_content_text_translations(
    english_df,
    french_df,
    french_content_df,
    comparison_lookup,
    english_uuid_to_french_indexes,
    audit_rows,
):

    if (
        COL_FRENCH_CONTENT
        not in french_content_df.columns
    ):

        return 0

    updated_count = 0

    # --------------------------------------------------------
    # Dedupe only AFTER exact French target resolution.
    # --------------------------------------------------------

    seen = set()

    for (
        _,
        row,
    ) in french_content_df.iterrows():

        french_translation = clean_value(
            row.get(
                COL_FRENCH_CONTENT
            )
        )

        if not french_translation:
            continue

        row_key = clean_value(
            row.get(
                COL_ROW_KEY
            )
        )

        (
            comparison,
            resolution,
        ) = resolve_translation_row(
            row,
            comparison_lookup,
            english_df,
            french_df,
            english_uuid_to_french_indexes,
        )

        if not comparison:

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                community=row.get(
                    COL_COMMUNITY
                ),
                title=row.get(
                    COL_TITLE
                ),
                update_type=(
                    "Content Text"
                ),
                customer_value=(
                    french_translation
                ),
                status=(
                    "COMPARISON_ROW_NOT_FOUND"
                ),
            )

            continue

        changed_text = clean_value(
            comparison.get(
                COL_CONTENT_CHANGES,
                row.get(
                    COL_CONTENT_CHANGES
                ),
            )
        )

        context_before = clean_value(
            comparison.get(
                COL_CONTEXT_BEFORE
            )
        )

        context_after = clean_value(
            comparison.get(
                COL_CONTEXT_AFTER
            )
        )

        try:

            block_count = int(
                float(
                    comparison.get(
                        COL_BLOCK_COUNT,
                        0,
                    )
                    or
                    0
                )
            )

        except (
            ValueError,
            TypeError,
        ):

            block_count = 0

        if (
            resolution.get(
                "status"
            )
            !=
            "FOUND"
        ):

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    resolution.get(
                        "english_uuid",
                        "",
                    )
                ),
                english_index=(
                    resolution.get(
                        "english_index"
                    )
                ),
                community=comparison.get(
                    COL_COMMUNITY
                ),
                title=comparison.get(
                    COL_TITLE
                ),
                match_method=(
                    resolution.get(
                        "method",
                        "",
                    )
                ),
                update_type=(
                    "Content Text"
                ),
                source_value=(
                    changed_text
                ),
                customer_value=(
                    french_translation
                ),
                status=(
                    resolution.get(
                        "status"
                    )
                ),
            )

            continue

        english_index = (
            resolution[
                "english_index"
            ]
        )

        french_index = (
            resolution[
                "french_index"
            ]
        )

        english_uuid = (
            resolution[
                "english_uuid"
            ]
        )

        french_uuid = clean_value(
            french_df.at[
                french_index,
                COL_UUID
            ]
        )

        # ====================================================
        # DEDUPE AFTER EXACT TARGET IS KNOWN
        # ====================================================

        dedupe_key = (
            french_uuid,
            normalize_text(
                changed_text
            ),
            normalize_text(
                french_translation
            ),
        )

        if dedupe_key in seen:

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                community=(
                    comparison.get(
                        COL_COMMUNITY
                    )
                ),
                title=(
                    comparison.get(
                        COL_TITLE
                    )
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Content Text"
                ),
                source_value=(
                    changed_text
                ),
                customer_value=(
                    french_translation
                ),
                status=(
                    "SKIPPED_DUPLICATE"
                ),
            )

            continue

        seen.add(
            dedupe_key
        )

        english_contents = parse_contents(
            english_df.at[
                english_index,
                COL_CONTENTS
            ]
        )

        french_contents = parse_contents(
            french_df.at[
                french_index,
                COL_CONTENTS
            ]
        )

        if not english_contents:

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                community=(
                    comparison.get(
                        COL_COMMUNITY
                    )
                ),
                title=(
                    comparison.get(
                        COL_TITLE
                    )
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Content Text"
                ),
                source_value=(
                    changed_text
                ),
                customer_value=(
                    french_translation
                ),
                status=(
                    "ENGLISH_CONTENT_EMPTY"
                ),
            )

            continue

        english_heading = clean_value(
            comparison.get(
                COL_TRAIN_HEADING
            )
        )

        english_url_name = clean_value(
            comparison.get(
                COL_TRAIN_URL_NAME
            )
        )

        english_url = clean_value(
            comparison.get(
                COL_TRAIN_URL
            )
        )

        (
            english_item_index,
            english_item_method,
        ) = find_content_item_index(
            english_contents,
            heading=(
                english_heading
            ),
            url_name=(
                english_url_name
            ),
            url=(
                english_url
            ),
        )

        if english_item_index is None:

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                community=(
                    comparison.get(
                        COL_COMMUNITY
                    )
                ),
                title=(
                    comparison.get(
                        COL_TITLE
                    )
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Content Text"
                ),
                source_value=(
                    changed_text
                ),
                customer_value=(
                    french_translation
                ),
                status=(
                    "ENGLISH_CONTENT_ITEM_NOT_FOUND"
                ),
                details=(
                    english_item_method
                ),
            )

            continue

        english_item = (
            english_contents[
                english_item_index
            ]
        )

        english_content = str(
            english_item.get(
                "content"
            )
            or
            ""
        )

        changed_blocks = split_blocks(
            changed_text,
            block_count,
        )

        french_blocks = split_blocks(
            french_translation,
            block_count,
        )

        before_blocks = split_blocks(
            context_before,
            block_count,
        )

        after_blocks = split_blocks(
            context_after,
            block_count,
        )

        if block_count <= 1:

            changed_blocks = [
                changed_text
            ]

            french_blocks = [
                french_translation
            ]

            before_blocks = [
                context_before
            ]

            after_blocks = [
                context_after
            ]

        if (
            block_count > 1
            and
            (
                len(
                    changed_blocks
                )
                !=
                block_count
                or
                len(
                    french_blocks
                )
                !=
                block_count
            )
        ):

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                community=(
                    comparison.get(
                        COL_COMMUNITY
                    )
                ),
                title=(
                    comparison.get(
                        COL_TITLE
                    )
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Content Text"
                ),
                source_value=(
                    changed_text
                ),
                customer_value=(
                    french_translation
                ),
                status=(
                    "REVIEW_REQUIRED"
                ),
                details=(
                    "Multiple blocks do not align. "
                    f"Expected={block_count}; "
                    f"English={len(changed_blocks)}; "
                    f"French={len(french_blocks)}"
                ),
            )

            continue

        # ====================================================
        # EXISTING FRENCH CONTENT
        # ====================================================

        if french_contents:

            # ------------------------------------------------
            # Prefer same content item index.
            # ------------------------------------------------

            if (
                english_item_index
                <
                len(
                    french_contents
                )
            ):

                french_item_index = (
                    english_item_index
                )

            else:

                french_item_index = (
                    None
                )

            if french_item_index is None:

                add_audit(
                    audit_rows,
                    source_sheet=(
                        "French Content Translation"
                    ),
                    row_key=row_key,
                    english_uuid=(
                        english_uuid
                    ),
                    french_uuid=(
                        french_uuid
                    ),
                    english_index=(
                        english_index
                    ),
                    french_index=(
                        french_index
                    ),
                    community=(
                        comparison.get(
                            COL_COMMUNITY
                        )
                    ),
                    title=(
                        comparison.get(
                            COL_TITLE
                        )
                    ),
                    match_method=(
                        resolution[
                            "method"
                        ]
                    ),
                    update_type=(
                        "Content Text"
                    ),
                    source_value=(
                        changed_text
                    ),
                    customer_value=(
                        french_translation
                    ),
                    status=(
                        "FRENCH_CONTENT_ITEM_NOT_FOUND"
                    ),
                )

                continue

            french_item = (
                french_contents[
                    french_item_index
                ]
            )

            if not isinstance(
                french_item,
                dict,
            ):

                add_audit(
                    audit_rows,
                    source_sheet=(
                        "French Content Translation"
                    ),
                    row_key=row_key,
                    english_uuid=(
                        english_uuid
                    ),
                    french_uuid=(
                        french_uuid
                    ),
                    english_index=(
                        english_index
                    ),
                    french_index=(
                        french_index
                    ),
                    community=(
                        comparison.get(
                            COL_COMMUNITY
                        )
                    ),
                    title=(
                        comparison.get(
                            COL_TITLE
                        )
                    ),
                    match_method=(
                        resolution[
                            "method"
                        ]
                    ),
                    update_type=(
                        "Content Text"
                    ),
                    source_value=(
                        changed_text
                    ),
                    customer_value=(
                        french_translation
                    ),
                    status=(
                        "INVALID_FRENCH_CONTENT_STRUCTURE"
                    ),
                )

                continue

            old_serialized = (
                serialize_contents(
                    french_contents
                )
            )

            working_content = str(
                french_item.get(
                    "content"
                )
                or
                ""
            )

            success = True

            details = []

            for block_index in range(
                len(
                    changed_blocks
                )
            ):

                (
                    working_content,
                    changed,
                    detail,
                ) = update_existing_french_content(
                    english_content,
                    working_content,
                    changed_blocks[
                        block_index
                    ],
                    french_blocks[
                        block_index
                    ],
                    (
                        before_blocks[
                            block_index
                        ]
                        if (
                            block_index
                            <
                            len(
                                before_blocks
                            )
                        )
                        else ""
                    ),
                    (
                        after_blocks[
                            block_index
                        ]
                        if (
                            block_index
                            <
                            len(
                                after_blocks
                            )
                        )
                        else ""
                    ),
                )

                details.append(
                    f"Block "
                    f"{block_index + 1}: "
                    f"{detail}"
                )

                if not changed:

                    success = False

                    break

            if success:

                french_item[
                    "content"
                ] = (
                    working_content
                )

                new_serialized = (
                    serialize_contents(
                        french_contents
                    )
                )

                french_df.at[
                    french_index,
                    COL_CONTENTS
                ] = (
                    new_serialized
                )

                updated_count += 1

                status = (
                    "UPDATED"
                )

            else:

                new_serialized = (
                    old_serialized
                )

                status = (
                    "REVIEW_REQUIRED"
                )

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                community=(
                    comparison.get(
                        COL_COMMUNITY
                    )
                ),
                title=(
                    comparison.get(
                        COL_TITLE
                    )
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Content Text"
                ),
                source_value=(
                    changed_text
                ),
                customer_value=(
                    french_translation
                ),
                old_value=(
                    old_serialized
                ),
                new_value=(
                    new_serialized
                ),
                status=(
                    status
                ),
                details=(
                    "; ".join(
                        details
                    )
                ),
            )

            continue

        # ====================================================
        # FRENCH CONTENTS IS []
        #
        # BUILD USING ENGLISH STRUCTURE
        # ====================================================

        new_french_contents = (
            deepcopy(
                english_contents
            )
        )

        template_item = (
            new_french_contents[
                english_item_index
            ]
        )

        working_content = (
            english_content
        )

        success = True

        details = []

        for block_index in range(
            len(
                changed_blocks
            )
        ):

            (
                working_content,
                changed,
                detail,
            ) = build_content_from_english_template(
                working_content,
                changed_blocks[
                    block_index
                ],
                french_blocks[
                    block_index
                ],
                (
                    before_blocks[
                        block_index
                    ]
                    if (
                        block_index
                        <
                        len(
                            before_blocks
                        )
                    )
                    else ""
                ),
                (
                    after_blocks[
                        block_index
                    ]
                    if (
                        block_index
                        <
                        len(
                            after_blocks
                        )
                    )
                    else ""
                ),
            )

            details.append(
                f"Block "
                f"{block_index + 1}: "
                f"{detail}"
            )

            if not changed:

                success = False

                break

        if success:

            template_item[
                "content"
            ] = (
                working_content
            )

            new_serialized = (
                serialize_contents(
                    new_french_contents
                )
            )

            french_df.at[
                french_index,
                COL_CONTENTS
            ] = (
                new_serialized
            )

            updated_count += 1

            status = (
                "UPDATED_FROM_EMPTY_FRENCH_CONTENT"
            )

        else:

            new_serialized = (
                "[]"
            )

            status = (
                "REVIEW_REQUIRED"
            )

        add_audit(
            audit_rows,
            source_sheet=(
                "French Content Translation"
            ),
            row_key=row_key,
            english_uuid=(
                english_uuid
            ),
            french_uuid=(
                french_uuid
            ),
            english_index=(
                english_index
            ),
            french_index=(
                french_index
            ),
            community=(
                comparison.get(
                    COL_COMMUNITY
                )
            ),
            title=(
                comparison.get(
                    COL_TITLE
                )
            ),
            match_method=(
                resolution[
                    "method"
                ]
            ),
            update_type=(
                "Content Text"
            ),
            source_value=(
                changed_text
            ),
            customer_value=(
                french_translation
            ),
            old_value=(
                "[]"
            ),
            new_value=(
                new_serialized
            ),
            status=(
                status
            ),
            details=(
                "; ".join(
                    details
                )
                +
                "; French contents was empty. "
                "English contents used as structural template."
            ),
        )

    return updated_count


# ============================================================
# HEADING TRANSLATIONS
# ============================================================

def apply_heading_translations(
    english_df,
    french_df,
    french_content_df,
    comparison_lookup,
    english_uuid_to_french_indexes,
    audit_rows,
):

    if (
        COL_FRENCH_HEADING
        not in french_content_df.columns
    ):

        return 0

    updated_count = 0
    seen = set()

    for (
        _,
        row,
    ) in french_content_df.iterrows():

        new_heading = clean_value(
            row.get(
                COL_FRENCH_HEADING
            )
        )

        if not new_heading:
            continue

        row_key = clean_value(
            row.get(
                COL_ROW_KEY
            )
        )

        (
            comparison,
            resolution,
        ) = resolve_translation_row(
            row,
            comparison_lookup,
            english_df,
            french_df,
            english_uuid_to_french_indexes,
        )

        if (
            not comparison
            or
            resolution.get(
                "status"
            )
            !=
            "FOUND"
        ):

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    resolution.get(
                        "english_uuid",
                        "",
                    )
                ),
                english_index=(
                    resolution.get(
                        "english_index"
                    )
                ),
                community=(
                    row.get(
                        COL_COMMUNITY
                    )
                ),
                title=(
                    row.get(
                        COL_TITLE
                    )
                ),
                match_method=(
                    resolution.get(
                        "method",
                        "",
                    )
                ),
                update_type=(
                    "Heading"
                ),
                customer_value=(
                    new_heading
                ),
                status=(
                    resolution.get(
                        "status",
                        "COMPARISON_ROW_NOT_FOUND",
                    )
                ),
            )

            continue

        english_index = (
            resolution[
                "english_index"
            ]
        )

        french_index = (
            resolution[
                "french_index"
            ]
        )

        english_uuid = (
            resolution[
                "english_uuid"
            ]
        )

        french_uuid = clean_value(
            french_df.at[
                french_index,
                COL_UUID
            ]
        )

        old_english_heading = clean_value(
            comparison.get(
                COL_TRAIN_HEADING
            )
        )

        dedupe_key = (
            french_uuid,
            normalized_key(
                old_english_heading
            ),
            normalize_text(
                new_heading
            ),
        )

        if dedupe_key in seen:

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Heading"
                ),
                source_value=(
                    old_english_heading
                ),
                customer_value=(
                    new_heading
                ),
                status=(
                    "SKIPPED_DUPLICATE"
                ),
            )

            continue

        seen.add(
            dedupe_key
        )

        english_contents = parse_contents(
            english_df.at[
                english_index,
                COL_CONTENTS
            ]
        )

        french_contents = parse_contents(
            french_df.at[
                french_index,
                COL_CONTENTS
            ]
        )

        if not french_contents:

            # Content text phase may not have constructed it.
            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Heading"
                ),
                source_value=(
                    old_english_heading
                ),
                customer_value=(
                    new_heading
                ),
                status=(
                    "FRENCH_CONTENT_EMPTY"
                ),
            )

            continue

        (
            english_item_index,
            item_method,
        ) = find_content_item_index(
            english_contents,
            heading=(
                old_english_heading
            ),
            url_name=(
                comparison.get(
                    COL_TRAIN_URL_NAME
                )
            ),
            url=(
                comparison.get(
                    COL_TRAIN_URL
                )
            ),
        )

        if (
            english_item_index is None
            or
            english_item_index
            >=
            len(
                french_contents
            )
        ):

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Heading"
                ),
                source_value=(
                    old_english_heading
                ),
                customer_value=(
                    new_heading
                ),
                status=(
                    "CONTENT_ITEM_NOT_FOUND"
                ),
                details=(
                    item_method
                ),
            )

            continue

        item = (
            french_contents[
                english_item_index
            ]
        )

        old_value = (
            item.get(
                "heading"
            )
        )

        if values_equal(
            old_value,
            new_heading,
        ):

            status = (
                "NO_CHANGE"
            )

        else:

            item[
                "heading"
            ] = (
                new_heading
            )

            french_df.at[
                french_index,
                COL_CONTENTS
            ] = (
                serialize_contents(
                    french_contents
                )
            )

            updated_count += 1

            status = (
                "UPDATED"
            )

        add_audit(
            audit_rows,
            source_sheet=(
                "French Content Translation"
            ),
            row_key=row_key,
            english_uuid=(
                english_uuid
            ),
            french_uuid=(
                french_uuid
            ),
            english_index=(
                english_index
            ),
            french_index=(
                french_index
            ),
            match_method=(
                resolution[
                    "method"
                ]
            ),
            update_type=(
                "Heading"
            ),
            source_value=(
                old_english_heading
            ),
            customer_value=(
                new_heading
            ),
            old_value=(
                old_value
            ),
            new_value=(
                new_heading
            ),
            status=(
                status
            ),
        )

    return updated_count


# ============================================================
# URL NAME TRANSLATIONS
# ============================================================

def apply_url_name_translations(
    english_df,
    french_df,
    french_content_df,
    comparison_lookup,
    english_uuid_to_french_indexes,
    audit_rows,
):

    if (
        COL_FRENCH_URL_NAME
        not in french_content_df.columns
    ):

        return 0

    updated_count = 0
    seen = set()

    for (
        _,
        row,
    ) in french_content_df.iterrows():

        new_label = clean_value(
            row.get(
                COL_FRENCH_URL_NAME
            )
        )

        if not new_label:
            continue

        row_key = clean_value(
            row.get(
                COL_ROW_KEY
            )
        )

        (
            comparison,
            resolution,
        ) = resolve_translation_row(
            row,
            comparison_lookup,
            english_df,
            french_df,
            english_uuid_to_french_indexes,
        )

        if (
            not comparison
            or
            resolution.get(
                "status"
            )
            !=
            "FOUND"
        ):

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    resolution.get(
                        "english_uuid",
                        "",
                    )
                ),
                english_index=(
                    resolution.get(
                        "english_index"
                    )
                ),
                community=(
                    row.get(
                        COL_COMMUNITY
                    )
                ),
                title=(
                    row.get(
                        COL_TITLE
                    )
                ),
                match_method=(
                    resolution.get(
                        "method",
                        "",
                    )
                ),
                update_type=(
                    "URL Name"
                ),
                customer_value=(
                    new_label
                ),
                status=(
                    resolution.get(
                        "status",
                        "COMPARISON_ROW_NOT_FOUND",
                    )
                ),
            )

            continue

        english_index = (
            resolution[
                "english_index"
            ]
        )

        french_index = (
            resolution[
                "french_index"
            ]
        )

        english_uuid = (
            resolution[
                "english_uuid"
            ]
        )

        french_uuid = clean_value(
            french_df.at[
                french_index,
                COL_UUID
            ]
        )

        old_label = clean_value(
            comparison.get(
                COL_TRAIN_URL_NAME
            )
        )

        old_url = clean_value(
            comparison.get(
                COL_TRAIN_URL
            )
        )

        dedupe_key = (
            french_uuid,
            normalized_key(
                old_label
            ),
            old_url,
            normalize_text(
                new_label
            ),
        )

        if dedupe_key in seen:

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "URL Name"
                ),
                source_value=(
                    old_label
                ),
                customer_value=(
                    new_label
                ),
                status=(
                    "SKIPPED_DUPLICATE"
                ),
            )

            continue

        seen.add(
            dedupe_key
        )

        contents = parse_contents(
            french_df.at[
                french_index,
                COL_CONTENTS
            ]
        )

        if not contents:

            add_audit(
                audit_rows,
                source_sheet=(
                    "French Content Translation"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "URL Name"
                ),
                source_value=(
                    old_label
                ),
                customer_value=(
                    new_label
                ),
                status=(
                    "FRENCH_CONTENT_EMPTY"
                ),
            )

            continue

        old_serialized = (
            serialize_contents(
                contents
            )
        )

        success = False
        detail = (
            "LABEL_NOT_FOUND"
        )

        for item in contents:

            if not isinstance(
                item,
                dict,
            ):
                continue

            content = str(
                item.get(
                    "content"
                )
                or
                ""
            )

            (
                new_content,
                changed,
                detail,
            ) = replace_markdown_label(
                content,
                old_label,
                new_label,
                old_url,
            )

            if changed:

                item[
                    "content"
                ] = (
                    new_content
                )

                success = True

                break

        if success:

            new_serialized = (
                serialize_contents(
                    contents
                )
            )

            french_df.at[
                french_index,
                COL_CONTENTS
            ] = (
                new_serialized
            )

            updated_count += 1

            status = (
                "UPDATED"
            )

        else:

            new_serialized = (
                old_serialized
            )

            status = (
                "REVIEW_REQUIRED"
            )

        add_audit(
            audit_rows,
            source_sheet=(
                "French Content Translation"
            ),
            row_key=row_key,
            english_uuid=(
                english_uuid
            ),
            french_uuid=(
                french_uuid
            ),
            english_index=(
                english_index
            ),
            french_index=(
                french_index
            ),
            match_method=(
                resolution[
                    "method"
                ]
            ),
            update_type=(
                "URL Name"
            ),
            source_value=(
                old_label
            ),
            customer_value=(
                new_label
            ),
            old_value=(
                old_serialized
            ),
            new_value=(
                new_serialized
            ),
            status=(
                status
            ),
            details=(
                detail
            ),
        )

    return updated_count


# ============================================================
# CONTENT URL UPDATES
# ============================================================

def apply_content_url_updates(
    english_df,
    french_df,
    french_urls_df,
    comparison_lookup,
    english_uuid_to_french_indexes,
    audit_rows,
):

    if (
        COL_FRENCH_CONTENT_URL
        not in french_urls_df.columns
    ):

        return 0

    updated_count = 0
    seen = set()

    for (
        _,
        row,
    ) in french_urls_df.iterrows():

        new_url = clean_value(
            row.get(
                COL_FRENCH_CONTENT_URL
            )
        )

        if not new_url:
            continue

        row_key = clean_value(
            row.get(
                COL_ROW_KEY
            )
        )

        (
            comparison,
            resolution,
        ) = resolve_translation_row(
            row,
            comparison_lookup,
            english_df,
            french_df,
            english_uuid_to_french_indexes,
        )

        if (
            not comparison
            or
            resolution.get(
                "status"
            )
            !=
            "FOUND"
        ):

            add_audit(
                audit_rows,
                source_sheet=(
                    "French URLs"
                ),
                row_key=row_key,
                english_uuid=(
                    resolution.get(
                        "english_uuid",
                        "",
                    )
                ),
                english_index=(
                    resolution.get(
                        "english_index"
                    )
                ),
                community=(
                    row.get(
                        COL_COMMUNITY
                    )
                ),
                title=(
                    row.get(
                        COL_TITLE
                    )
                ),
                match_method=(
                    resolution.get(
                        "method",
                        "",
                    )
                ),
                update_type=(
                    "Content URL"
                ),
                customer_value=(
                    new_url
                ),
                status=(
                    resolution.get(
                        "status",
                        "COMPARISON_ROW_NOT_FOUND",
                    )
                ),
            )

            continue

        english_index = (
            resolution[
                "english_index"
            ]
        )

        french_index = (
            resolution[
                "french_index"
            ]
        )

        english_uuid = (
            resolution[
                "english_uuid"
            ]
        )

        french_uuid = clean_value(
            french_df.at[
                french_index,
                COL_UUID
            ]
        )

        old_url = clean_value(
            comparison.get(
                COL_TRAIN_URL,
                row.get(
                    COL_TRAIN_URL
                ),
            )
        )

        dedupe_key = (
            french_uuid,
            old_url,
            new_url,
        )

        if dedupe_key in seen:

            add_audit(
                audit_rows,
                source_sheet=(
                    "French URLs"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Content URL"
                ),
                source_value=(
                    old_url
                ),
                customer_value=(
                    new_url
                ),
                status=(
                    "SKIPPED_DUPLICATE"
                ),
            )

            continue

        seen.add(
            dedupe_key
        )

        contents = parse_contents(
            french_df.at[
                french_index,
                COL_CONTENTS
            ]
        )

        if not contents:

            add_audit(
                audit_rows,
                source_sheet=(
                    "French URLs"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Content URL"
                ),
                source_value=(
                    old_url
                ),
                customer_value=(
                    new_url
                ),
                status=(
                    "FRENCH_CONTENT_EMPTY"
                ),
            )

            continue

        old_serialized = (
            serialize_contents(
                contents
            )
        )

        success = False
        detail = (
            "URL_NOT_FOUND"
        )

        for item in contents:

            if not isinstance(
                item,
                dict,
            ):
                continue

            content = str(
                item.get(
                    "content"
                )
                or
                ""
            )

            (
                new_content,
                changed,
                detail,
            ) = replace_markdown_url(
                content,
                old_url,
                new_url,
            )

            if changed:

                item[
                    "content"
                ] = (
                    new_content
                )

                success = True

                break

        if success:

            new_serialized = (
                serialize_contents(
                    contents
                )
            )

            french_df.at[
                french_index,
                COL_CONTENTS
            ] = (
                new_serialized
            )

            updated_count += 1

            status = (
                "UPDATED"
            )

        else:

            new_serialized = (
                old_serialized
            )

            status = (
                "REVIEW_REQUIRED"
            )

        add_audit(
            audit_rows,
            source_sheet=(
                "French URLs"
            ),
            row_key=row_key,
            english_uuid=(
                english_uuid
            ),
            french_uuid=(
                french_uuid
            ),
            english_index=(
                english_index
            ),
            french_index=(
                french_index
            ),
            match_method=(
                resolution[
                    "method"
                ]
            ),
            update_type=(
                "Content URL"
            ),
            source_value=(
                old_url
            ),
            customer_value=(
                new_url
            ),
            old_value=(
                old_serialized
            ),
            new_value=(
                new_serialized
            ),
            status=(
                status
            ),
            details=(
                detail
            ),
        )

    return updated_count


# ============================================================
# BUTTON LINK UPDATES
# ============================================================

def apply_button_link_updates(
    english_df,
    french_df,
    french_urls_df,
    comparison_lookup,
    english_uuid_to_french_indexes,
    audit_rows,
):

    if (
        COL_FRENCH_BUTTON_LINK
        not in french_urls_df.columns
    ):

        return 0

    updated_count = 0
    seen = set()

    for (
        _,
        row,
    ) in french_urls_df.iterrows():

        new_value = clean_value(
            row.get(
                COL_FRENCH_BUTTON_LINK
            )
        )

        if not new_value:
            continue

        row_key = clean_value(
            row.get(
                COL_ROW_KEY
            )
        )

        (
            comparison,
            resolution,
        ) = resolve_translation_row(
            row,
            comparison_lookup,
            english_df,
            french_df,
            english_uuid_to_french_indexes,
        )

        if (
            not comparison
            or
            resolution.get(
                "status"
            )
            !=
            "FOUND"
        ):

            add_audit(
                audit_rows,
                source_sheet=(
                    "French URLs"
                ),
                row_key=row_key,
                english_uuid=(
                    resolution.get(
                        "english_uuid",
                        "",
                    )
                ),
                english_index=(
                    resolution.get(
                        "english_index"
                    )
                ),
                community=(
                    row.get(
                        COL_COMMUNITY
                    )
                ),
                title=(
                    row.get(
                        COL_TITLE
                    )
                ),
                match_method=(
                    resolution.get(
                        "method",
                        "",
                    )
                ),
                update_type=(
                    "Button Link"
                ),
                customer_value=(
                    new_value
                ),
                status=(
                    resolution.get(
                        "status",
                        "COMPARISON_ROW_NOT_FOUND",
                    )
                ),
            )

            continue

        english_index = (
            resolution[
                "english_index"
            ]
        )

        french_index = (
            resolution[
                "french_index"
            ]
        )

        english_uuid = (
            resolution[
                "english_uuid"
            ]
        )

        french_uuid = clean_value(
            french_df.at[
                french_index,
                COL_UUID
            ]
        )

        dedupe_key = (
            french_uuid,
            new_value,
        )

        if dedupe_key in seen:

            add_audit(
                audit_rows,
                source_sheet=(
                    "French URLs"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Button Link"
                ),
                customer_value=(
                    new_value
                ),
                status=(
                    "SKIPPED_DUPLICATE"
                ),
            )

            continue

        seen.add(
            dedupe_key
        )

        if (
            COL_BUTTON_LINK
            not in french_df.columns
        ):

            add_audit(
                audit_rows,
                source_sheet=(
                    "French URLs"
                ),
                row_key=row_key,
                english_uuid=(
                    english_uuid
                ),
                french_uuid=(
                    french_uuid
                ),
                english_index=(
                    english_index
                ),
                french_index=(
                    french_index
                ),
                match_method=(
                    resolution[
                        "method"
                    ]
                ),
                update_type=(
                    "Button Link"
                ),
                customer_value=(
                    new_value
                ),
                status=(
                    "BUTTON_LINK_COLUMN_NOT_FOUND"
                ),
            )

            continue

        old_value = (
            french_df.at[
                french_index,
                COL_BUTTON_LINK
            ]
        )

        if values_equal(
            old_value,
            new_value,
        ):

            status = (
                "NO_CHANGE"
            )

        else:

            french_df.at[
                french_index,
                COL_BUTTON_LINK
            ] = (
                new_value
            )

            updated_count += 1

            status = (
                "UPDATED"
            )

        add_audit(
            audit_rows,
            source_sheet=(
                "French URLs"
            ),
            row_key=row_key,
            english_uuid=(
                english_uuid
            ),
            french_uuid=(
                french_uuid
            ),
            english_index=(
                english_index
            ),
            french_index=(
                french_index
            ),
            match_method=(
                resolution[
                    "method"
                ]
            ),
            update_type=(
                "Button Link"
            ),
            customer_value=(
                new_value
            ),
            old_value=(
                old_value
            ),
            new_value=(
                new_value
            ),
            status=(
                status
            ),
        )

    return updated_count


# ============================================================
# VALIDATE FINAL OUTPUT
# ============================================================

def validate_final_output(
    original_df,
    final_df,
):

    errors = []

    if (
        len(
            original_df
        )
        !=
        len(
            final_df
        )
    ):

        errors.append(
            "French row count changed."
        )

    if (
        list(
            original_df.columns
        )
        !=
        list(
            final_df.columns
        )
    ):

        errors.append(
            "French columns changed."
        )

    if (
        COL_LOCALE
        in final_df.columns
    ):

        invalid = (
            final_df[
                COL_LOCALE
            ]
            .fillna("")
            .astype(str)
            .ne(
                "fr-CA"
            )
        )

        if invalid.any():

            errors.append(
                "Output contains non-fr-CA records."
            )

    if errors:

        raise RuntimeError(
            "\nFinal validation failed:\n"
            +
            "\n".join(
                f"- {error}"
                for error
                in errors
            )
        )


# ============================================================
# WRITE FINAL EXCEL
# ============================================================

def write_final_excel(
    dataframe,
):

    with pd.ExcelWriter(
        FINAL_XLSX,
        engine="openpyxl",
    ) as writer:

        dataframe.to_excel(
            writer,
            sheet_name=(
                "French Updated"
            ),
            index=False,
        )

        worksheet = (
            writer.book[
                "French Updated"
            ]
        )

        worksheet.freeze_panes = (
            "A2"
        )

        for cell in (
            worksheet[
                1
            ]
        ):

            cell.font = Font(
                bold=True
            )

            cell.alignment = Alignment(
                wrap_text=True,
                vertical="top",
            )

        for column_cells in (
            worksheet.iter_cols()
        ):

            width = max(
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

            worksheet.column_dimensions[
                column_cells[
                    0
                ].column_letter
            ].width = min(
                width + 2,
                30,
            )

        if (
            worksheet.max_row > 0
            and
            worksheet.max_column > 0
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

            table = Table(
                displayName=(
                    "FrenchUpdatedTable"
                ),
                ref=(
                    f"A1:{last_cell}"
                ),
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


# ============================================================
# WRITE AUDIT
# ============================================================

def write_audit(
    audit_rows,
):

    audit_df = pd.DataFrame(
        audit_rows,
        columns=(
            AUDIT_COLUMNS
        ),
    )

    if audit_df.empty:

        summary_df = pd.DataFrame(
            [
                {
                    "Status":
                        "NO CUSTOMER INPUT",

                    "Count":
                        0,
                }
            ]
        )

    else:

        summary_df = (
            audit_df[
                "Status"
            ]
            .value_counts(
                dropna=False
            )
            .rename_axis(
                "Status"
            )
            .reset_index(
                name="Count"
            )
        )

    with pd.ExcelWriter(
        AUDIT_XLSX,
        engine="openpyxl",
    ) as writer:

        audit_df.to_excel(
            writer,
            sheet_name="Audit",
            index=False,
        )

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
        )

        for worksheet in (
            writer.book.worksheets
        ):

            worksheet.freeze_panes = (
                "A2"
            )

            for cell in (
                worksheet[
                    1
                ]
            ):

                cell.font = Font(
                    bold=True
                )

                cell.alignment = Alignment(
                    wrap_text=True,
                )

            for column_cells in (
                worksheet.iter_cols()
            ):

                width = max(
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

                worksheet.column_dimensions[
                    column_cells[
                        0
                    ].column_letter
                ].width = min(
                    width + 2,
                    40,
                )


# ============================================================
# MAIN
# ============================================================

def main():

    banner(
        "CMS FRENCH TRANSLATION SYNC - STAGE 2"
    )

    print()
    print(
        "English Train:"
    )
    print(
        ENGLISH_TRAIN_FILE
    )

    print()
    print(
        "French Train:"
    )
    print(
        FRENCH_TRAIN_FILE
    )

    print()
    print(
        "Translation workbook:"
    )
    print(
        TRANSLATION_FILE
    )

    # ========================================================
    # LOAD
    # ========================================================

    (
        english_df,
        french_df,
        comparison_df,
        french_urls_df,
        french_content_df,
    ) = load_inputs()

    original_french_df = (
        french_df.copy(
            deep=True
        )
    )

    comparison_lookup = (
        build_comparison_lookup(
            comparison_df
        )
    )

    (
        english_uuid_to_index,
        french_uuid_to_index,
        english_uuid_to_french_indexes,
    ) = build_uuid_maps(
        english_df,
        french_df,
    )

    # ========================================================
    # MAPPING DIAGNOSTIC
    # ========================================================

    mapped_english_uuids = len(
        english_uuid_to_french_indexes
    )

    banner(
        "STAGE 2 MAPPING"
    )

    print(
        f"English Train records: "
        f"{len(english_df)}"
    )

    print(
        f"French Train records: "
        f"{len(french_df)}"
    )

    print(
        f"English UUIDs: "
        f"{len(english_uuid_to_index)}"
    )

    print(
        f"French UUIDs: "
        f"{len(french_uuid_to_index)}"
    )

    print(
        f"English UUIDs referenced by "
        f"French localizations: "
        f"{mapped_english_uuids}"
    )

    print(
        f"Comparison rows: "
        f"{len(comparison_df)}"
    )

    print(
        f"French Content Translation rows: "
        f"{len(french_content_df)}"
    )

    print(
        f"French URLs rows: "
        f"{len(french_urls_df)}"
    )

    # ========================================================
    # CUSTOMER INPUTS
    # ========================================================

    banner(
        "CUSTOMER INPUT SUMMARY"
    )

    heading_inputs = (
        count_inputs(
            french_content_df,
            COL_FRENCH_HEADING,
        )
    )

    url_name_inputs = (
        count_inputs(
            french_content_df,
            COL_FRENCH_URL_NAME,
        )
    )

    content_inputs = (
        count_inputs(
            french_content_df,
            COL_FRENCH_CONTENT,
        )
    )

    button_inputs = (
        count_inputs(
            french_urls_df,
            COL_FRENCH_BUTTON_LINK,
        )
    )

    content_url_inputs = (
        count_inputs(
            french_urls_df,
            COL_FRENCH_CONTENT_URL,
        )
    )

    print(
        f"French translation of column B: "
        f"{heading_inputs}"
    )

    print(
        f"French translation of column C: "
        f"{url_name_inputs}"
    )

    print(
        f"Content Text Changes French: "
        f"{content_inputs}"
    )

    print(
        f"French Button Link corresponding "
        f"to column D: "
        f"{button_inputs}"
    )

    print(
        f"French Content URL corresponding "
        f"to column J: "
        f"{content_url_inputs}"
    )

    print()

    print(
        "Total customer inputs: "
        f"{heading_inputs + url_name_inputs + content_inputs + button_inputs + content_url_inputs}"
    )

    audit_rows = []

    # ========================================================
    # CONTENT FIRST
    # ========================================================

    banner(
        "APPLYING CONTENT TEXT TRANSLATIONS"
    )

    content_updated = (
        apply_content_text_translations(
            english_df,
            french_df,
            french_content_df,
            comparison_lookup,
            english_uuid_to_french_indexes,
            audit_rows,
        )
    )

    print(
        "Content text records updated: "
        f"{content_updated}"
    )

    # ========================================================
    # HEADING
    # ========================================================

    banner(
        "APPLYING HEADING TRANSLATIONS"
    )

    heading_updated = (
        apply_heading_translations(
            english_df,
            french_df,
            french_content_df,
            comparison_lookup,
            english_uuid_to_french_indexes,
            audit_rows,
        )
    )

    print(
        "Heading records updated: "
        f"{heading_updated}"
    )

    # ========================================================
    # URL NAME
    # ========================================================

    banner(
        "APPLYING URL NAME TRANSLATIONS"
    )

    url_name_updated = (
        apply_url_name_translations(
            english_df,
            french_df,
            french_content_df,
            comparison_lookup,
            english_uuid_to_french_indexes,
            audit_rows,
        )
    )

    print(
        "URL name records updated: "
        f"{url_name_updated}"
    )

    # ========================================================
    # CONTENT URL
    # ========================================================

    banner(
        "APPLYING FRENCH CONTENT URL UPDATES"
    )

    content_url_updated = (
        apply_content_url_updates(
            english_df,
            french_df,
            french_urls_df,
            comparison_lookup,
            english_uuid_to_french_indexes,
            audit_rows,
        )
    )

    print(
        "Content URL records updated: "
        f"{content_url_updated}"
    )

    # ========================================================
    # BUTTON LINK
    # ========================================================

    banner(
        "APPLYING FRENCH BUTTON LINK UPDATES"
    )

    button_updated = (
        apply_button_link_updates(
            english_df,
            french_df,
            french_urls_df,
            comparison_lookup,
            english_uuid_to_french_indexes,
            audit_rows,
        )
    )

    print(
        "Button link records updated: "
        f"{button_updated}"
    )

    # ========================================================
    # VALIDATION
    # ========================================================

    banner(
        "VALIDATING FINAL OUTPUT"
    )

    validate_final_output(
        original_french_df,
        french_df,
    )

    print(
        f"Columns: PASS "
        f"({len(french_df.columns)})"
    )

    print(
        f"Rows: PASS "
        f"({len(french_df)})"
    )

    changed_records = 0

    for index in (
        french_df.index
    ):

        before = (
            original_french_df
            .loc[
                index
            ]
            .fillna("")
            .astype(str)
        )

        after = (
            french_df
            .loc[
                index
            ]
            .fillna("")
            .astype(str)
        )

        if not (
            before
            ==
            after
        ).all():

            changed_records += 1

    print(
        "French records changed: "
        f"{changed_records}"
    )

    # ========================================================
    # SAVE
    # ========================================================

    FINAL_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    french_df.to_csv(
        FINAL_CSV,
        index=False,
        encoding="utf-8-sig",
    )

    write_final_excel(
        french_df
    )

    write_audit(
        audit_rows
    )

    # ========================================================
    # SUMMARY
    # ========================================================

    banner(
        "UPDATE SUMMARY"
    )

    print(
        f"Content text updates : "
        f"{content_updated}"
    )

    print(
        f"Heading updates      : "
        f"{heading_updated}"
    )

    print(
        f"URL name updates     : "
        f"{url_name_updated}"
    )

    print(
        f"Content URL updates  : "
        f"{content_url_updated}"
    )

    print(
        f"Button link updates  : "
        f"{button_updated}"
    )

    print()

    print(
        f"French records changed: "
        f"{changed_records}"
    )

    # ========================================================
    # AUDIT SUMMARY
    # ========================================================

    if audit_rows:

        audit_df = pd.DataFrame(
            audit_rows
        )

        print()
        print(
            "Audit statuses:"
        )

        for (
            status,
            count,
        ) in (
            audit_df[
                "Status"
            ]
            .value_counts()
            .items()
        ):

            print(
                f"  {status:<40} "
                f"{count}"
            )

    # ========================================================
    # OUTPUT FILES
    # ========================================================

    banner(
        "FILES CREATED"
    )

    print()
    print(
        "Final CSV:"
    )
    print(
        FINAL_CSV
    )

    print()
    print(
        "Final Excel:"
    )
    print(
        FINAL_XLSX
    )

    print()
    print(
        "Audit:"
    )
    print(
        AUDIT_XLSX
    )

    print()

    print(
        "IMPORTANT:"
    )

    print(
        "Review REVIEW_REQUIRED and "
        "UPDATED_FROM_EMPTY_FRENCH_CONTENT "
        "records in the audit before CMS import."
    )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()