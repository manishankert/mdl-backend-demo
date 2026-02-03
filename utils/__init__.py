# utils/__init__.py
from .text_utils import (
    sanitize,
    short_text,
    norm_ref,
    norm_txt,
    title_case,
    title_with_article,
    no_article,
    title_with_acronyms,
    allcaps,
    with_The_allcaps,
    with_the_allcaps,
    format_name_standard_case,
)
from .docx_utils import (
    shade_cell,
    set_col_widths,
    tight_paragraph,
    as_oxml,
    insert_after,
    apply_grid_borders,
    remove_paragraph,
    clear_runs,
    para_text,
    rewrite_para_text,
)

__all__ = [
    # text utils
    "sanitize",
    "short_text",
    "norm_ref",
    "norm_txt",
    "title_case",
    "title_with_article",
    "no_article",
    "title_with_acronyms",
    "allcaps",
    "with_The_allcaps",
    "with_the_allcaps",
    "format_name_standard_case",
    # docx utils
    "shade_cell",
    "set_col_widths",
    "tight_paragraph",
    "as_oxml",
    "insert_after",
    "apply_grid_borders",
    "remove_paragraph",
    "clear_runs",
    "para_text",
    "rewrite_para_text",
]
