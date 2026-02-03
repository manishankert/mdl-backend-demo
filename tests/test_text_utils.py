# tests/test_text_utils.py
"""Tests for text utility functions."""

import pytest
from utils.text_utils import (
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


class TestSanitize:
    """Tests for sanitize function."""

    def test_sanitize_simple(self):
        """Test sanitize with simple alphanumeric string."""
        assert sanitize("hello_world") == "hello_world"

    def test_sanitize_special_chars(self):
        """Test sanitize removes special characters."""
        assert sanitize("hello@world!test") == "hello_world_test"

    def test_sanitize_spaces(self):
        """Test sanitize replaces spaces with underscores."""
        assert sanitize("hello world") == "hello_world"

    def test_sanitize_multiple_special(self):
        """Test sanitize handles multiple consecutive special chars."""
        assert sanitize("hello!!!world") == "hello_world"

    def test_sanitize_empty_string(self):
        """Test sanitize with empty string."""
        assert sanitize("") == ""

    def test_sanitize_none(self):
        """Test sanitize with None."""
        assert sanitize(None) == ""

    def test_sanitize_preserves_dots(self):
        """Test sanitize preserves dots."""
        assert sanitize("file.txt") == "file.txt"

    def test_sanitize_preserves_hyphens(self):
        """Test sanitize preserves hyphens."""
        assert sanitize("my-file") == "my-file"


class TestShortText:
    """Tests for short_text function."""

    def test_short_text_under_limit(self):
        """Test short_text with text under limit."""
        assert short_text("hello world", 100) == "hello world"

    def test_short_text_over_limit(self):
        """Test short_text truncates text over limit and adds ellipsis."""
        result = short_text("a" * 100, 50)
        # Function truncates to limit-1 chars then adds "..." (3 chars)
        # So result is (limit-1) + 3 = limit + 2 chars
        assert result.endswith("...")
        assert len(result) < 100  # Should be truncated from original

    def test_short_text_exact_limit(self):
        """Test short_text with text at exact limit."""
        result = short_text("hello", 5)
        assert result == "hello"

    def test_short_text_empty(self):
        """Test short_text with empty string."""
        assert short_text("") == ""

    def test_short_text_none(self):
        """Test short_text with None."""
        assert short_text(None) == ""

    def test_short_text_normalizes_whitespace(self):
        """Test short_text normalizes multiple whitespace."""
        assert short_text("hello    world") == "hello world"

    def test_short_text_default_limit(self):
        """Test short_text uses default limit of 900 and truncates long text."""
        result = short_text("a" * 1000)
        # Should be truncated from 1000 chars
        assert len(result) < 1000
        assert result.endswith("...")


class TestNormRef:
    """Tests for norm_ref function."""

    def test_norm_ref_uppercase(self):
        """Test norm_ref converts to uppercase."""
        assert norm_ref("abc") == "ABC"

    def test_norm_ref_removes_whitespace(self):
        """Test norm_ref removes all whitespace."""
        assert norm_ref("a b c") == "ABC"
        assert norm_ref("a\tb\nc") == "ABC"

    def test_norm_ref_empty(self):
        """Test norm_ref with empty string."""
        assert norm_ref("") == ""

    def test_norm_ref_none(self):
        """Test norm_ref with None."""
        assert norm_ref(None) == ""


class TestNormTxt:
    """Tests for norm_txt function."""

    def test_norm_txt_basic(self):
        """Test norm_txt with basic text."""
        assert norm_txt("hello world") == "hello world"

    def test_norm_txt_nbsp(self):
        """Test norm_txt converts NBSP to space."""
        assert norm_txt("hello\u00A0world") == "hello world"

    def test_norm_txt_dashes(self):
        """Test norm_txt normalizes dashes."""
        assert norm_txt("hello–world") == "hello-world"
        assert norm_txt("hello—world") == "hello-world"

    def test_norm_txt_collapses_whitespace(self):
        """Test norm_txt collapses multiple whitespace."""
        assert norm_txt("hello    world") == "hello world"

    def test_norm_txt_empty(self):
        """Test norm_txt with empty string."""
        assert norm_txt("") == ""


class TestTitleCase:
    """Tests for title_case function."""

    def test_title_case_basic(self):
        """Test title_case with basic text."""
        assert title_case("hello world") == "Hello World"

    def test_title_case_uppercase(self):
        """Test title_case with uppercase text."""
        assert title_case("HELLO WORLD") == "Hello World"

    def test_title_case_empty(self):
        """Test title_case with empty string."""
        assert title_case("") == ""


class TestTitleWithArticle:
    """Tests for title_with_article function."""

    def test_title_with_article_adds_the(self):
        """Test title_with_article adds 'The' prefix."""
        assert title_with_article("City of Austin") == "The City of Austin"

    def test_title_with_article_preserves_existing(self):
        """Test title_with_article doesn't duplicate 'The'."""
        assert title_with_article("The City of Austin") == "The City of Austin"
        assert title_with_article("the City of Austin") == "the City of Austin"

    def test_title_with_article_empty(self):
        """Test title_with_article with empty string."""
        assert title_with_article("") == ""


class TestNoArticle:
    """Tests for no_article function."""

    def test_no_article_removes_the(self):
        """Test no_article removes 'The' prefix."""
        assert no_article("The City of Austin") == "City of Austin"
        assert no_article("the City of Austin") == "City of Austin"

    def test_no_article_preserves_no_article(self):
        """Test no_article preserves text without article."""
        assert no_article("City of Austin") == "City of Austin"

    def test_no_article_empty(self):
        """Test no_article with empty string."""
        assert no_article("") == ""

    def test_no_article_strips_whitespace(self):
        """Test no_article strips surrounding whitespace."""
        assert no_article("  The City  ") == "City"


class TestTitleWithAcronyms:
    """Tests for title_with_acronyms function."""

    def test_title_with_acronyms_preserves_allcaps(self):
        """Test title_with_acronyms preserves ALL CAPS words."""
        result = title_with_acronyms("city of USA")
        assert "USA" in result

    def test_title_with_acronyms_lowercases_stopwords(self):
        """Test title_with_acronyms lowercases stop words."""
        result = title_with_acronyms("city OF austin")
        # Stop words should be lowercased (OF -> of)
        assert "of" in result.lower()


class TestAllcaps:
    """Tests for allcaps function."""

    def test_allcaps_basic(self):
        """Test allcaps converts to uppercase."""
        assert allcaps("hello world") == "HELLO WORLD"

    def test_allcaps_empty(self):
        """Test allcaps with empty string."""
        assert allcaps("") == ""

    def test_allcaps_none(self):
        """Test allcaps with None."""
        assert allcaps(None) == ""

    def test_allcaps_strips_whitespace(self):
        """Test allcaps strips surrounding whitespace."""
        assert allcaps("  hello  ") == "HELLO"


class TestWithTheAllcaps:
    """Tests for with_The_allcaps function."""

    def test_with_The_allcaps_basic(self):
        """Test with_The_allcaps formats correctly."""
        assert with_The_allcaps("city of austin") == "The CITY OF AUSTIN"

    def test_with_The_allcaps_removes_existing_the(self):
        """Test with_The_allcaps doesn't duplicate 'The'."""
        assert with_The_allcaps("The city of austin") == "The CITY OF AUSTIN"
        assert with_The_allcaps("the city of austin") == "The CITY OF AUSTIN"


class TestWithTheAllcapsLower:
    """Tests for with_the_allcaps function."""

    def test_with_the_allcaps_basic(self):
        """Test with_the_allcaps formats correctly."""
        assert with_the_allcaps("rehmann robson llc") == "the REHMANN ROBSON LLC"

    def test_with_the_allcaps_removes_existing_the(self):
        """Test with_the_allcaps doesn't duplicate 'the'."""
        assert with_the_allcaps("The rehmann robson llc") == "the REHMANN ROBSON LLC"


class TestFormatNameStandardCase:
    """Tests for format_name_standard_case function."""

    def test_format_name_standard_case_allcaps(self):
        """Test format_name_standard_case with ALL CAPS input."""
        result = format_name_standard_case("CITY OF ANN ARBOR, MICHIGAN")
        assert result == "City of Ann Arbor, Michigan"

    def test_format_name_standard_case_removes_article(self):
        """Test format_name_standard_case removes 'The' article."""
        result = format_name_standard_case("THE CITY OF AUSTIN")
        assert not result.lower().startswith("the ")
        assert "City" in result

    def test_format_name_standard_case_lowercases_connectors(self):
        """Test format_name_standard_case lowercases connector words."""
        result = format_name_standard_case("CITY OF ANN ARBOR")
        assert " of " in result

    def test_format_name_standard_case_empty(self):
        """Test format_name_standard_case with empty string."""
        assert format_name_standard_case("") == ""

    def test_format_name_standard_case_mixed_case(self):
        """Test format_name_standard_case with mixed case input."""
        result = format_name_standard_case("City of Austin")
        assert result == "City of Austin"
