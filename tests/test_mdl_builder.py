# tests/test_mdl_builder.py
"""Tests for MDL builder service functions."""

import pytest
from unittest.mock import patch, MagicMock
from datetime import datetime

from services.mdl_builder import (
    format_letter_date,
    summarize_finding_text,
    load_finding_mappings,
    best_summary_label,
    build_mdl_model_from_fac,
    render_mdl_html,
)


class TestFormatLetterDate:
    """Tests for format_letter_date function."""

    def test_format_letter_date_with_date(self):
        """Test format_letter_date with specific ISO date."""
        iso_date, long_date = format_letter_date("2024-06-15")

        assert iso_date == "2024-06-15"
        assert "June" in long_date
        assert "15" in long_date
        assert "2024" in long_date

    def test_format_letter_date_without_date(self):
        """Test format_letter_date without date uses current date."""
        iso_date, long_date = format_letter_date(None)

        # Should return today's date in both formats
        today = datetime.utcnow()
        assert str(today.year) in iso_date
        assert str(today.year) in long_date


class TestSummarizeFindingText:
    """Tests for summarize_finding_text function."""

    def test_summarize_finding_text_empty(self):
        """Test summarize_finding_text with empty string."""
        assert summarize_finding_text("") == ""

    def test_summarize_finding_text_short(self):
        """Test summarize_finding_text with short text."""
        text = "The City did not verify suspension status."
        result = summarize_finding_text(text)

        assert result == text.strip()

    def test_summarize_finding_text_truncates_long(self):
        """Test summarize_finding_text truncates long text."""
        text = "This is a very long finding. " * 100
        result = summarize_finding_text(text, max_chars=200)

        assert len(result) <= 200

    def test_summarize_finding_text_skips_metadata_sentences(self):
        """Test summarize_finding_text skips metadata sentences."""
        text = (
            "Assistance Listing Number 21.027. "
            "The City failed to implement proper controls. "
            "This resulted in a significant deficiency."
        )
        result = summarize_finding_text(text)

        # Should not include the Assistance Listing sentence
        assert "Assistance Listing Number" not in result

    def test_summarize_finding_text_normalizes_whitespace(self):
        """Test summarize_finding_text normalizes whitespace."""
        text = "Too    much    whitespace    here."
        result = summarize_finding_text(text)

        assert "    " not in result


class TestLoadFindingMappings:
    """Tests for load_finding_mappings function."""

    def test_load_finding_mappings_no_path(self):
        """Test load_finding_mappings with no path."""
        type_map, summary_labels = load_finding_mappings(None)

        assert type_map == {}
        assert summary_labels == []

    def test_load_finding_mappings_nonexistent_file(self):
        """Test load_finding_mappings with nonexistent file."""
        type_map, summary_labels = load_finding_mappings("/nonexistent/file.xlsx")

        assert type_map == {}
        assert summary_labels == []


class TestBestSummaryLabel:
    """Tests for best_summary_label function."""

    def test_best_summary_label_empty_summary(self):
        """Test best_summary_label with empty summary."""
        result = best_summary_label("", ["Label 1", "Label 2"])
        assert result is None

    def test_best_summary_label_empty_labels(self):
        """Test best_summary_label with empty labels list."""
        result = best_summary_label("Some summary", [])
        assert result is None

    def test_best_summary_label_finds_match(self):
        """Test best_summary_label finds closest match."""
        labels = [
            "Lack of evidence of suspension and debarment verification",
            "Inadequate subrecipient monitoring",
            "Missing procurement documentation",
        ]
        summary = "The auditor found lack of evidence for suspension verification."

        result = best_summary_label(summary, labels)

        # Should find a match (closest one)
        assert result is not None
        assert result in labels


class TestBuildMdlModelFromFac:
    """Tests for build_mdl_model_from_fac function."""

    def test_build_mdl_model_basic(self, sample_fac_general, sample_fac_findings,
                                   sample_fac_findings_text, sample_fac_caps,
                                   sample_federal_awards):
        """Test build_mdl_model_from_fac creates valid model."""
        model = build_mdl_model_from_fac(
            auditee_name="City of Ann Arbor",
            ein="386004534",
            audit_year=2023,
            fac_general=sample_fac_general,
            fac_findings=sample_fac_findings,
            fac_findings_text=sample_fac_findings_text,
            fac_caps=sample_fac_caps,
            federal_awards=sample_federal_awards,
        )

        assert model["auditee_name"] == "City of Ann Arbor"
        assert "38-6004534" in model["ein"]
        assert model["audit_year"] == 2023
        assert "letter_date_iso" in model
        assert "programs" in model

    def test_build_mdl_model_empty_findings(self):
        """Test build_mdl_model_from_fac with empty findings."""
        model = build_mdl_model_from_fac(
            auditee_name="Test City",
            ein="123456789",
            audit_year=2023,
            fac_general=[],
            fac_findings=[],
            fac_findings_text=[],
            fac_caps=[],
            federal_awards=[],
        )

        assert model["auditee_name"] == "Test City"
        assert model["programs"] == []

    def test_build_mdl_model_formats_ein(self):
        """Test build_mdl_model_from_fac formats EIN correctly."""
        model = build_mdl_model_from_fac(
            auditee_name="Test City",
            ein="123456789",
            audit_year=2023,
            fac_general=[],
            fac_findings=[],
            fac_findings_text=[],
            fac_caps=[],
            federal_awards=[],
        )

        assert model["ein"] == "12-3456789"

    def test_build_mdl_model_with_treasury_filter(self, sample_fac_findings,
                                                   sample_fac_findings_text,
                                                   sample_fac_caps,
                                                   sample_federal_awards):
        """Test build_mdl_model_from_fac with treasury listings filter."""
        model = build_mdl_model_from_fac(
            auditee_name="Test City",
            ein="123456789",
            audit_year=2023,
            fac_general=[],
            fac_findings=sample_fac_findings,
            fac_findings_text=sample_fac_findings_text,
            fac_caps=sample_fac_caps,
            federal_awards=sample_federal_awards,
            treasury_listings=["21.027"],
        )

        # Should filter to only treasury programs
        for program in model.get("programs", []):
            aln = program.get("assistance_listing", "")
            if aln != "Unknown":
                assert aln in ["21.027"]


class TestRenderMdlHtml:
    """Tests for render_mdl_html function."""

    def test_render_mdl_html_basic(self, sample_mdl_model):
        """Test render_mdl_html generates valid HTML."""
        html = render_mdl_html(sample_mdl_model)

        assert "<div" in html
        assert "</div>" in html
        assert "City of Ann Arbor" in html or "CITY OF ANN ARBOR" in html

    def test_render_mdl_html_includes_date(self, sample_mdl_model):
        """Test render_mdl_html includes letter date."""
        html = render_mdl_html(sample_mdl_model)

        # Should include formatted date
        assert "June" in html or "2024" in html

    def test_render_mdl_html_includes_ein(self, sample_mdl_model):
        """Test render_mdl_html includes EIN."""
        html = render_mdl_html(sample_mdl_model)

        assert "EIN" in html
        assert "38-6004534" in html

    def test_render_mdl_html_includes_programs(self, sample_mdl_model):
        """Test render_mdl_html includes program tables."""
        html = render_mdl_html(sample_mdl_model)

        # Should include table elements
        assert "<table" in html

    def test_render_mdl_html_includes_treasury_info(self, sample_mdl_model):
        """Test render_mdl_html includes Treasury department info."""
        html = render_mdl_html(sample_mdl_model)

        assert "DEPARTMENT OF THE TREASURY" in html
        assert "WASHINGTON" in html

    def test_render_mdl_html_includes_contact_email(self, sample_mdl_model):
        """Test render_mdl_html includes contact email placeholder or email text."""
        html = render_mdl_html(sample_mdl_model)

        # The email section should contain mailto link or email reference
        assert "mailto:" in html or "email" in html.lower()
