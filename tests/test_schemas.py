# tests/test_schemas.py
"""Tests for Pydantic schema models."""

import pytest
from pydantic import ValidationError
from models.schemas import (
    BuildDocx,
    BuildFromFAC,
    BuildByReport,
    BuildByReportTemplated,
    BuildAuto,
)


class TestBuildDocx:
    """Tests for BuildDocx schema."""

    def test_build_docx_required_fields(self):
        """Test BuildDocx requires auditee_name, ein, and audit_year."""
        model = BuildDocx(auditee_name="Test City", ein="12-3456789", audit_year=2023)
        assert model.auditee_name == "Test City"
        assert model.ein == "12-3456789"
        assert model.audit_year == 2023

    def test_build_docx_missing_required_field(self):
        """Test BuildDocx raises error when required field is missing."""
        with pytest.raises(ValidationError):
            BuildDocx(auditee_name="Test City", ein="12-3456789")

    def test_build_docx_optional_fields_default_none(self):
        """Test BuildDocx optional fields default to None."""
        model = BuildDocx(auditee_name="Test City", ein="12-3456789", audit_year=2023)
        assert model.body_html is None
        assert model.body_html_b64 is None
        assert model.dest_path is None
        assert model.filename is None

    def test_build_docx_with_optional_fields(self):
        """Test BuildDocx with optional fields provided."""
        model = BuildDocx(
            auditee_name="Test City",
            ein="12-3456789",
            audit_year=2023,
            body_html="<p>Test</p>",
            filename="test.docx"
        )
        assert model.body_html == "<p>Test</p>"
        assert model.filename == "test.docx"


class TestBuildFromFAC:
    """Tests for BuildFromFAC schema."""

    def test_build_from_fac_required_fields(self):
        """Test BuildFromFAC requires auditee_name, ein, and audit_year."""
        model = BuildFromFAC(auditee_name="Test City", ein="12-3456789", audit_year=2023)
        assert model.auditee_name == "Test City"
        assert model.ein == "12-3456789"
        assert model.audit_year == 2023

    def test_build_from_fac_list_defaults(self):
        """Test BuildFromFAC list fields default to empty lists."""
        model = BuildFromFAC(auditee_name="Test City", ein="12-3456789", audit_year=2023)
        assert model.fac_general == []
        assert model.fac_findings == []
        assert model.fac_findings_text == []
        assert model.fac_caps == []
        assert model.federal_awards == []

    def test_build_from_fac_with_data(self):
        """Test BuildFromFAC with FAC data provided."""
        model = BuildFromFAC(
            auditee_name="Test City",
            ein="12-3456789",
            audit_year=2023,
            fac_general=[{"report_id": "test-123"}],
            fac_findings=[{"reference_number": "2023-001"}]
        )
        assert len(model.fac_general) == 1
        assert model.fac_general[0]["report_id"] == "test-123"
        assert len(model.fac_findings) == 1


class TestBuildByReport:
    """Tests for BuildByReport schema."""

    def test_build_by_report_required_fields(self):
        """Test BuildByReport requires report_id in addition to base fields."""
        model = BuildByReport(
            auditee_name="Test City",
            ein="12-3456789",
            audit_year=2023,
            report_id="2023-06-GSAFAC-0000123456"
        )
        assert model.report_id == "2023-06-GSAFAC-0000123456"

    def test_build_by_report_missing_report_id(self):
        """Test BuildByReport raises error when report_id is missing."""
        with pytest.raises(ValidationError):
            BuildByReport(auditee_name="Test City", ein="12-3456789", audit_year=2023)

    def test_build_by_report_defaults(self):
        """Test BuildByReport default values."""
        model = BuildByReport(
            auditee_name="Test City",
            ein="12-3456789",
            audit_year=2023,
            report_id="test-123"
        )
        assert model.only_flagged is False
        assert model.max_refs == 15
        assert model.include_awards is True


class TestBuildByReportTemplated:
    """Tests for BuildByReportTemplated schema."""

    def test_build_by_report_templated_inherits(self):
        """Test BuildByReportTemplated inherits from BuildByReport."""
        model = BuildByReportTemplated(
            auditee_name="Test City",
            ein="12-3456789",
            audit_year=2023,
            report_id="test-123"
        )
        assert model.report_id == "test-123"
        assert model.max_refs == 15

    def test_build_by_report_templated_additional_fields(self):
        """Test BuildByReportTemplated has additional optional fields."""
        model = BuildByReportTemplated(
            auditee_name="Test City",
            ein="12-3456789",
            audit_year=2023,
            report_id="test-123",
            auditor_name="Rehmann Robson LLC",
            fy_end_text="June 30, 2023",
            city="Ann Arbor",
            state="MI"
        )
        assert model.auditor_name == "Rehmann Robson LLC"
        assert model.fy_end_text == "June 30, 2023"
        assert model.city == "Ann Arbor"
        assert model.state == "MI"

    def test_build_by_report_templated_treasury_listings(self):
        """Test BuildByReportTemplated treasury_listings field."""
        model = BuildByReportTemplated(
            auditee_name="Test City",
            ein="12-3456789",
            audit_year=2023,
            report_id="test-123",
            treasury_listings=["21.027", "21.023"]
        )
        assert model.treasury_listings == ["21.027", "21.023"]


class TestBuildAuto:
    """Tests for BuildAuto schema."""

    def test_build_auto_required_fields(self):
        """Test BuildAuto requires only ein and audit_year (auditee_name is optional)."""
        model = BuildAuto(ein="12-3456789", audit_year=2023)
        assert model.ein == "12-3456789"
        assert model.audit_year == 2023
        assert model.auditee_name is None  # Optional, defaults to None

    def test_build_auto_with_auditee_name(self):
        """Test BuildAuto with optional auditee_name provided."""
        model = BuildAuto(auditee_name="Test City", ein="12-3456789", audit_year=2023)
        assert model.auditee_name == "Test City"
        assert model.ein == "12-3456789"
        assert model.audit_year == 2023

    def test_build_auto_with_county_name(self):
        """Test BuildAuto with optional county_name provided."""
        model = BuildAuto(ein="12-3456789", audit_year=2023, county_name="Travis County")
        assert model.county_name == "Travis County"

    def test_build_auto_defaults(self):
        """Test BuildAuto default values."""
        model = BuildAuto(ein="12-3456789", audit_year=2023)
        assert model.max_refs == 15
        assert model.only_flagged is False
        assert model.include_awards is True
        assert model.include_no_qc_line is True
        assert model.include_no_cap_line is False
        assert model.auditee_name is None
        assert model.county_name is None

    def test_build_auto_all_optional_fields(self):
        """Test BuildAuto with all optional fields."""
        model = BuildAuto(
            auditee_name="Test City",
            ein="12-3456789",
            audit_year=2023,
            dest_path="/custom/path",
            recipient_name="City of Test",
            fy_end_text="June 30, 2023",
            auditor_name="Test Auditor LLC",
            street_address="123 Main St",
            city="Test City",
            state="TX",
            zip_code="12345",
            poc_name="John Doe",
            poc_title="Finance Director",
            template_path="/templates/custom.docx",
            treasury_contact_email="test@treasury.gov",
            treasury_listings=["21.027"]
        )
        assert model.dest_path == "/custom/path"
        assert model.recipient_name == "City of Test"
        assert model.treasury_contact_email == "test@treasury.gov"
        assert model.treasury_listings == ["21.027"]
