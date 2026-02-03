# tests/test_routes.py
"""Tests for API routes."""

import pytest
import base64
from unittest.mock import patch, MagicMock
from fastapi.testclient import TestClient


class TestHealthzRoute:
    """Tests for /healthz endpoint."""

    def test_healthz_returns_ok(self, test_client):
        """Test healthz endpoint returns ok status."""
        response = test_client.get("/healthz")

        assert response.status_code == 200
        data = response.json()
        assert data["ok"] is True
        assert "time" in data


class TestEchoRoute:
    """Tests for /echo endpoint."""

    def test_echo_returns_payload(self, test_client):
        """Test echo endpoint returns the received payload."""
        payload = {"test": "data", "number": 123}
        response = test_client.post("/echo", json=payload)

        assert response.status_code == 200
        data = response.json()
        assert data["received"] == payload
        assert "ts" in data


class TestDebugEnvRoute:
    """Tests for /debug/env endpoint."""

    def test_debug_env_returns_key_status(self, test_client, set_test_env_vars):
        """Test debug/env endpoint returns API key status."""
        response = test_client.get("/debug/env")

        assert response.status_code == 200
        data = response.json()
        assert "fac_api_key_present" in data
        assert "fac_api_key_masked" in data


class TestDebugStorageRoute:
    """Tests for /debug/storage endpoint."""

    def test_debug_storage_returns_status(self, test_client, set_test_env_vars):
        """Test debug/storage endpoint returns storage status."""
        response = test_client.get("/debug/storage")

        assert response.status_code == 200
        data = response.json()
        assert "using_storage" in data


class TestLocalFileRoute:
    """Tests for /local/{path} endpoint."""

    def test_get_local_file_not_found(self, test_client):
        """Test local file endpoint returns 404 for non-existent file."""
        response = test_client.get("/local/nonexistent/file.docx")

        assert response.status_code == 404


class TestBuildDocxRoute:
    """Tests for /build-docx endpoint."""

    def test_build_docx_missing_body(self, test_client, set_test_env_vars):
        """Test build-docx endpoint requires body_html or body_html_b64."""
        payload = {
            "auditee_name": "Test City",
            "ein": "123456789",
            "audit_year": 2023
        }

        response = test_client.post("/build-docx", json=payload)

        assert response.status_code == 400
        assert "body_html" in response.text.lower()

    @patch('routes.docx_routes.save_local_and_url')
    def test_build_docx_with_html(self, mock_save, test_client, set_test_env_vars):
        """Test build-docx endpoint with HTML body."""
        mock_save.return_value = "http://localhost:8000/local/test.docx"

        payload = {
            "auditee_name": "Test City",
            "ein": "123456789",
            "audit_year": 2023,
            "body_html": "<p>Test content</p>"
        }

        response = test_client.post("/build-docx", json=payload)

        assert response.status_code == 200
        data = response.json()
        assert "url" in data

    @patch('routes.docx_routes.save_local_and_url')
    def test_build_docx_with_base64(self, mock_save, test_client, set_test_env_vars):
        """Test build-docx endpoint with base64-encoded HTML body."""
        mock_save.return_value = "http://localhost:8000/local/test.docx"

        html_content = "<p>Base64 test content</p>"
        html_b64 = base64.b64encode(html_content.encode()).decode()

        payload = {
            "auditee_name": "Test City",
            "ein": "123456789",
            "audit_year": 2023,
            "body_html_b64": html_b64
        }

        response = test_client.post("/build-docx", json=payload)

        assert response.status_code == 200
        data = response.json()
        assert "url" in data


class TestBuildDocxFromFacRoute:
    """Tests for /build-docx-from-fac endpoint."""

    @patch('routes.docx_routes.save_local_and_url')
    def test_build_docx_from_fac_basic(self, mock_save, test_client, set_test_env_vars,
                                        sample_fac_general, sample_fac_findings,
                                        sample_fac_findings_text, sample_fac_caps,
                                        sample_federal_awards):
        """Test build-docx-from-fac endpoint creates document."""
        mock_save.return_value = "http://localhost:8000/local/test.docx"

        payload = {
            "auditee_name": "Test City",
            "ein": "123456789",
            "audit_year": 2023,
            "fac_general": sample_fac_general,
            "fac_findings": sample_fac_findings,
            "fac_findings_text": sample_fac_findings_text,
            "fac_caps": sample_fac_caps,
            "federal_awards": sample_federal_awards
        }

        response = test_client.post("/build-docx-from-fac", json=payload)

        assert response.status_code == 200
        data = response.json()
        assert "url" in data

    @patch('routes.docx_routes.save_local_and_url')
    def test_build_docx_from_fac_empty_findings(self, mock_save, test_client, set_test_env_vars):
        """Test build-docx-from-fac endpoint with no findings."""
        mock_save.return_value = "http://localhost:8000/local/test.docx"

        payload = {
            "auditee_name": "Test City",
            "ein": "123456789",
            "audit_year": 2023,
            "fac_general": [],
            "fac_findings": [],
            "fac_findings_text": [],
            "fac_caps": [],
            "federal_awards": []
        }

        response = test_client.post("/build-docx-from-fac", json=payload)

        assert response.status_code == 200
        data = response.json()
        assert "url" in data


class TestBuildDocxByReportRoute:
    """Tests for /build-docx-by-report endpoint."""

    @patch('routes.docx_routes.fac_get')
    @patch('routes.docx_routes.save_local_and_url')
    def test_build_docx_by_report_basic(self, mock_save, mock_fac_get, test_client, set_test_env_vars):
        """Test build-docx-by-report endpoint."""
        mock_save.return_value = "http://localhost:8000/local/test.docx"
        mock_fac_get.return_value = []

        payload = {
            "auditee_name": "Test City",
            "ein": "123456789",
            "audit_year": 2023,
            "report_id": "2023-06-GSAFAC-0000123456"
        }

        response = test_client.post("/build-docx-by-report", json=payload)

        assert response.status_code == 200
        data = response.json()
        assert "url" in data


class TestBuildMdlDocxByReportTemplatedRoute:
    """Tests for /build-mdl-docx-by-report-templated endpoint."""

    @patch('routes.docx_routes.fac_get')
    def test_build_mdl_docx_by_report_templated_no_template(self, mock_fac_get, test_client, set_test_env_vars):
        """Test build-mdl-docx-by-report-templated endpoint without template."""
        mock_fac_get.return_value = []

        payload = {
            "auditee_name": "Test City",
            "ein": "123456789",
            "audit_year": 2023,
            "report_id": "2023-06-GSAFAC-0000123456"
        }

        response = test_client.post("/build-mdl-docx-by-report-templated", json=payload)

        assert response.status_code == 200
        data = response.json()
        # Should indicate template path not provided
        assert "ok" in data


class TestBuildMdlDocxAutoRoute:
    """Tests for /build-mdl-docx-auto endpoint."""

    @patch('routes.docx_routes.fac_get')
    def test_build_mdl_docx_auto_no_fac_records(self, mock_fac_get, test_client, set_test_env_vars):
        """Test build-mdl-docx-auto endpoint when no FAC records found for EIN."""
        mock_fac_get.return_value = []

        payload = {
            "ein": "123456789",
            "audit_year": 2023
        }

        response = test_client.post("/build-mdl-docx-auto", json=payload)

        assert response.status_code == 200
        data = response.json()
        assert data["ok"] is False
        assert "No FAC records found" in data["message"]

    @patch('routes.docx_routes.fac_get')
    def test_build_mdl_docx_auto_no_report_for_year(self, mock_fac_get, test_client, set_test_env_vars):
        """Test build-mdl-docx-auto endpoint when no report found for input year."""
        mock_fac_get.side_effect = [
            # First call: latest year (found)
            [{
                "report_id": "2024-06-GSAFAC-0000123456",
                "audit_year": 2024,
                "auditee_name": "Test City",
                "fac_accepted_date": "2024-06-15"
            }],
            # Second call: input year (not found)
            []
        ]

        payload = {
            "ein": "123456789",
            "audit_year": 2023
        }

        response = test_client.post("/build-mdl-docx-auto", json=payload)

        assert response.status_code == 200
        data = response.json()
        assert data["ok"] is False
        assert "No FAC report found" in data["message"]
        assert "2023" in data["message"]

    @patch('routes.docx_routes.fac_get')
    @patch('routes.docx_routes.aln_overrides_from_summary')
    @patch('routes.docx_routes.build_docx_from_template')
    @patch('routes.docx_routes.postprocess_docx')
    @patch('routes.docx_routes.save_local_and_url')
    def test_build_mdl_docx_auto_success_minimal_input(self, mock_save, mock_postprocess,
                                          mock_build_template, mock_aln_overrides,
                                          mock_fac_get, test_client, set_test_env_vars):
        """Test build-mdl-docx-auto endpoint with minimal input (only ein and audit_year)."""
        mock_fac_get.side_effect = [
            # First call: latest year for auditee_name
            [{
                "report_id": "2024-06-GSAFAC-0000123456",
                "audit_year": 2024,
                "fac_accepted_date": "2024-06-15",
                "auditee_address_line_1": "123 Main St",
                "auditee_city": "Austin",
                "auditee_state": "TX",
                "auditee_zip": "78701",
                "auditor_firm_name": "Test Auditors LLC",
                "fy_end_date": "2024-06-30",
                "auditee_contact_name": "John Doe",
                "auditee_contact_title": "Finance Director",
                "auditee_name": "Test City From FAC"
            }],
            # Second call: input year general info
            [{
                "report_id": "2023-06-GSAFAC-0000123456",
                "fac_accepted_date": "2023-06-15",
                "auditee_address_line_1": "123 Main St",
                "auditee_city": "Austin",
                "auditee_state": "TX",
                "auditee_zip": "78701",
                "auditor_firm_name": "Test Auditors LLC",
                "fy_end_date": "2023-06-30",
                "auditee_contact_name": "John Doe",
                "auditee_contact_title": "Finance Director",
                "auditee_name": "Test City Old Name"
            }],
            # Third call: findings
            [],
            # Fourth call: federal awards
            []
        ]
        mock_aln_overrides.return_value = ({}, {})
        mock_build_template.return_value = b"test document bytes"
        mock_postprocess.return_value = b"processed document bytes"
        mock_save.return_value = "http://localhost:8000/local/test.docx"

        # Minimal payload - no auditee_name provided
        payload = {
            "ein": "123456789",
            "audit_year": 2023
        }

        response = test_client.post("/build-mdl-docx-auto", json=payload)

        assert response.status_code == 200
        data = response.json()
        assert data["ok"] is True
        assert "url" in data

    @patch('routes.docx_routes.fac_get')
    @patch('routes.docx_routes.aln_overrides_from_summary')
    @patch('routes.docx_routes.build_docx_from_template')
    @patch('routes.docx_routes.postprocess_docx')
    @patch('routes.docx_routes.save_local_and_url')
    def test_build_mdl_docx_auto_success_with_auditee_name(self, mock_save, mock_postprocess,
                                          mock_build_template, mock_aln_overrides,
                                          mock_fac_get, test_client, set_test_env_vars):
        """Test build-mdl-docx-auto endpoint success case with auditee_name provided."""
        mock_fac_get.side_effect = [
            # First call: latest year for auditee_name (FAC name takes precedence)
            [{
                "report_id": "2023-06-GSAFAC-0000123456",
                "audit_year": 2023,
                "fac_accepted_date": "2023-06-15",
                "auditee_address_line_1": "123 Main St",
                "auditee_city": "Austin",
                "auditee_state": "TX",
                "auditee_zip": "78701",
                "auditor_firm_name": "Test Auditors LLC",
                "fy_end_date": "2023-06-30",
                "auditee_contact_name": "John Doe",
                "auditee_contact_title": "Finance Director",
                "auditee_name": "Test City From FAC"
            }],
            # Second call: input year general info
            [{
                "report_id": "2023-06-GSAFAC-0000123456",
                "fac_accepted_date": "2023-06-15",
                "auditee_address_line_1": "123 Main St",
                "auditee_city": "Austin",
                "auditee_state": "TX",
                "auditee_zip": "78701",
                "auditor_firm_name": "Test Auditors LLC",
                "fy_end_date": "2023-06-30",
                "auditee_contact_name": "John Doe",
                "auditee_contact_title": "Finance Director",
                "auditee_name": "Test City From FAC"
            }],
            # Third call: findings
            [],
            # Fourth call: federal awards
            []
        ]
        mock_aln_overrides.return_value = ({}, {})
        mock_build_template.return_value = b"test document bytes"
        mock_postprocess.return_value = b"processed document bytes"
        mock_save.return_value = "http://localhost:8000/local/test.docx"

        payload = {
            "auditee_name": "Test City",
            "ein": "123456789",
            "audit_year": 2023
        }

        response = test_client.post("/build-mdl-docx-auto", json=payload)

        assert response.status_code == 200
        data = response.json()
        assert data["ok"] is True
        assert "url" in data


class TestRouteValidation:
    """Tests for route input validation."""

    def test_build_docx_requires_auditee_name(self, test_client):
        """Test build-docx requires auditee_name."""
        payload = {
            "ein": "123456789",
            "audit_year": 2023,
            "body_html": "<p>Test</p>"
        }

        response = test_client.post("/build-docx", json=payload)

        assert response.status_code == 422

    def test_build_docx_requires_ein(self, test_client):
        """Test build-docx requires ein."""
        payload = {
            "auditee_name": "Test City",
            "audit_year": 2023,
            "body_html": "<p>Test</p>"
        }

        response = test_client.post("/build-docx", json=payload)

        assert response.status_code == 422

    def test_build_docx_requires_audit_year(self, test_client):
        """Test build-docx requires audit_year."""
        payload = {
            "auditee_name": "Test City",
            "ein": "123456789",
            "body_html": "<p>Test</p>"
        }

        response = test_client.post("/build-docx", json=payload)

        assert response.status_code == 422
