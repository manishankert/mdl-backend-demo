# tests/test_fac_api.py
"""Tests for FAC API service functions."""

import pytest
from unittest.mock import patch, MagicMock
from fastapi import HTTPException

from services.fac_api import fac_headers, fac_get, or_param, from_fac_general


class TestFacHeaders:
    """Tests for fac_headers function."""

    def test_fac_headers_with_env_var(self):
        """Test fac_headers returns headers with API key from environment."""
        with patch.dict('os.environ', {'FAC_API_KEY': 'test-key-123'}):
            headers = fac_headers()
            assert headers == {"X-Api-Key": "test-key-123"}

    def test_fac_headers_no_key_raises_error(self):
        """Test fac_headers raises HTTPException when no API key is set."""
        with patch.dict('os.environ', {'FAC_API_KEY': ''}, clear=True):
            with patch('services.fac_api.FAC_KEY', None):
                with pytest.raises(HTTPException) as exc_info:
                    fac_headers()
                assert exc_info.value.status_code == 500
                assert "FAC_API_KEY not configured" in str(exc_info.value.detail)


class TestOrParam:
    """Tests for or_param function."""

    def test_or_param_single_value(self):
        """Test or_param with single value."""
        result = or_param("field", ["value1"])
        assert result == "(field.eq.value1)"

    def test_or_param_multiple_values(self):
        """Test or_param with multiple values."""
        result = or_param("status", ["active", "pending", "completed"])
        expected = "(status.eq.active,status.eq.pending,status.eq.completed)"
        assert result == expected

    def test_or_param_empty_list(self):
        """Test or_param with empty list."""
        result = or_param("field", [])
        assert result == "()"


class TestFacGet:
    """Tests for fac_get function."""

    @patch('services.fac_api.requests.get')
    @patch('services.fac_api.fac_headers')
    def test_fac_get_success(self, mock_headers, mock_get):
        """Test fac_get returns JSON on success."""
        mock_headers.return_value = {"X-Api-Key": "test-key"}
        mock_response = MagicMock()
        mock_response.json.return_value = [{"id": 1, "name": "test"}]
        mock_response.raise_for_status.return_value = None
        mock_get.return_value = mock_response

        result = fac_get("/test-path", {"param": "value"})

        assert result == [{"id": 1, "name": "test"}]
        mock_get.assert_called_once()

    @patch('services.fac_api.requests.get')
    @patch('services.fac_api.fac_headers')
    def test_fac_get_http_error(self, mock_headers, mock_get):
        """Test fac_get raises HTTPException on HTTP error."""
        import requests

        mock_headers.return_value = {"X-Api-Key": "test-key"}
        mock_response = MagicMock()
        mock_response.status_code = 404
        mock_response.text = "Not found"
        mock_response.raise_for_status.side_effect = requests.HTTPError()
        mock_get.return_value = mock_response

        with pytest.raises(HTTPException):
            fac_get("/test-path", {"param": "value"})


class TestFromFacGeneral:
    """Tests for from_fac_general function."""

    def test_from_fac_general_empty(self):
        """Test from_fac_general with empty list."""
        result = from_fac_general([])
        assert result == {}

    def test_from_fac_general_none(self):
        """Test from_fac_general with None."""
        result = from_fac_general(None)
        assert result == {}

    def test_from_fac_general_extracts_fields(self):
        """Test from_fac_general extracts expected fields."""
        data = [{
            "auditee_address_line_1": "123 Main St",
            "auditee_city": "Austin",
            "auditee_state": "TX",
            "auditee_zip": "78701",
            "auditor_firm_name": "Test Auditors LLC",
            "fy_end_date": "2023-06-30",
            "auditee_contact_name": "John Doe",
            "auditee_contact_title": "Finance Director",
        }]

        result = from_fac_general(data)

        assert result["street_address"] == "123 Main St"
        assert result["city"] == "Austin"
        assert result["state"] == "TX"
        assert result["zip_code"] == "78701"
        assert result["auditor_name"] == "Test Auditors LLC"
        assert result["poc_name"] == "John Doe"
        assert result["poc_title"] == "Finance Director"

    def test_from_fac_general_formats_date(self):
        """Test from_fac_general formats fiscal year end date."""
        data = [{
            "fy_end_date": "2023-06-30",
            "auditee_address_line_1": "",
            "auditee_city": "",
            "auditee_state": "",
            "auditee_zip": "",
            "auditor_firm_name": "",
            "auditee_contact_name": "",
            "auditee_contact_title": "",
        }]

        result = from_fac_general(data)

        # Should be formatted as "Month Day, Year"
        assert result["period_end_text"] is not None
        assert "June" in result["period_end_text"]
        assert "2023" in result["period_end_text"]

    def test_from_fac_general_handles_missing_fields(self):
        """Test from_fac_general handles missing optional fields."""
        data = [{}]

        result = from_fac_general(data)

        assert result["street_address"] == ""
        assert result["city"] == ""
        assert result["state"] == ""
        assert result["zip_code"] == ""
        assert result["auditor_name"] == ""
        assert result["poc_name"] == ""
        assert result["poc_title"] == ""
