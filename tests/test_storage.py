# tests/test_storage.py
"""Tests for storage service functions."""

import os
import pytest
from unittest.mock import patch, MagicMock

from services.storage import parse_conn_str, save_local_and_url


class TestParseConnStr:
    """Tests for parse_conn_str function."""

    def test_parse_conn_str_empty(self):
        """Test parse_conn_str with empty string."""
        result = parse_conn_str("")
        assert result["AccountName"] is None
        assert result["AccountKey"] is None
        assert result["BlobEndpoint"] is None

    def test_parse_conn_str_none(self):
        """Test parse_conn_str with None."""
        result = parse_conn_str(None)
        assert result["AccountName"] is None
        assert result["AccountKey"] is None
        assert result["BlobEndpoint"] is None

    def test_parse_conn_str_development_storage(self):
        """Test parse_conn_str with development storage emulator."""
        result = parse_conn_str("UseDevelopmentStorage=true")
        assert result["AccountName"] == "devstoreaccount1"
        assert result["AccountKey"] is not None
        assert result["BlobEndpoint"] == "http://127.0.0.1:10000/devstoreaccount1"

    def test_parse_conn_str_standard(self):
        """Test parse_conn_str with standard connection string."""
        conn_str = "AccountName=myaccount;AccountKey=mykey;BlobEndpoint=https://myaccount.blob.core.windows.net"
        result = parse_conn_str(conn_str)
        assert result["AccountName"] == "myaccount"
        assert result["AccountKey"] == "mykey"
        assert result["BlobEndpoint"] == "https://myaccount.blob.core.windows.net"

    def test_parse_conn_str_partial(self):
        """Test parse_conn_str with partial connection string."""
        conn_str = "AccountName=testaccount;AccountKey=testkey"
        result = parse_conn_str(conn_str)
        assert result["AccountName"] == "testaccount"
        assert result["AccountKey"] == "testkey"
        assert result["BlobEndpoint"] is None


class TestSaveLocalAndUrl:
    """Tests for save_local_and_url function."""

    def test_save_local_and_url(self, tmp_path):
        """Test save_local_and_url saves file and returns URL."""
        with patch('services.storage.LOCAL_SAVE_DIR', str(tmp_path)):
            with patch('services.storage.PUBLIC_BASE_URL', 'http://localhost:8000'):
                blob_name = "test/output.docx"
                data = b"test document content"

                result = save_local_and_url(blob_name, data)

                # Check file was created
                full_path = tmp_path / "test" / "output.docx"
                assert full_path.exists()
                assert full_path.read_bytes() == data

                # Check URL format
                assert result == "http://localhost:8000/local/test/output.docx"

    def test_save_local_and_url_creates_directories(self, tmp_path):
        """Test save_local_and_url creates nested directories."""
        with patch('services.storage.LOCAL_SAVE_DIR', str(tmp_path)):
            with patch('services.storage.PUBLIC_BASE_URL', 'http://localhost:8000'):
                blob_name = "deep/nested/path/file.docx"
                data = b"content"

                result = save_local_and_url(blob_name, data)

                full_path = tmp_path / "deep" / "nested" / "path" / "file.docx"
                assert full_path.exists()


class TestBlobServiceClient:
    """Tests for blob_service_client function."""

    def test_blob_service_client_no_connection_string(self):
        """Test blob_service_client raises error when connection string not set."""
        from services.storage import blob_service_client

        with patch('services.storage.AZURE_CONN_STR', None):
            with pytest.raises(RuntimeError, match="AZURE_STORAGE_CONNECTION_STRING not set"):
                blob_service_client()

    def test_blob_service_client_with_blob_endpoint(self):
        """Test blob_service_client with BlobEndpoint in connection string."""
        from services.storage import blob_service_client

        conn_str = "AccountName=test;AccountKey=testkey;BlobEndpoint=https://test.blob.core.windows.net"

        with patch('services.storage.AZURE_CONN_STR', conn_str):
            with patch('services.storage.BlobServiceClient') as mock_client:
                blob_service_client()
                mock_client.assert_called_once()


class TestUploadAndSas:
    """Tests for upload_and_sas function."""

    def test_upload_and_sas_no_connection_string(self):
        """Test upload_and_sas raises error when connection string not set."""
        from services.storage import upload_and_sas

        with patch('services.storage.AZURE_CONN_STR', None):
            with pytest.raises(RuntimeError, match="AZURE_STORAGE_CONNECTION_STRING not set"):
                upload_and_sas("container", "blob.docx", b"data")
