# tests/test_config.py
"""Tests for configuration module."""

import os
import pytest


class TestConfig:
    """Test configuration values and environment variables."""

    def test_fac_base_default(self):
        """Test FAC_BASE has correct default value."""
        from config import FAC_BASE
        # Default should be set when env var is not present
        assert FAC_BASE is not None
        assert "fac.gov" in FAC_BASE

    def test_azure_container_default(self):
        """Test AZURE_CONTAINER has correct default value."""
        from config import AZURE_CONTAINER
        assert AZURE_CONTAINER == "mdl-output"

    def test_azurite_sas_version_default(self):
        """Test AZURITE_SAS_VERSION has correct default value."""
        from config import AZURITE_SAS_VERSION
        assert AZURITE_SAS_VERSION == "2021-08-06"

    def test_local_save_dir_default(self):
        """Test LOCAL_SAVE_DIR has correct default value."""
        from config import LOCAL_SAVE_DIR
        assert LOCAL_SAVE_DIR == "./_out"

    def test_public_base_url_default(self):
        """Test PUBLIC_BASE_URL has correct default value."""
        from config import PUBLIC_BASE_URL
        assert PUBLIC_BASE_URL == "http://localhost:8000"

    def test_treasury_programs_contains_slfrf(self):
        """Test TREASURY_PROGRAMS contains SLFRF program."""
        from config import TREASURY_PROGRAMS
        assert "21.027" in TREASURY_PROGRAMS
        assert "SLFRF" in TREASURY_PROGRAMS["21.027"]

    def test_treasury_programs_contains_era(self):
        """Test TREASURY_PROGRAMS contains ERA program."""
        from config import TREASURY_PROGRAMS
        assert "21.023" in TREASURY_PROGRAMS
        assert "ERA" in TREASURY_PROGRAMS["21.023"]

    def test_treasury_programs_contains_haf(self):
        """Test TREASURY_PROGRAMS contains HAF program."""
        from config import TREASURY_PROGRAMS
        assert "21.026" in TREASURY_PROGRAMS
        assert "HAF" in TREASURY_PROGRAMS["21.026"]

    def test_treasury_programs_contains_cpf(self):
        """Test TREASURY_PROGRAMS contains CPF program."""
        from config import TREASURY_PROGRAMS
        assert "21.029" in TREASURY_PROGRAMS
        assert "CPF" in TREASURY_PROGRAMS["21.029"]

    def test_treasury_programs_contains_ssbci(self):
        """Test TREASURY_PROGRAMS contains SSBCI program."""
        from config import TREASURY_PROGRAMS
        assert "21.031" in TREASURY_PROGRAMS
        assert "SSBCI" in TREASURY_PROGRAMS["21.031"]

    def test_treasury_programs_contains_latcf(self):
        """Test TREASURY_PROGRAMS contains LATCF program."""
        from config import TREASURY_PROGRAMS
        assert "21.032" in TREASURY_PROGRAMS
        assert "LATCF" in TREASURY_PROGRAMS["21.032"]

    def test_lowercase_words_contains_common_words(self):
        """Test LOWERCASE_WORDS contains expected connector words."""
        from config import LOWERCASE_WORDS
        expected = {"and", "of", "the", "for", "to", "in", "on", "at", "by", "with", "from"}
        assert LOWERCASE_WORDS == expected
