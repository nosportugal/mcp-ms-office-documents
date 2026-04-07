"""Tests for AWS S3 upload backend configuration and client creation.

Validates three credential scenarios:
1. Explicit credentials (AWS_ACCESS_KEY, AWS_SECRET_ACCESS_KEY, AWS_REGION, S3_BUCKET)
2. Default credential chain with no creds (IRSA on EKS, SSO locally)
3. Default credential chain with optional region override

Also verifies validation rules:
- Partial credentials (key without secret or vice versa) are rejected
- Missing S3_BUCKET is rejected
- Explicit credentials without AWS_REGION are rejected
- Empty string env vars normalize to the default credential chain
"""

import sys
import os
from pathlib import Path
from unittest.mock import patch, MagicMock

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Env vars used by S3Settings via Config.from_env()
S3_ENV_KEYS = [
    "UPLOAD_STRATEGY",
    "AWS_ACCESS_KEY",
    "AWS_SECRET_ACCESS_KEY",
    "AWS_REGION",
    "S3_BUCKET",
]


def _build_env(**overrides):
    """Return a minimal S3 environment dict with given overrides.

    Keys not in *overrides* are removed from the environment so that
    previous test state doesn't leak.
    """
    env = {k: os.environ.get(k) for k in S3_ENV_KEYS}  # snapshot
    # Clear all S3-related vars
    clean = {k: v for k, v in os.environ.items() if k not in S3_ENV_KEYS}
    clean["UPLOAD_STRATEGY"] = "S3"
    clean.update(overrides)
    return clean


def _get_fresh_config(env: dict):
    """Return a fresh Config from *env*, bypassing the singleton cache."""
    import config as cfg_mod
    cfg_mod._CONFIG = None
    with patch.dict(os.environ, env, clear=True):
        return cfg_mod.get_config()


# ---------------------------------------------------------------------------
# S3Settings validation tests
# ---------------------------------------------------------------------------

class TestS3SettingsValidation:
    """Test S3Settings Pydantic model validation rules."""

    def test_explicit_credentials_all_set(self):
        """All four env vars provided — classic / original behaviour."""
        env = _build_env(
            AWS_ACCESS_KEY="<TEST_AWS_ACCESS_KEY>",
            AWS_SECRET_ACCESS_KEY="<TEST_AWS_SECRET_KEY>",
            AWS_REGION="us-east-1",
            S3_BUCKET="my-test-bucket",
        )
        cfg = _get_fresh_config(env)

        assert cfg.storage.s3 is not None
        assert cfg.storage.s3.use_explicit_credentials is True
        assert cfg.storage.s3.access_key == "<TEST_AWS_ACCESS_KEY>"
        assert cfg.storage.s3.secret_key == "<TEST_AWS_SECRET_KEY>"
        assert cfg.storage.s3.region == "us-east-1"
        assert cfg.storage.s3.bucket == "my-test-bucket"

    def test_default_chain_no_credentials(self):
        """Only S3_BUCKET set — should use default credential chain (IRSA / SSO)."""
        env = _build_env(S3_BUCKET="irsa-bucket")
        cfg = _get_fresh_config(env)

        assert cfg.storage.s3 is not None
        assert cfg.storage.s3.use_explicit_credentials is False
        assert cfg.storage.s3.access_key is None
        assert cfg.storage.s3.secret_key is None
        assert cfg.storage.s3.region is None
        assert cfg.storage.s3.bucket == "irsa-bucket"

    def test_default_chain_with_region(self):
        """S3_BUCKET + AWS_REGION set, no credentials — region is passed through."""
        env = _build_env(AWS_REGION="eu-west-1", S3_BUCKET="regional-bucket")
        cfg = _get_fresh_config(env)

        assert cfg.storage.s3.use_explicit_credentials is False
        assert cfg.storage.s3.region == "eu-west-1"
        assert cfg.storage.s3.bucket == "regional-bucket"

    def test_empty_strings_normalize_to_default_chain(self):
        """Empty-string env vars should behave as if unset."""
        env = _build_env(
            AWS_ACCESS_KEY="",
            AWS_SECRET_ACCESS_KEY="",
            AWS_REGION="",
            S3_BUCKET="empty-strings-bucket",
        )
        cfg = _get_fresh_config(env)

        assert cfg.storage.s3.use_explicit_credentials is False
        assert cfg.storage.s3.access_key is None
        assert cfg.storage.s3.secret_key is None
        assert cfg.storage.s3.region is None

    def test_partial_credentials_rejected(self):
        """Only one of access_key / secret_key set — should raise ValueError."""
        env = _build_env(
            AWS_ACCESS_KEY="AKID",
            AWS_REGION="us-east-1",
            S3_BUCKET="partial-bucket",
        )
        with pytest.raises(ValueError, match="both be set or both be omitted"):
            _get_fresh_config(env)

    def test_missing_bucket_rejected(self):
        """Empty S3_BUCKET should raise ValueError."""
        env = _build_env(S3_BUCKET="")
        with pytest.raises(ValueError, match="S3_BUCKET"):
            _get_fresh_config(env)

    def test_explicit_credentials_without_region_rejected(self):
        """Explicit credentials without AWS_REGION should raise ValueError."""
        env = _build_env(
            AWS_ACCESS_KEY="AKID",
            AWS_SECRET_ACCESS_KEY="secret",
            S3_BUCKET="no-region-bucket",
        )
        with pytest.raises(ValueError, match="AWS_REGION is required"):
            _get_fresh_config(env)


# ---------------------------------------------------------------------------
# S3 client creation tests
# ---------------------------------------------------------------------------

class TestS3ClientCreation:
    """Test _create_s3_client builds the boto3 client correctly."""

    def test_client_with_explicit_credentials(self):
        """Explicit credentials → boto3.client called with access key, secret, region, endpoint."""
        env = _build_env(
            AWS_ACCESS_KEY="AKID",
            AWS_SECRET_ACCESS_KEY="secret",
            AWS_REGION="us-west-2",
            S3_BUCKET="explicit-bucket",
        )
        cfg = _get_fresh_config(env)

        mock_client = MagicMock()
        with patch("boto3.client", return_value=mock_client) as patched:
            from upload_tools.backends.s3 import _create_s3_client
            result = _create_s3_client(cfg.storage.s3)

            patched.assert_called_once_with(
                's3',
                region_name="us-west-2",
                aws_access_key_id="AKID",
                aws_secret_access_key="secret",
                endpoint_url="https://s3.us-west-2.amazonaws.com",
            )
            assert result is mock_client

    def test_client_with_default_chain_no_region(self):
        """Default chain, no region → boto3.client called with no extra kwargs."""
        env = _build_env(S3_BUCKET="chain-bucket")
        cfg = _get_fresh_config(env)

        mock_client = MagicMock()
        with patch("boto3.client", return_value=mock_client) as patched:
            from upload_tools.backends.s3 import _create_s3_client
            result = _create_s3_client(cfg.storage.s3)

            patched.assert_called_once_with('s3')
            assert result is mock_client

    def test_client_with_default_chain_and_region(self):
        """Default chain + region → boto3.client called with region_name only."""
        env = _build_env(AWS_REGION="ap-southeast-1", S3_BUCKET="region-chain-bucket")
        cfg = _get_fresh_config(env)

        mock_client = MagicMock()
        with patch("boto3.client", return_value=mock_client) as patched:
            from upload_tools.backends.s3 import _create_s3_client
            result = _create_s3_client(cfg.storage.s3)

            patched.assert_called_once_with('s3', region_name="ap-southeast-1")
            assert result is mock_client


# ---------------------------------------------------------------------------
# upload_to_s3 integration tests (mocked boto3)
# ---------------------------------------------------------------------------

class TestUploadToS3:
    """Test the upload_to_s3 function end-to-end with mocked boto3."""

    def test_upload_returns_presigned_url(self):
        """Successful upload should return a string with the presigned URL."""
        from io import BytesIO
        from upload_tools.backends.s3 import upload_to_s3

        env = _build_env(S3_BUCKET="upload-bucket")
        cfg = _get_fresh_config(env)

        mock_client = MagicMock()
        mock_client.generate_presigned_url.return_value = "https://s3.amazonaws.com/upload-bucket/test.docx?signed"

        with patch("upload_tools.backends.s3._create_s3_client", return_value=mock_client):
            file_obj = BytesIO(b"test content")
            result = upload_to_s3(file_obj, "test.docx", cfg.storage.s3, 3600)

        assert result is not None
        assert "https://s3.amazonaws.com/upload-bucket/test.docx?signed" in result
        assert "3600 seconds" in result

        mock_client.upload_fileobj.assert_called_once()
        mock_client.generate_presigned_url.assert_called_once_with(
            'get_object',
            Params={'Bucket': 'upload-bucket', 'Key': 'test.docx'},
            ExpiresIn=3600,
        )

    def test_upload_returns_none_when_no_config(self):
        """upload_to_s3 should return None when s3cfg is None."""
        from upload_tools.backends.s3 import upload_to_s3
        from io import BytesIO

        result = upload_to_s3(BytesIO(b"data"), "file.docx", None, 3600)
        assert result is None

    def test_upload_returns_none_on_no_credentials_error(self):
        """upload_to_s3 should return None and log when credentials are missing."""
        from io import BytesIO
        from botocore.exceptions import NoCredentialsError
        from upload_tools.backends.s3 import upload_to_s3

        env = _build_env(S3_BUCKET="no-creds-bucket")
        cfg = _get_fresh_config(env)

        mock_client = MagicMock()
        mock_client.upload_fileobj.side_effect = NoCredentialsError()

        with patch("upload_tools.backends.s3._create_s3_client", return_value=mock_client):
            result = upload_to_s3(BytesIO(b"data"), "file.docx", cfg.storage.s3, 3600)

        assert result is None

    def test_upload_returns_none_on_client_error(self):
        """upload_to_s3 should return None on ClientError."""
        from io import BytesIO
        from botocore.exceptions import ClientError
        from upload_tools.backends.s3 import upload_to_s3

        env = _build_env(S3_BUCKET="error-bucket")
        cfg = _get_fresh_config(env)

        mock_client = MagicMock()
        mock_client.upload_fileobj.side_effect = ClientError(
            {"Error": {"Code": "AccessDenied", "Message": "Forbidden"}},
            "PutObject",
        )

        with patch("upload_tools.backends.s3._create_s3_client", return_value=mock_client):
            result = upload_to_s3(BytesIO(b"data"), "file.docx", cfg.storage.s3, 3600)

        assert result is None
