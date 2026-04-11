"""AWS S3 upload backend.

Supports two credential modes:

1. **Explicit credentials** — ``AWS_ACCESS_KEY``, ``AWS_SECRET_ACCESS_KEY``,
   and ``AWS_REGION`` are set as environment variables. The S3 client is
   created with these values directly.

2. **AWS default credential chain** — When the above env vars are *not* set,
   ``boto3`` automatically discovers credentials from the standard chain:

   - Environment variables (``AWS_ACCESS_KEY_ID``, etc.)
   - Shared credential / config files (``~/.aws/credentials``)
   - AWS SSO sessions (``aws sso login``) — for local development
   - ECS container credentials
   - **IRSA (IAM Roles for Service Accounts)** — for AWS EKS
   - EC2 instance metadata (IMDSv2)

   This mode requires no credential env vars at all; only ``S3_BUCKET`` is
   needed. Region is resolved automatically from the environment, AWS config,
   or instance metadata.

See Also:
    https://boto3.amazonaws.com/v1/documentation/api/latest/guide/credentials.html
"""

import logging
from ..utils import get_content_type

logger = logging.getLogger(__name__)


def _create_s3_client(s3cfg):
    """Create a boto3 S3 client using either explicit or default credentials.

    Args:
        s3cfg: An ``S3Settings`` instance from config.

    Returns:
        A boto3 S3 client ready for use.

    When ``s3cfg.use_explicit_credentials`` is True, the client is created
    with the provided access key, secret key, region, and a region-specific
    endpoint URL.

    When explicit credentials are not provided, boto3's default credential
    chain is used. This automatically supports:

    - IRSA on EKS (via ``AWS_WEB_IDENTITY_TOKEN_FILE`` and
      ``AWS_ROLE_ARN`` injected by the EKS pod identity webhook)
    - AWS SSO sessions (after running ``aws sso login``)
    - Instance profiles, environment variables, config files, etc.
    """
    import boto3  # type: ignore

    if s3cfg.use_explicit_credentials:
        logger.info("Creating S3 client with explicit credentials (region: %s)", s3cfg.region)
        return boto3.client(
            's3',
            region_name=s3cfg.region,
            aws_access_key_id=s3cfg.access_key,
            aws_secret_access_key=s3cfg.secret_key,
            endpoint_url=f'https://s3.{s3cfg.region}.amazonaws.com',
        )

    # Use the default credential chain (IRSA, SSO, instance profile, etc.)
    logger.info(
        "Creating S3 client using default credential chain (region: %s)",
        s3cfg.region or "auto-detected",
    )
    client_kwargs = {"region_name": s3cfg.region} if s3cfg.region else {}
    return boto3.client('s3', **client_kwargs)


def upload_to_s3(file_object, file_name: str, s3cfg, signed_url_expires_in: int):
    """Upload a file to S3 and return a pre-signed download URL.

    Args:
        file_object: A file-like object (must support ``seek`` and ``read``).
        file_name: The S3 object key (destination path/name in the bucket).
        s3cfg: An ``S3Settings`` instance with bucket and optional credentials.
        signed_url_expires_in: TTL in seconds for the pre-signed download URL.

    Returns:
        A string containing the pre-signed URL and expiry info, or ``None``
        on failure.
    """
    if not s3cfg:
        logger.error("S3 configuration not provided")
        return None

    # Lazy import to avoid requiring boto3/botocore unless S3 strategy is used
    try:
        from botocore.exceptions import NoCredentialsError, ClientError  # type: ignore
    except Exception:
        logger.error("boto3/botocore are not installed. Please add them to requirements and install.")
        return None

    content_type = get_content_type(file_name)

    try:
        s3_client = _create_s3_client(s3cfg)

        # Upload the file to S3
        file_object.seek(0)
        s3_client.upload_fileobj(
            Fileobj=file_object,
            Bucket=s3cfg.bucket,
            Key=file_name,
            ExtraArgs={'ContentType': content_type},
        )

        # Generate a pre-signed URL valid for configured duration
        url = s3_client.generate_presigned_url(
            'get_object',
            Params={'Bucket': s3cfg.bucket, 'Key': file_name},
            ExpiresIn=signed_url_expires_in,
        )

        return f"Link to created document to be shared with user in markdown format: {url} . Link is valid for {signed_url_expires_in} seconds."

    except FileNotFoundError:
        logger.error(f"The file {file_object} was not found.")
        return None
    except NoCredentialsError:
        logger.error(
            "AWS credentials are not available. When running without explicit "
            "AWS_ACCESS_KEY / AWS_SECRET_ACCESS_KEY, ensure credentials are "
            "available via the AWS default credential chain: IRSA (EKS), "
            "aws sso login (local), instance profile, or ~/.aws/credentials."
        )
        return None
    except ClientError as e:
        logger.error(f"Client error: {e}")
        return None
    except Exception as e:
        logger.error(f"Unexpected error uploading to S3: {e}")
        return None
