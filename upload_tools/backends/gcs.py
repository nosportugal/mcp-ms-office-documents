import logging
from datetime import timedelta
from ..utils import get_content_type

logger = logging.getLogger(__name__)


def upload_to_gcs(file_object, file_name: str, gcscfg, signed_url_expires_in: int):
    """Upload a file to a GCS bucket and return a signed URL valid for configured duration."""

    if not gcscfg:
        logger.error("GCS configuration not provided")
        return None

    # Lazy import to avoid requiring google-cloud-storage unless GCS strategy is used
    try:
        from google.cloud import storage  # type: ignore
        from google.cloud.exceptions import GoogleCloudError  # type: ignore
    except Exception:
        logger.error("google-cloud-storage is not installed. Please add it to requirements and install.")
        return None

    content_type = get_content_type(file_name)

    try:
        # Use explicit service-account JSON when configured, otherwise fall
        # back to Application Default Credentials (ADC).  ADC automatically
        # picks up GKE Workload Identity / OIDC-federated service accounts.
        if gcscfg.credentials_path:
            storage_client = storage.Client.from_service_account_json(gcscfg.credentials_path)
        else:
            storage_client = storage.Client()

        bucket = storage_client.bucket(gcscfg.bucket)
        blob = bucket.blob(file_name)

        # Upload the file to GCS
        file_object.seek(0)  # Reset file pointer to beginning
        blob.upload_from_file(file_object, content_type=content_type)

        # Generate a signed URL valid for configured duration.
        # When using Workload Identity Federation (WIF) or other federated
        # credentials without a local private key, the default signing path
        # fails.  In that case we use the IAM signBlob API by passing the
        # service account email and an access token explicitly.
        signing_kwargs = {
            "version": "v4",
            "expiration": timedelta(seconds=signed_url_expires_in),
            "method": "GET",
        }

        if gcscfg.credentials_path:
            # Key-file credentials can sign locally — no extra args needed.
            url = blob.generate_signed_url(**signing_kwargs)
        else:
            # Federated / ADC credentials: delegate signing to IAM signBlob.
            import google.auth
            import google.auth.transport.requests

            credentials, _ = google.auth.default()
            credentials.refresh(google.auth.transport.requests.Request())

            signing_kwargs["service_account_email"] = credentials.service_account_email
            signing_kwargs["access_token"] = credentials.token
            url = blob.generate_signed_url(**signing_kwargs)

        return f"Link to created document to be shared with user in markdown format: {url} . Link is valid for {signed_url_expires_in} seconds."

    except GoogleCloudError as e:  # type: ignore[name-defined]
        logger.error(f"Google Cloud error: {e}")
        return None
    except Exception as e:
        logger.error(f"Error uploading to GCS: {e}")
        return None
