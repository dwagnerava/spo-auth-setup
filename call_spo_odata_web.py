import json
import os
import sys
import urllib.error
import urllib.request
from getpass import getpass


def get_optional_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        return ""
    return value.strip()


def load_dotenv(file_path: str = ".env") -> None:
    if not os.path.exists(file_path):
        return

    with open(file_path, "r", encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue

            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


def get_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        print(f"Missing required environment variable: {name}")
        sys.exit(1)
    return value


def get_access_token_with_certificate(tenant_id: str, client_id: str, sharepoint_host: str) -> str:
    cert_pfx_path = get_optional_env("AZURE_CERT_PFX_PATH")
    if not cert_pfx_path:
        return ""

    if not os.path.exists(cert_pfx_path):
        print(f"Certificate file not found: {cert_pfx_path}")
        sys.exit(1)

    try:
        import msal
        from cryptography.hazmat.primitives import hashes
        from cryptography.hazmat.primitives.serialization import Encoding, NoEncryption, PrivateFormat, pkcs12
    except ImportError:
        print("Missing dependencies for certificate auth. Install with: pip install msal cryptography")
        sys.exit(1)

    cert_data: bytes
    with open(cert_pfx_path, "rb") as cert_file:
        cert_data = cert_file.read()

    cert_password = get_optional_env("AZURE_CERT_PFX_PASSWORD")
    cert_password_bytes = cert_password.encode("utf-8") if cert_password else None

    try:
        key, cert, _ = pkcs12.load_key_and_certificates(cert_data, cert_password_bytes)
    except ValueError:
        if cert_password:
            print("PFX password from AZURE_CERT_PFX_PASSWORD appears invalid.")
            sys.exit(1)

        prompted_password = getpass("Enter PFX password: ")
        prompted_password_bytes = prompted_password.encode("utf-8") if prompted_password else None
        try:
            key, cert, _ = pkcs12.load_key_and_certificates(cert_data, prompted_password_bytes)
        except ValueError:
            print("Invalid PFX password or PKCS12 data.")
            sys.exit(1)

    if not key or not cert:
        print("Could not load private key and certificate from PFX file.")
        sys.exit(1)

    private_key_pem = key.private_bytes(
        encoding=Encoding.PEM,
        format=PrivateFormat.PKCS8,
        encryption_algorithm=NoEncryption(),
    ).decode("utf-8")

    thumbprint = cert.fingerprint(hashes.SHA1()).hex()

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential={
            "private_key": private_key_pem,
            "thumbprint": thumbprint,
        },
    )

    result = app.acquire_token_for_client(scopes=[f"https://{sharepoint_host}/.default"])
    access_token = result.get("access_token")
    if not access_token:
        print(f"Token request failed: {result}")
        sys.exit(1)

    return access_token


def get_access_token_with_secret(tenant_id: str, client_id: str, client_secret: str, sharepoint_host: str) -> str:
    try:
        import msal
    except ImportError:
        print("Missing dependency for secret auth. Install with: pip install msal")
        sys.exit(1)

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )

    result = app.acquire_token_for_client(scopes=[f"https://{sharepoint_host}/.default"])
    access_token = result.get("access_token")
    if not access_token:
        print(f"Token request failed: {result}")
        sys.exit(1)

    return access_token


def normalize_site_path(site_path: str) -> str:
    clean = site_path.strip()
    if not clean.startswith("/"):
        clean = "/" + clean
    return clean.rstrip("/")


def main() -> None:
    load_dotenv()

    tenant_id = get_env("AZURE_TENANT_ID")
    client_id = get_env("AZURE_CLIENT_ID")
    sharepoint_host = get_env("SHAREPOINT_HOST")
    site_path = normalize_site_path(get_env("SHAREPOINT_SITE_PATH"))

    cert_pfx_path = get_optional_env("AZURE_CERT_PFX_PATH")
    if cert_pfx_path:
        print(f"Auth mode: certificate ({cert_pfx_path})")
        access_token = get_access_token_with_certificate(tenant_id, client_id, sharepoint_host)
    else:
        print("Auth mode: client_secret")
        client_secret = get_env("AZURE_CLIENT_SECRET")
        access_token = get_access_token_with_secret(tenant_id, client_id, client_secret, sharepoint_host)

    endpoint = f"https://{sharepoint_host}{site_path}/_api/web"
    request = urllib.request.Request(endpoint, method="GET")
    request.add_header("Authorization", f"Bearer {access_token}")
    request.add_header("Accept", "application/json;odata=nometadata")

    try:
        with urllib.request.urlopen(request) as response:
            raw = response.read().decode("utf-8")
            print(f"GET {endpoint}")
            print(f"HTTP {response.status}")
            try:
                print(json.dumps(json.loads(raw), indent=2))
            except json.JSONDecodeError:
                print(raw)
    except urllib.error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        print(f"SharePoint call failed ({exc.code}): {error_body}")
        sys.exit(1)


if __name__ == "__main__":
    main()
