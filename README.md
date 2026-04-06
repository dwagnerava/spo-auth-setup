# SPO Auth Setup

This repository helps you validate app-only authentication to SharePoint Online and run a simple OData call to `/_api/web`.

It includes:
- `create_spo_app_cert.ps1`: creates a self-signed certificate and exports `.cer` + `.pfx`
- `call_spo_odata_web.py`: acquires an Entra ID token (certificate or client secret) and calls SharePoint REST

## Prerequisites

- Windows with PowerShell 5.1+ or PowerShell 7+
- Python 3.10+
- Access to Microsoft Entra admin center and SharePoint admin center
- Permission to create or update an App Registration

## 1) Install Python dependencies

From the repo root:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 2) Create a certificate for app auth (optional but recommended)

Run:

```powershell
.\create_spo_app_cert.ps1
```

Useful options:

```powershell
# Custom subject and validity
.\create_spo_app_cert.ps1 -CertSubject "CN=spo-auth-setup" -ValidYears 3

# Non-interactive/no password (local test only)
.\create_spo_app_cert.ps1 -NoPasswordPrompt
```

Outputs are written to `cert-output`:
- `<name>.cer`: upload to App Registration certificates
- `<name>.pfx`: used locally by the Python script

The script prints the certificate thumbprint and file paths after creation.

## 3) Configure App Registration (Entra ID)

1. Create or open an App Registration.
2. Copy:
- Application (client) ID
- Directory (tenant) ID
3. Add API permissions:
- `SharePoint` -> `Application permissions` -> `Sites.Read.All` (or broader if needed)
4. Grant admin consent for your tenant.
5. If using certificate auth:
- Go to `Certificates & secrets` -> `Certificates`
- Upload the generated `.cer` file
6. If using secret auth instead:
- Create a client secret and save its value

## 4) Grant site-level app access in SharePoint (if required)

If your tenant uses site-scoped app permissions, grant access to the target site using your preferred admin process (for example with PnP PowerShell `Grant-PnPAzureADAppSitePermission`).

At minimum, verify the app can access the site specified by `SHAREPOINT_SITE_PATH`.

## 5) Create your `.env` file

Create a `.env` file in the repo root.

### Certificate auth example

```dotenv
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
SHAREPOINT_HOST=contoso.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/YourSite
AZURE_CERT_PFX_PATH=cert-output/spo-auth-setup.pfx
AZURE_CERT_PFX_PASSWORD=your-pfx-password
```

### Client secret auth example

```dotenv
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=your-client-secret
SHAREPOINT_HOST=contoso.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/YourSite
```

Notes:
- If `AZURE_CERT_PFX_PATH` is present, the Python script uses certificate auth.
- If `AZURE_CERT_PFX_PATH` is not set, the script uses client secret auth.
- If the PFX password is omitted from `.env`, the script prompts you.

## 6) Run the SharePoint OData call

```powershell
python .\call_spo_odata_web.py
```

Expected success output:
- Auth mode line
- `GET https://<tenant>.sharepoint.com/sites/<site>/_api/web`
- `HTTP 200`
- JSON response body

## Troubleshooting

- `Missing required environment variable`:
  - Check `.env` key names and values.
- `Token request failed`:
  - Verify tenant ID/client ID, API permissions, and admin consent.
- `Certificate file not found`:
  - Confirm `AZURE_CERT_PFX_PATH` is correct.
- `Invalid PFX password or PKCS12 data`:
  - Re-enter password or recreate certificate.
- SharePoint `401/403`:
  - Confirm SharePoint permissions and any site-scoped app grants.
- Module errors:
  - Re-run `pip install -r requirements.txt` in the active venv.

## Repository notes

- `.gitignore` excludes local secrets (`.env`) and generated cert artifacts (`cert-output/`, `.pfx`, `.cer`, etc.).
- Keep `.pfx` and secrets out of source control.
