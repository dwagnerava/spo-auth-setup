param(
    [string]$CertSubject = "CN=spo-app-only",
    [int]$ValidYears = 2,
    [string]$OutputFolder = "cert-output",
    [switch]$NoPasswordPrompt
)

$ErrorActionPreference = "Stop"

$outDir = Join-Path -Path $PSScriptRoot -ChildPath $OutputFolder
if (-not (Test-Path -Path $outDir)) {
    New-Item -Path $outDir -ItemType Directory | Out-Null
}

$notAfter = (Get-Date).AddYears($ValidYears)

$certParams = @{
    Subject = $CertSubject
    CertStoreLocation = "Cert:\CurrentUser\My"
    KeyAlgorithm = "RSA"
    KeyLength = 2048
    HashAlgorithm = "SHA256"
    KeyExportPolicy = "Exportable"
    KeySpec = "Signature"
    Provider = "Microsoft Enhanced RSA and AES Cryptographic Provider"
    NotAfter = $notAfter
}

$cert = New-SelfSignedCertificate @certParams

if (-not $cert) {
    throw "Certificate creation failed."
}

$safeName = ($CertSubject -replace "^CN=", "" -replace "[^a-zA-Z0-9._-]", "_")
$cerPath = Join-Path -Path $outDir -ChildPath "$safeName.cer"
$pfxPath = Join-Path -Path $outDir -ChildPath "$safeName.pfx"

Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null

if ($NoPasswordPrompt) {
    # Only for local testing. Use a password-protected PFX for shared or CI environments.
    $pfxBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, "")
    [System.IO.File]::WriteAllBytes($pfxPath, $pfxBytes)
}
else {
    $securePassword = Read-Host -Prompt "Enter password for PFX file" -AsSecureString
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $securePassword | Out-Null
}

Write-Host ""
Write-Host "Certificate created successfully."
Write-Host "Thumbprint: $($cert.Thumbprint)"
Write-Host "Store path: Cert:\CurrentUser\My\$($cert.Thumbprint)"
Write-Host "Public cert (.cer): $cerPath"
Write-Host "Private key (.pfx): $pfxPath"
Write-Host ""
Write-Host "Next step: upload the .cer file to your app registration."
