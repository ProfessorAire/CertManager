param(
  [Parameter(Mandatory = $true, HelpMessage = "Root directory where the CA certificate will be stored.")]
  [string]$TargetDirectory,
  
  [Parameter(Mandatory = $true, HelpMessage = "Country code (e.g., US).")]
  [string]$Country,
  
  [Parameter(Mandatory = $true, HelpMessage = "Organization/unit name.")]
  [string]$OrgUnit,
  
  [Parameter(Mandatory = $true, HelpMessage = "CA certificate name (without extension).")]
  [string]$CaCertName
)

# Validate that OpenSSL exists
if (-Not (Get-Command "openssl" -ErrorAction SilentlyContinue))
{
  Write-Error "OpenSSL is not installed or not found in PATH."
  return
}

Write-Host "Creating Root CA certificate for organization: $OrgUnit" -ForegroundColor Cyan

# Step 1: Ensure the directory exists
$caDir = Join-Path $TargetDirectory $CaCertName
if (-Not (Test-Path $caDir))
{
  Write-Host "Creating directory: $caDir"
  New-Item -ItemType Directory -Path $caDir | Out-Null
}

# Step 2: Create the configuration file
$cnfPath = Join-Path $caDir "$CaCertName.cnf"
$cnContent = @"
[ req ]
default_bits = 4096
prompt = no
default_md = sha256
x509_extensions = v3_ca
distinguished_name = dn

[ dn ]
C  = $Country
O  = $OrgUnit
CN = $OrgUnit Private Root CA

[ v3_ca ]
subjectKeyIdentifier = hash
authorityKeyIdentifier = keyid:always
basicConstraints = critical, CA:true
keyUsage = critical, keyCertSign, cRLSign
"@

Write-Host "Creating configuration file: $cnfPath"
$cnContent | Out-File -FilePath $cnfPath -Encoding ascii

# Step 3: Generate the Root CA private key
$keyPath = Join-Path $caDir "$CaCertName.key"
Write-Host "Generating Root CA private key: $keyPath"
openssl genrsa -out $keyPath 4096

if (-Not (Test-Path $keyPath))
{
  Write-Error "Failed to generate Root CA private key."
  return
}

# Step 4: Create the Root CA certificate
$pemPath = Join-Path $caDir "$CaCertName.pem"
Write-Host "Creating Root CA certificate: $pemPath"
openssl req -x509 -new -key $keyPath -sha256 -days 3650 -out $pemPath -config $cnfPath

if (-Not (Test-Path $pemPath))
{
  Write-Error "Failed to create Root CA certificate."
  return
}

# Step 5: Verify the Root CA certificate
Write-Host "`nVerifying Root CA certificate..." -ForegroundColor Yellow
$verifyOutput = openssl x509 -in $pemPath -noout -text | Select-String -Pattern "Basic Constraints" -Context 0,3

if ($verifyOutput)
{
  Write-Host "Verification output:" -ForegroundColor Green
  $verifyOutput | ForEach-Object { Write-Host $_ }
  
  # Check for expected values
  $verifyText = $verifyOutput | Out-String
  if ($verifyText -match "CA:TRUE" -and $verifyText -match "Certificate Sign, CRL Sign")
  {
    Write-Host "`nRoot CA certificate verified successfully!" -ForegroundColor Green
  }
  else
  {
    Write-Warning "Root CA certificate may not have expected constraints."
  }
}
else
{
  Write-Warning "Could not verify Basic Constraints in certificate."
}

# Step 6: Create the DER format certificate
$derPath = Join-Path $caDir "$CaCertName.der"
Write-Host "`nCreating DER format certificate: $derPath"
openssl x509 -in $pemPath -outform der -out $derPath

if (Test-Path $derPath)
{
  Write-Host "DER certificate created successfully." -ForegroundColor Green
}
else
{
  Write-Error "Failed to create DER format certificate."
  return
}

# Step 7: Create the CER format certificate
$cerPath = Join-Path $caDir "$CaCertName.cer"
Write-Host "`nCreating CER format certificate: $cerPath"
openssl x509 -in $pemPath -outform der -out $cerPath
if (Test-Path $cerPath)
{
  Write-Host "CER certificate created successfully." -ForegroundColor Green
}
else
{
  Write-Error "Failed to create CER format certificate."
  return
}

Write-Host "`nRoot CA creation complete!" -ForegroundColor Cyan
Write-Host "  Private Key: $keyPath"
Write-Host "  Certificate (PEM): $pemPath"
Write-Host "  Certificate (DER): $derPath"
Write-Host "  Configuration: $cnfPath"
