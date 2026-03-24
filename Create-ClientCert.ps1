param(
  [Parameter(Mandatory = $true, HelpMessage = "Root directory where certificates are stored.")]
  [string]$TargetDirectory,
        
  [Parameter(Mandatory = $true, HelpMessage = "Fully Qualified Domain Name (FQDN) for the certificate. Such as server.example.com")]
  [string]$Fqdn,
        
  [Parameter(Mandatory = $true, HelpMessage = "Array of IP addresses to include in the certificate SAN (e.g., '192.168.1.1,10.0.0.1').")]
  [string[]]$IpAddress,

  [Parameter(HelpMessage = "Password for the .pfx file. Leave empty for no password.")]
  [string]$CertPassword = "",
  
  [Parameter(Mandatory = $true, HelpMessage = "Country code (e.g., US).")]
  [string]$Country,
  
  [Parameter(Mandatory = $true, HelpMessage = "Organization/unit name.")]
  [string]$OrgUnit,
  
  [Parameter(Mandatory = $true, HelpMessage = "CA certificate name (without extension).")]
  [string]$CaCertName,
  
  [Parameter(Mandatory = $true, HelpMessage = "Path to the CA certificate directory.")]
  [string]$CaCertificatePath,
  
  [Parameter(HelpMessage = "Password for the CA private key.")]
  [string]$CaKeyPassword = "",
  
  [Parameter(HelpMessage = "Use legacy encryption (3DES, SHA1) for older devices like Crestron Series 3.")]
  [switch]$UseLegacyEncryption
)

# Set paths based on parameters
$caCertificate = Join-Path $CaCertificatePath "$CaCertName.pem"
$caCertKey = Join-Path $CaCertificatePath "$CaCertName.key"
$certsRootDirectory = $TargetDirectory
    
# Validate that OpenSSL exists
if (-Not (Get-Command "openssl" -ErrorAction SilentlyContinue))
{
  Write-Error "OpenSSL is not installed or not found in PATH."
  return
}

# Validate that the CA certificate file exists
if (-Not (Test-Path $caCertificate))
{
  Write-Error "CA certificate file not found: $caCertificate"
  return
}

# Validate that fqdn is a valid hostname or fully qualified domain name
if ($Fqdn -notmatch '^[a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$')
{
  Write-Error "Invalid FQDN format: $Fqdn. Must be a valid hostname or fully qualified domain name (e.g., server or server.example.com)"
  return
}

# Parse and validate IP addresses (support comma-separated list)
$ipv4Regex = '^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'

foreach ($ip in $IpAddress)
{
  if ($ip -notmatch $ipv4Regex)
  {
    Write-Error "Invalid IP address format: $ip. Must be a valid IPv4 address (e.g., 192.168.1.50)"
    return
  }
}

if ($IpAddress.Count -eq 0)
{
  Write-Error "At least one valid IP address must be provided."
  return
}

if ($IpAddress.Count -gt 1)
{
  Write-Host "Creating certificate for: $Fqdn with IPs: $($IpAddress -join ', ')"
}
else
{
  Write-Host "Creating certificate for: $Fqdn with IP: $($IpAddress[0])"
}
  
# Make sure a directory for the FQDN exists
$certDir = Join-Path $certsRootDirectory $Fqdn
if (-Not (Test-Path $certDir))
{
  New-Item -ItemType Directory -Path $certDir | Out-Null
}

# Check if a cnf file already exists for this FQDN in the certDir
if (Test-Path $certDir\$Fqdn.cnf)
{
  Remove-Item $certDir\$Fqdn.cnf -Force
}

  # Build the alt_names section with all IP addresses
  $altNamesSection = "DNS.1 = $Fqdn`n"
  for ($i = 0; $i -lt $IpAddress.Count; $i++)
  {
    $altNamesSection += "IP.$($i + 1) = $($IpAddress[$i])`n"
  }
  
  # Create OpenSSL config file with SANs
  $cnfContent = @"
[ req ]
default_bits = 2048
prompt = no
default_md = sha256
distinguished_name = dn
req_extensions = v3_req

[ dn ]
C  = $Country
O  = $OrgUnit
CN = $Fqdn

[ v3_req ]
subjectAltName = @alt_names

[ alt_names ]
$altNamesSection
"@

  $cnfPath = "$certDir\$Fqdn.cnf"
  $cnfContent | Out-File -FilePath $cnfPath -Encoding ascii

# Generate a private key for the server
Write-Host "Generating private key for $Fqdn"
openssl genrsa -out "$certDir\$Fqdn.key" 2048

# For Crestron Series 3 compatibility, ensure the key is in traditional PKCS#1 format
# (BEGIN RSA PRIVATE KEY instead of BEGIN PRIVATE KEY)
if ($UseLegacyEncryption)
{
  Write-Host "Converting private key to traditional PKCS#1 format for legacy compatibility..."
  # Create a temporary file for the conversion
  $tempKey = "$certDir\$Fqdn.key.tmp"
  $rsaKeyResult = openssl rsa -in "$certDir\$Fqdn.key" -out $tempKey -traditional 2>&1
  
  if ($LASTEXITCODE -eq 0)
  {
    # Replace the original key with the converted one
    Move-Item -Path $tempKey -Destination "$certDir\$Fqdn.key" -Force
    Write-Host "Private key converted to PKCS#1 format successfully."
  }
  else
  {
    Write-Warning "Failed to convert key to PKCS#1 format: $rsaKeyResult"
    # Clean up temp file if it exists
    if (Test-Path $tempKey) { Remove-Item $tempKey -Force }
  }
}

# Create a CSR for the server
Write-Host "Creating CSR for $Fqdn"
openssl req -new -key "$certDir\$Fqdn.key" -out "$certDir\$Fqdn.csr" -config "$certDir\$Fqdn.cnf"

# Sign the CSR with the CA to create the client certificate
Write-Host "Signing CSR to create certificate for $Fqdn"
if ([string]::IsNullOrWhiteSpace($CaKeyPassword))
{
  openssl x509 -req -in "$certDir\$Fqdn.csr" -CA $caCertificate -CAkey $caCertKey -CAcreateserial -out "$certDir\$Fqdn.crt" -days 397 -sha256 -extfile "$certDir\$Fqdn.cnf" -extensions v3_req
}
else
{
  openssl x509 -req -in "$certDir\$Fqdn.csr" -CA $caCertificate -CAkey $caCertKey -CAcreateserial -out "$certDir\$Fqdn.crt" -days 397 -sha256 -extfile "$certDir\$Fqdn.cnf" -extensions v3_req -passin "pass:$CaKeyPassword"
}

# Create a file with the full certificate chain
Write-Host "Creating fullchain certificate for $Fqdn"
Get-Content "$certDir\$Fqdn.crt", $caCertificate | Set-Content -Path "$certDir\$Fqdn-fullchain.crt" -Encoding ascii

# Verify the Certificate Signature, checking that it outputs OK as a result and erroring otherwise.
Write-Host "Verifying certificate for $Fqdn"
$verifyOutput = openssl verify -CAfile $caCertificate "$certDir\$Fqdn.crt"
if ($verifyOutput -notmatch "$Fqdn.crt: OK")
{
  Write-Error "Certificate verification failed for $Fqdn.crt. Output: $verifyOutput"
  return
}
else {
  Write-Host "Certificate verification succeeded for $Fqdn.crt."
}

# Inspect the certificate details
$details = openssl x509 -in "$certDir\$Fqdn.crt" -text -noout -subject -issuer -dates -ext subjectAltName
# Verify the certificate details, ensuring the expected values are present (detailed checks)
# Normalize details to a single string for reliable regex matching across lines
$detailsText = $details -join "`n"

$failed = $false

# Check Subject CN
if ($detailsText -notmatch "Subject:.*CN ?= ?$Fqdn")
{
  Write-Error "Subject CN mismatch. Expected CN = '$Fqdn'."
  $failed = $true
}

# Check Issuer CN (expected CA name format used previously)
$expectedIssuer = "$OrgUnit Private Root CA"
if ($detailsText -notmatch "Issuer:.*CN ?= ?$expectedIssuer")
{
  Write-Error "Issuer CN mismatch. Expected Issuer CN = '$expectedIssuer'."
  $failed = $true
}

# Check validity dates
if ($detailsText -notmatch "Not Before")
{
  Write-Error "Missing 'Not Before' date in certificate details."
  $failed = $true
}
if ($detailsText -notmatch "Not After")
{
  Write-Error "Missing 'Not After' date in certificate details."
  $failed = $true
}

# Check Subject Alternative Names for DNS and IP
# OpenSSL formats SANs like: "DNS:example.com, IP Address:192.168.0.1, IP Address:192.168.1.1"
# Verify DNS entry exists
if ($detailsText -notmatch "(?s)X509v3 Subject Alternative Name:.*DNS:\s*$Fqdn")
{
  Write-Error "SubjectAltName mismatch. Expected DNS:$Fqdn in certificate."
  $failed = $true
}

# Verify each IP address exists in the SANs
foreach ($ip in $IpAddress)
{
  $escapedIp = [regex]::Escape($ip)
  if ($detailsText -notmatch "IP Address:\s*$escapedIp")
  {
    Write-Error "SubjectAltName mismatch. Expected IP Address:$ip in certificate."
    $failed = $true
  }
}

if ($failed)
{
  Write-Error "Certificate details verification failed for $Fqdn. Full details:`n$detailsText"
  return
}
else
{
  Write-Host "Certificate details verified successfully for $Fqdn."
}


# Create a pfx file for easier import into browsers/OS
Write-Host "Creating PFX file for $Fqdn"
$pfxPath = "$certDir\$Fqdn.pfx"

if ($UseLegacyEncryption)
{
  Write-Host "Using legacy encryption (3DES, SHA1) for compatibility with older devices"
  openssl pkcs12 -export -out $pfxPath -inkey "$certDir\$Fqdn.key" -in "$certDir\$Fqdn.crt" -passout pass:$CertPassword -keypbe PBE-SHA1-3DES -certpbe PBE-SHA1-3DES -macalg sha1
}
else
{
  openssl pkcs12 -export -out $pfxPath -inkey "$certDir\$Fqdn.key" -in "$certDir\$Fqdn.crt" -passout pass:$CertPassword
}

# Verify PFX file was created
if (Test-Path $pfxPath)
{
  Write-Host "PFX file created successfully at: $pfxPath"
}
else
{
  Write-Error "Failed to create PFX file at: $pfxPath"
  return
}

# Create PEM and CER versions of the certificate for compatibility
$pemPath = Join-Path $certDir "$Fqdn.pem"
$cerPath = Join-Path $certDir "$Fqdn.cer"

Write-Host "Creating PEM and CER versions for $Fqdn"

try
{
  # Extract certificate (PEM) from the PFX (no private key)
  openssl pkcs12 -in $pfxPath -nokeys -out $pemPath -passin "pass:$CertPassword"

  # Convert PEM to DER-formatted .cer
  openssl x509 -in $pemPath -outform DER -out $cerPath

  Write-Host "Created PEM: $pemPath"
  Write-Host "Created CER: $cerPath"
}
catch
{
  Write-Error "Failed to create PEM/CER from $pfxPath`: $($_.Exception.Message)"
}

Write-Host "`nCertificate creation process completed for $Fqdn." -ForegroundColor Green