# ManageCerts.ps1
# Script to manage certificates on various devices, including Crestron devices.

param(
  [Parameter(Mandatory = $false, HelpMessage = "Root directory where certificates are stored and managed.")]
  [string]$TargetDirectory = $PSScriptRoot,
  
  [Parameter(Mandatory = $false, HelpMessage = "FQDN or IP address of specific device(s) to manage. If not specified, all devices in configuration will be processed.")]
  [string[]]$Devices
)

# Ensure target directory exists
if (-not (Test-Path -Path $TargetDirectory))
{
  Write-Error "Target directory does not exist: $TargetDirectory"
  exit 1
}

# Configuration file path
$configPath = Join-Path $TargetDirectory "CertificateConfiguration.config"

# Function to create default configuration
function New-DefaultConfiguration
{
  param([string]$ConfigPath)
  
  Write-Host "`nNo configuration file found. Let's create one.`n" -ForegroundColor Yellow
  
  # Prompt for defaults
  $country = Read-Host "Enter country code (e.g., US)"
  if ([string]::IsNullOrWhiteSpace($country)) { $country = "US" }
  
  $orgUnit = Read-Host "Enter organization/unit name (e.g., MyOrg)"
  if ([string]::IsNullOrWhiteSpace($orgUnit)) { $orgUnit = "MyOrg" }
  
  $caCertName = Read-Host "Enter CA certificate filename (without extension, e.g., my-ca)"
  if ([string]::IsNullOrWhiteSpace($caCertName)) { $caCertName = "ca" }
  
  $certPassword = Read-Host "Enter default certificate password for Crestron devices (e.g., LogosRocks)"
  if ([string]::IsNullOrWhiteSpace($certPassword)) { $certPassword = "LogosRocks" }
  
  # Prompt for devices to manage
  $devices = Add-DevicesToConfiguration
  
  # Create configuration object
  $config = @{
    Country              = $country
    OrgUnit              = $orgUnit
    CACertificateName    = $caCertName
    CrestronCertPassword = $certPassword
    CertificatesToManage = $devices
  }
  
  # Save to file
  $config | ConvertTo-Json -Depth 10 | Set-Content -Path $ConfigPath -Encoding UTF8
  
  Write-Host "`nConfiguration saved to: $ConfigPath" -ForegroundColor Green
  Write-Host "You can manually edit this file to add/modify devices in the future.`n" -ForegroundColor Cyan
  
  return $config
}

# Function to prompt user to add devices
function Add-DevicesToConfiguration
{
  Write-Host "`nLet's add devices to manage. Press Enter on FQDN to finish.`n" -ForegroundColor Yellow
  $devices = @()
  
  do
  {
    $fqdn = Read-Host "Device FQDN or hostname (e.g., device.example.com or device)"
    if ([string]::IsNullOrWhiteSpace($fqdn)) { break }
    
    $ipInput = Read-Host "Device IP address(es) (comma-separated for multiple: 192.168.1.1,10.0.0.1)"
    # Parse IP addresses into array
    $ipAddresses = $ipInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    
    $connectionAddressInput = Read-Host "Connection address for device communication (leave blank to use '$fqdn')"
    
    $username = Read-Host "Device username"
    $password = Read-Host "Device password" -AsSecureString
    $passwordPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
      [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
    
    $updateType = Read-Host "Update type (Crestron3, Crestron4, CrestronTP60Series, CrestronTP70Series, SCP, TrueNAS, UniFi)"
    
    $deviceEntry = @{
      FQDN       = $fqdn
      IPAddress  = $ipAddresses
      Username   = $username
      Password   = $passwordPlain
      UpdateType = $updateType
    }
    
    # Only include ConnectionAddress if a value was provided
    if (-not [string]::IsNullOrWhiteSpace($connectionAddressInput))
    {
      $deviceEntry.ConnectionAddress = $connectionAddressInput
    }
    
    $devices += $deviceEntry
    
    Write-Host "Device added.`n" -ForegroundColor Green
  } while ($true)
  
  return $devices
}

# Load or create configuration
if (Test-Path -Path $configPath)
{
  Write-Host "Loading configuration from: $configPath"
  $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
  
  # Convert PSCustomObject to hashtable for easier access
  $certificatesToManage = @()
  foreach ($cert in $config.CertificatesToManage)
  {
    $certHash = @{
      FQDN       = $cert.FQDN
      IPAddress  = $cert.IPAddress
      Username   = $cert.Username
      Password   = $cert.Password
      UpdateType = $cert.UpdateType
    }
    
    # Include FileMappings if present
    if ($cert.PSObject.Properties['FileMappings'])
    {
      $certHash.FileMappings = $cert.FileMappings
    }
    
    # Include ApiKey if present (for TrueNAS)
    if ($cert.PSObject.Properties['ApiKey'])
    {
      $certHash.ApiKey = $cert.ApiKey
    }
    
    # Include ConnectionAddress if present
    if ($cert.PSObject.Properties['ConnectionAddress'])
    {
      $certHash.ConnectionAddress = $cert.ConnectionAddress
    }
    
    $certificatesToManage += $certHash
  }
  
  # Check if CertificatesToManage is empty and prompt to add devices
  if ($certificatesToManage.Count -eq 0)
  {
    Write-Host "`nNo devices found in configuration." -ForegroundColor Yellow
    $addDevices = Read-Host "Would you like to add devices now? (Y/N)"
    
    if ($addDevices -eq 'Y' -or $addDevices -eq 'y')
    {
      $newDevices = Add-DevicesToConfiguration
      
      if ($newDevices.Count -gt 0)
      {
        # Update the configuration with new devices
        $config.CertificatesToManage = $newDevices
        $config | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath -Encoding UTF8
        
        $certificatesToManage = $newDevices
        Write-Host "`nConfiguration updated with $($newDevices.Count) device(s)." -ForegroundColor Green
      }
      else
      {
        Write-Host "`nNo devices added. Exiting." -ForegroundColor Yellow
        exit 0
      }
    }
    else
    {
      Write-Host "`nNo devices to manage. Exiting." -ForegroundColor Yellow
      exit 0
    }
  }
}
else
{
  $config = New-DefaultConfiguration -ConfigPath $configPath
  
  # Convert to array of hashtables
  $certificatesToManage = @()
  foreach ($cert in $config.CertificatesToManage)
  {
    $certificatesToManage += $cert
  }
}

# Set paths based on configuration
$certsRootDirectory = $TargetDirectory
$caCertificateDir = Join-Path $TargetDirectory $($config.CACertificateName)
$caCertificatePath = Join-Path $caCertificateDir "$($config.CACertificateName).pem"

# Validate CA certificate exists or create it
if (-not (Test-Path -Path $caCertificatePath))
{
  Write-Host "CA certificate not found at: $caCertificatePath" -ForegroundColor Yellow
  Write-Host "Creating Root CA certificate..." -ForegroundColor Cyan
  
  $createRootCertScript = Join-Path $PSScriptRoot "Create-RootCert.ps1"
  
  if (-not (Test-Path $createRootCertScript))
  {
    Write-Error "Create-RootCert.ps1 script not found at: $createRootCertScript"
    exit 1
  }
  
  try
  {
    & $createRootCertScript -TargetDirectory $TargetDirectory -Country $config.Country -OrgUnit $config.OrgUnit -CaCertName $config.CACertificateName
    
    # Verify the certificate was created
    if (-not (Test-Path -Path $caCertificatePath))
    {
      Write-Error "Failed to create Root CA certificate."
      exit 1
    }
    
    Write-Host "Root CA certificate created successfully!" -ForegroundColor Green
  }
  catch
  {
    Write-Error "Error creating Root CA certificate: $($_.Exception.Message)"
    exit 1
  }
}

class RemoteFileUploadInfo
{
  [string]$Local
  [string]$Remote

  RemoteFileUploadInfo() { }

  RemoteFileUploadInfo([string]$Local, [string]$Remote)
  {
    $this.Local = $Local
    $this.Remote = $Remote
  }
}

function UploadCertsUsingScp
{
  param (
    [string]$CaCertPath,
    [string]$CertPath,
    [string]$CertDir,
    [string]$IPAddress,
    [string]$FQDN,
    [string]$Username,
    [SecureString]$Password,
    [hashtable]$FileMappings,
    [string]$SshKeyPath
  )

  Write-Host "Updating device $FQDN with certificates using SCP..."

  $creds = $null
  if (-not $SshKeyPath)
  {
    $creds = New-Object System.Management.Automation.PSCredential ($Username, $Password)
  }

  # Build file upload list based on mappings - no defaults
  $files = @()
  
  if ($FileMappings -and $FileMappings.Count -gt 0)
  {
    Write-Host "Using custom file mappings from configuration..."
    
    # Process each mapping
    foreach ($key in $FileMappings.Keys)
    {
      $remotePath = $FileMappings[$key]
      $localFile = $null
      
      # Determine which local file to use based on the key
      switch ($key)
      {
        "RootCA" { $localFile = $CaCertPath }
        "PFX" { $localFile = $CertPath }
        "PEM" { $localFile = Join-Path $CertDir "$FQDN.pem" }
        "CRT" { $localFile = Join-Path $CertDir "$FQDN.crt" }
        "KEY" { $localFile = Join-Path $CertDir "$FQDN.key" }
        default { Write-Warning "Unknown file mapping key: $key" }
      }
      
      if ($localFile -and (Test-Path $localFile))
      {
        $files += [RemoteFileUploadInfo]::new($localFile, $remotePath)
        Write-Host "  Mapping: $key -> $remotePath"
      }
      else
      {
        Write-Warning "Local file not found for mapping '$key': $localFile"
      }
    }
  }
  else
  {
    Write-Error "FileMappings configuration is required for Linux UpdateType. Please specify remote paths in your configuration."
    throw "FileMappings configuration is required for Linux UpdateType"
  }

  # Upload files via SCP
  Write-Host "Uploading certificate files to Linux device via SCP..."
  try
  {
    if ($SshKeyPath)
    {
      $uploadResults = UploadFilesViaScp -ComputerName $IPAddress -Files $files -KeyFilePath $SshKeyPath
    }
    else
    {
      $uploadResults = UploadFilesViaScp -ComputerName $IPAddress -Files $files -Credential $creds
    }
    
    $failedUploads = $uploadResults | Where-Object { -not $_.Success }
    if ($failedUploads.Count -gt 0)
    {
      $errorMsg = "Failed to upload $($failedUploads.Count) file(s) to Linux device."
      foreach ($fail in $failedUploads)
      {
        $errorMsg += "`n  - $($fail.LocalFile) -> $($fail.RemotePath): $($fail.Message)"
      }
      throw $errorMsg
    }
    
    Write-Host "All certificate files uploaded successfully to Linux device." -ForegroundColor Green
  }
  catch
  {
    Write-Error "Failed to upload certificates to Linux device: $($_.Exception.Message)"
    throw
  }
}

function UpdateUniFiCertificate
{
  param (
    [string]$CaCertPath,
    [string]$CertPath,
    [string]$CertDir,
    [string]$IPAddress,
    [string]$FQDN,
    [string]$Username,
    [SecureString]$Password,
    [string]$SshKeyPath
  )

  Write-Host "Updating UniFi DreamMachine Pro $FQDN with certificates..."

  $creds = $null
  if (-not $SshKeyPath)
  {
    $creds = New-Object System.Management.Automation.PSCredential ($Username, $Password)
  }

  # UniFi OS stores certificates in /data/unifi-core/config
  # We need to upload the certificate and key for both the main UI and direct access
  # unifi-core.crt/key: Main UniFi OS interface
  # unifi-core-direct.crt/key: Direct device access (when accessing by IP)
  $remoteCertPath = "/data/unifi-core/config/unifi-core.crt"
  $remoteKeyPath = "/data/unifi-core/config/unifi-core.key"
  $remoteDirectCertPath = "/data/unifi-core/config/unifi-core-direct.crt"
  $remoteDirectKeyPath = "/data/unifi-core/config/unifi-core-direct.key"
  
  # Prepare files to upload
  $crtPath = Join-Path $CertDir "$FQDN.crt"
  $keyPath = Join-Path $CertDir "$FQDN.key"
  
  if (-not (Test-Path $crtPath))
  {
    throw "Certificate file not found: $crtPath"
  }
  if (-not (Test-Path $keyPath))
  {
    throw "Private key file not found: $keyPath"
  }
  
  # Need to create a full chain certificate (server cert + CA cert)
  Write-Host "Creating full chain certificate..."
  $serverCert = Get-Content -Path $crtPath -Raw
  $caCert = Get-Content -Path $CaCertPath -Raw
  $fullChain = $serverCert.TrimEnd() + "`n" + $caCert.TrimEnd()
  
  # Write full chain to temporary file
  $tempChainPath = Join-Path $CertDir "$FQDN-fullchain.crt"
  Set-Content -Path $tempChainPath -Value $fullChain -NoNewline
  
  # Upload same certificates to both locations (main UI and direct access)
  $files = @(
    [RemoteFileUploadInfo]::new($tempChainPath, $remoteCertPath),
    [RemoteFileUploadInfo]::new($keyPath, $remoteKeyPath),
    [RemoteFileUploadInfo]::new($tempChainPath, $remoteDirectCertPath),
    [RemoteFileUploadInfo]::new($keyPath, $remoteDirectKeyPath)
  )

  try
  {
    # Upload files via SCP
    Write-Host "Uploading certificate files to UniFi device via SCP..."
    if ($SshKeyPath)
    {
      $uploadResults = UploadFilesViaScp -ComputerName $IPAddress -Files $files -KeyFilePath $SshKeyPath
    }
    else
    {
      $uploadResults = UploadFilesViaScp -ComputerName $IPAddress -Files $files -Credential $creds
    }
    
    $failedUploads = $uploadResults | Where-Object { -not $_.Success }
    if ($failedUploads.Count -gt 0)
    {
      $errorMsg = "Failed to upload $($failedUploads.Count) file(s) to UniFi device."
      foreach ($fail in $failedUploads)
      {
        $errorMsg += "`n  - $($fail.LocalFile) -> $($fail.RemotePath): $($fail.Message)"
      }
      throw $errorMsg
    }
    
    Write-Host "Certificate files uploaded successfully." -ForegroundColor Green
    
    # Execute commands to apply the certificate
    Write-Host "Applying certificate to UniFi OS..."
    $commands = @(
      "chmod 644 $remoteCertPath",
      "chmod 600 $remoteKeyPath",
      "chmod 644 $remoteDirectCertPath",
      "chmod 600 $remoteDirectKeyPath",
      "systemctl restart unifi-core"
    )
    
    if ($SshKeyPath)
    {
      $results = ExecuteSshCommands -ComputerName $IPAddress -Command $commands -KeyFilePath $SshKeyPath
    }
    else
    {
      $results = ExecuteSshCommands -ComputerName $IPAddress -Command $commands -Credential $creds
    }
    Write-Host $results
    Write-Host "UniFi certificate installation completed. The unifi-core service is restarting..." -ForegroundColor Green
    Write-Host "Note: It may take 30-60 seconds for the UniFi OS web interface to come back online." -ForegroundColor Yellow
  }
  catch
  {
    Write-Error "Failed to update UniFi device: $($_.Exception.Message)"
    throw
  }
  finally
  {
    # Clean up temporary full chain file
    if (Test-Path $tempChainPath)
    {
      Remove-Item -Path $tempChainPath -Force
    }
  }
}

function UpdateTrueNASCertificate
{
  param (
    [string]$CaCertPath,
    [string]$CertPath,
    [string]$CertDir,
    [string]$IPAddress,
    [string]$FQDN,
    [string]$ApiKey
  )

  Write-Host "Updating TrueNAS Scale device $FQDN with certificates..."

  # TrueNAS Scale API endpoint
  $baseUrl = "https://${IPAddress}/api/v2.0"
  
  # Read certificate files - need to read the .crt file, not the .pfx
  $crtPath = Join-Path $CertDir "$FQDN.crt"
  $keyPath = Join-Path $CertDir "$FQDN.key"
  
  if (-not (Test-Path $crtPath))
  {
    throw "Certificate file not found: $crtPath"
  }
  if (-not (Test-Path $keyPath))
  {
    throw "Private key file not found: $keyPath"
  }
  
  $certContent = Get-Content -Path $crtPath -Raw
  $keyContent = Get-Content -Path $keyPath -Raw
  
  # Validate certificate and key format
  if ($certContent -notmatch '-----BEGIN CERTIFICATE-----')
  {
    throw "Certificate file does not appear to be in PEM format"
  }
  if ($keyContent -notmatch '-----BEGIN (RSA )?PRIVATE KEY-----')
  {
    throw "Private key file does not appear to be in PEM format"
  }
  
  # Prepare headers
  $headers = @{
    "Authorization" = "Bearer $ApiKey"
    "Content-Type"  = "application/json"
  }
  
  # Create certificate payload - TrueNAS expects specific field names
  # Replace `r`n with `n and strip trailing newlines from certificate and key content
  $certContent = $certContent.TrimEnd("`r", "`n") -replace "`r`n", "`n"
  $keyContent = $keyContent.TrimEnd("`r", "`n") -replace "`r`n", "`n"
  $name = $FQDN -replace "\.", "-"
  Write-Host "Preparing to upload certificate with name: $name"
  $certPayload = @{
    csr_id      = 0
    name        = $name
    certificate = $certContent
    privatekey  = $keyContent
    create_type = "CERTIFICATE_CREATE_IMPORTED"
  } | ConvertTo-Json -Depth 10
  
  try
  {
    Write-Host "Uploading certificate to TrueNAS via API..."
    # Check if certificate already exists
    try
    {
      $existingCerts = Invoke-RestMethod -Uri "$baseUrl/certificate" -Headers $headers -Method Get -SkipCertificateCheck -ErrorAction Stop
      $existingCert = $existingCerts | Where-Object { $_.name -eq $name }
      
      if ($existingCert)
      {
        Write-Host "Certificate '$name' already exists (ID: $($existingCert.id)). Deleting old certificate..." -ForegroundColor Yellow
        Invoke-RestMethod -Uri "$baseUrl/certificate/id/$($existingCert.id)" -Headers $headers -Method Delete -SkipCertificateCheck | Out-Null
        Write-Host "Old certificate deleted." -ForegroundColor Green
        # Give TrueNAS a moment to process the deletion
        Start-Sleep -Seconds 2
      }
    }
    catch
    {
      Write-Warning "Could not check for existing certificates: $($_.Exception.Message)"
    }
    
    # Create new certificate
    Write-Host "Creating new certificate..."
    Write-Host "Certificate length: $($certContent.Length) bytes"
    Write-Host "Key length: $($keyContent.Length) bytes"
    Write-Host "Payload size: $($certPayload.Length) bytes"
    
    try
    {
      Invoke-RestMethod -Uri "$baseUrl/certificate" -Headers $headers -Method Post -Body $certPayload -SkipCertificateCheck -ErrorAction Stop | Out-Null
    }
    catch
    {
      Write-Error "API call failed: $($_.Exception.Message)"
      if ($_.ErrorDetails.Message)
      {
        Write-Error "API Error Details: $($_.ErrorDetails.Message)"
      }
      # Try to parse the response
      if ($_.Exception.Response)
      {
        Write-Host "Response Status: $($_.Exception.Response.StatusCode)"
      }
      throw
    }
    
    # TrueNAS API may not return any information about the certificate, so we need to retrieve it again to get the ID
    Write-Host "Retrieving newly created certificate ID..."
    Start-Sleep -Seconds 2
    $allCerts = Invoke-RestMethod -Uri "$baseUrl/certificate" -Headers $headers -Method Get -SkipCertificateCheck -ErrorAction Stop
    $newCert = $allCerts | Where-Object { $_.name -eq $name }
    if (-not $newCert)
    {
      throw "Could not find newly created certificate in certificate list."
    }
    
    $certId = $newCert.id
    Write-Host "Certificate uploaded successfully with ID: $certId" -ForegroundColor Green
    $verifiedCert = $newCert
    
    if ($verifiedCert)
    {
      Write-Host "Certificate verified in system with name: $($verifiedCert.name)" -ForegroundColor Green
        
      # Set as UI certificate
      Write-Host "Setting as UI certificate..."
      try
      {
        $uiPayload = @{ ui_certificate = $certId } | ConvertTo-Json
        Invoke-RestMethod -Uri "$baseUrl/system/general" -Headers $headers -Method Put -Body $uiPayload -SkipCertificateCheck -ErrorAction Stop | Out-Null
        Write-Host "Successfully set as UI certificate." -ForegroundColor Green
      }
      catch
      {
        Write-Error "Certificate uploaded but could not set as UI certificate: $($_.Exception.Message)"
        if ($_.ErrorDetails.Message)
        {
          Write-Error "Details: $($_.ErrorDetails.Message)"
        }
        
        throw "Failed to set certificate as UI certificate"
      }
    }
    else
    {
      Write-Error "Certificate verification failed - certificate with ID $certId not found in certificate list"
      Write-Host "Available certificates:"
      $allCerts | ForEach-Object { Write-Host "  ID: $($_.id), Name: $($_.name)" }
      throw "Certificate verification failed"
    }
  }
  catch
  {
    $errorMsg = $_.Exception.Message
    Write-Error "Failed to upload certificate to TrueNAS: $errorMsg"
    
    # Try to get more detailed error information
    if ($_.ErrorDetails.Message)
    {
      Write-Error "API Error Details: $($_.ErrorDetails.Message)"
    }
    
    throw
  }

  Write-Host "TrueNAS certificate installation completed successfully." -ForegroundColor Green
}

function UpdateCrestronCertificate
{
  param (
    [string]$CaCertPath,
    [string]$CertPath,
    [string]$IPAddress,
    [string]$FQDN,
    [string]$Username,
    [SecureString]$Password,
    [string]$Series,
    [string]$OrgUnit,
    [string]$ConnectionAddress = ""
  )

  # Determine the address to use for SFTP/SSH connections
  $computerName = if ([string]::IsNullOrWhiteSpace($ConnectionAddress)) { $FQDN } else { $ConnectionAddress }

  Write-Host "Updating Crestron device $FQDN with certificate from $CertPath for user $Username"

  # Upload the CA certificate to the Crestron device using SFTP to place it in the "/sys" directory

  $creds = New-Object System.Management.Automation.PSCredential ($Username, $Password)

  # Call the reusable SFTP upload function to place the files in the correct locations, using the RemoteFileUploadInfo class
  if ($Series -eq "4")
  {
    Write-Host "Preparing to upload files to Crestron Series 4 device..."
    $files = @(
      [RemoteFileUploadInfo]::new($CaCertPath, "/cert/root_cert.cer"),
      [RemoteFileUploadInfo]::new($CertPath, "/cert/webserver_cert.pfx"),
      [RemoteFileUploadInfo]::new($CertPath, "/cert/websocket_cert.pfx")
    )
    
    $installCACommand = "certificate ADD ROOT root_cert.cer"
    $installWebServerCommand = "certificate ADDF webserver_cert.pfx WEBSERVER $certPassword"
    $installWebSocketCommand = "certificate ADDF websocket_cert.pfx WEBSOCKET $certPassword"

    $commands = @($installCACommand, $installWebServerCommand, $installWebSocketCommand)
  }
  elseif ($Series -eq "TP70" -or $Series -eq "TP60")
  {
    # Touchpanels do not like changing file names after they are uploaded, so we have to upload with the correct names directly
    # Create copies of the certs correctly named for all the files.
    $certDir = Split-Path $CertPath
    $rootCaCertPath = Join-Path $certDir "root_cert.cer"
    $webserverCertPath = Join-Path $certDir "webserver_cert.pfx"
    $websocketCertPath = Join-Path $certDir "websocket_cert.pfx"
    $sipCertPath = Join-Path $certDir "sip_cert.pfx"
    $streamCertPath = Join-Path $certDir "stream_cert.pfx"
    $clientCertPath = Join-Path $certDir "client_cert.pfx"
    Copy-Item -Path $CaCertPath -Destination $rootCaCertPath -Force
    Copy-Item -Path $CertPath -Destination $webserverCertPath -Force
    Copy-Item -Path $CertPath -Destination $websocketCertPath -Force
    Copy-Item -Path $CertPath -Destination $sipCertPath -Force
    Copy-Item -Path $CertPath -Destination $streamCertPath -Force
    Copy-Item -Path $CertPath -Destination $clientCertPath -Force

    Write-Host "Preparing to upload files to Crestron 70 Series TP, or other compatible device..."
    $files = @(
      [RemoteFileUploadInfo]::new($rootCaCertPath, "/User/Cert/root_cert.cer"),
      [RemoteFileUploadInfo]::new($webserverCertPath, "/User/Cert/webserver_cert.pfx"),
      [RemoteFileUploadInfo]::new($websocketCertPath, "/User/Cert/websocket_cert.pfx"),
      [RemoteFileUploadInfo]::new($sipCertPath, "/User/Cert/sip_cert.pfx"),
      [RemoteFileUploadInfo]::new($streamCertPath, "/User/Cert/stream_cert.pfx"),
      [RemoteFileUploadInfo]::new($clientCertPath, "/User/Cert/client_cert.pfx")
    )
    
    $installCACommand = "certificate ADD ROOT root_cert.cer"
    $installWebServerCommand = "certificate ADDF webserver_cert.pfx WEBSERVER $certPassword"
    $installWebSocketCommand = "certificate ADDF websocket_cert.pfx WEBSOCKET $certPassword"
    $installSipCommand = "certificate ADDF sip_cert.pfx SIP $certPassword"
    $installStreamCommand = "certificate ADDF stream_cert.pfx STREAM $certPassword"
    $installClientCommand = "certificate ADDF client_cert.pfx CLIENT $certPassword"

    # Only the 70 Series compatible panels accept the client cert.
    if ($Series -eq "TP70") {
      $commands = @($installCACommand, $installWebServerCommand, $installWebSocketCommand, $installSipCommand, $installStreamCommand, $installClientCommand)
    } else {
      $commands = @($installCACommand, $installWebServerCommand, $installWebSocketCommand, $installSipCommand, $installStreamCommand)
    }
  }
  elseif ($Series -eq "3")
  {
    Write-Host "Preparing to upload files to Crestron Series 3 device..."
    
    # Get paths to the required certificate files
    $certDir = Split-Path $CertPath
    $srvCertPath = Join-Path $certDir "$FQDN.crt"
    $srvKeyPath = Join-Path $certDir "$FQDN.key"
    $rootCAPem = $CaCertPath -replace '\.cer$', '.pem'
    
    # For Series 3, upload to /user directory first, then move to /sys via SSH
    # This is because /sys is not directly accessible via SFTP
    $files = @(
      [RemoteFileUploadInfo]::new($CaCertPath, "/CERT/root_cert.cer"),
      [RemoteFileUploadInfo]::new($CertPath, "/CERT/machine_cert.pfx"),
      [RemoteFileUploadInfo]::new($CertPath, "/CERT/websocket_cert.pfx")
      [RemoteFileUploadInfo]::new($rootCAPem, "/user/rootCA_cert.cer"),
      [RemoteFileUploadInfo]::new($srvCertPath, "/user/srv_cert.cer"),
      [RemoteFileUploadInfo]::new($srvKeyPath, "/user/srv_key.pem")
    )
    
    # Commands to delete old files in /sys, move new files, and install certificates
    $deleteRootCACommand = "delete /sys/rootCA_cert.cer"
    $deleteSrvCertCommand = "delete /sys/srv_cert.cer"
    $deleteSrvKeyCommand = "delete /sys/srv_key.pem"
    $moveRootCACommand = "move /user/rootCA_cert.cer /sys/rootCA_cert.cer"
    $moveSrvCertCommand = "move /user/srv_cert.cer /sys/srv_cert.cer"
    $moveSrvKeyCommand = "move /user/srv_key.pem /sys/srv_key.pem"
    $installCACommand = "certificate ADD ROOT root_cert.cer"
    $installWebSocketCommand = "certificate ADD WEBSOCKET $certPassword"
    $installMachineCommand = "certificate ADD MACHINE $certPassword"
    $installSrvCommand = "ssl CA"

    $commands = @(
      $deleteRootCACommand,
      $deleteSrvCertCommand, 
      $deleteSrvKeyCommand,
      $moveRootCACommand,
      $moveSrvCertCommand,
      $moveSrvKeyCommand,      
      $installCACommand,
      $installWebSocketCommand,
      $installMachineCommand,
      $installSrvCommand
    )
    
    # Note: For intermediate certs, use "certificate add" or "certificate addf" commands
    # The root CA and server cert files will be in /sys after the move commands
  }
  else
  {
    Write-Error "Unsupported Crestron series: $Series"
  }

  UploadFilesViaSftp -ComputerName $computerName -Files $files -Credential $creds

  # Check if root CA is already installed and remove it if needed
  Write-Host "Checking for existing root CA certificates..."
  try
  {
    $listRootResult = ExecuteSshCommand -ComputerName $computerName -Command "certificate list root" -Credential $creds
    
    if ($listRootResult)
    {
      $rootListText = $listRootResult -join "`n"
      Write-Host "Current root certificates:"
      Write-Host $rootListText
      
      # Parse the output to find matching CA cert
      # Look for the CN from the CA certificate (e.g., "EVandS Private Root CA")
      $caCN = "$($OrgUnit) Private Root CA"
      
      # Check if our CA is in the list
      if ($rootListText -match [regex]::Escape($caCN))
      {
        Write-Host "Found existing root CA '$caCN' - removing before reinstall..." -ForegroundColor Yellow
        
        # Extract the UID from the line containing our CA
        $caLine = $rootListText -split "`n" | Where-Object { $_ -match [regex]::Escape($caCN) } | Select-Object -First 1
        
        if ($caLine -match '\|\s*([A-F0-9]+)\s*\|')
        {
          $uid = $matches[1].Trim()
          Write-Host "Removing root CA with UID: $uid"
          
          $removeCommand = "certificate rem root $($OrgUnit) Private Root CA $uid"
          ExecuteSshCommand -ComputerName $computerName -Command $removeCommand -Credential $creds | Out-Null
          
          Write-Host "Root CA removed successfully" -ForegroundColor Green
        }
        else
        {
          Write-Warning "Could not parse UID from root certificate list"
        }
      }
      else
      {
        Write-Host "Root CA '$caCN' not found in existing certificates - proceeding with fresh install" -ForegroundColor Green
      }
    }
  }
  catch
  {
    Write-Warning "Could not check/remove existing root CA: $($_.Exception.Message)"
    Write-Host "Proceeding with installation anyway..."
  }

  # Execute the command to install the certificate on the Crestron device
  Write-Host "Installing certificates on Crestron device..."
  
  $commandResults = ExecuteSshCommands -ComputerName $computerName -Command $commands -Credential $creds
  
  # Check if any commands failed
  $failedCommands = $commandResults | Where-Object { -not $_.Success }
  if ($failedCommands.Count -gt 0)
  {
    $errorMsg = "Certificate installation failed on Crestron device. $($failedCommands.Count) command(s) failed."
    $errorMsg += "`nFailed Commands:`n"
    foreach ($fail in $failedCommands)
    {
      $commandText = [string]$fail.Command
      # Redact known certificate password if present
      try
      {
        if ($null -ne $config -and $config.CrestronCertPassword)
        {
          $redacted = $commandText -replace [regex]::Escape([string]$config.CrestronCertPassword), '***'
        }
        else
        {
          $redacted = $commandText
        }
      }
      catch
      {
        $redacted = $commandText
      }
      $errorMsg += "Command: $redacted`nMessage: $($fail.Message)`nOutput: $($fail.Output)`n`n"
    }

    Write-Error $errorMsg
    throw $errorMsg
  }
  
  Write-Host "Certificates installed successfully on Crestron device." -ForegroundColor Green
}


function UploadFilesViaScp
{
  param(
    [Parameter(Mandatory)][string]$ComputerName,
    [Parameter(Mandatory)][RemoteFileUploadInfo[]]$Files,
    [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential,
    [Parameter(Mandatory = $false)][string]$KeyFilePath
  )

  if (-not (Get-Module -ListAvailable -Name Posh-SSH))
  {
    Write-Host "Posh-SSH module not found. Installing..."
    Install-Module -Name Posh-SSH -Force -Scope CurrentUser -ErrorAction Stop
  }
  Import-Module Posh-SSH -ErrorAction Stop

  $results = @()
  $sshSession = $null

  try
  {
    # Create SSH session for directory creation and file removal
    Write-Host "Opening SSH session to $ComputerName..."
    if ($KeyFilePath)
    {
      Write-Host "Using SSH key authentication: $KeyFilePath"
      $sshSession = New-SSHSession -ComputerName $ComputerName -KeyFile $KeyFilePath -AcceptKey -ErrorAction Stop
    }
    else
    {
      $sshSession = New-SSHSession -ComputerName $ComputerName -Credential $Credential -AcceptKey -ErrorAction Stop
    }
    $sessionId = if ($sshSession -is [System.Array]) { $sshSession[0].SessionId } else { $sshSession.SessionId }

    foreach ($fileToUpload in $Files)
    {
      $localPath = $fileToUpload.Local
      $remotePath = $fileToUpload.Remote

      if (-not (Test-Path $localPath))
      {
        Write-Warning "Local file not found: $localPath"
        $results += [PSCustomObject]@{
          LocalFile  = $localPath
          RemotePath = $remotePath
          Success    = $false
          Message    = "Local file not found"
        }
        continue
      }

      try
      {
        # Extract directory and filename
        $remoteDir = [System.IO.Path]::GetDirectoryName($remotePath).Replace('\', '/')
        $remoteFile = [System.IO.Path]::GetFileName($remotePath)
        
        # Ensure remote directory exists
        if (-not [string]::IsNullOrWhiteSpace($remoteDir))
        {
          Write-Host "Ensuring remote directory exists: $remoteDir"
          $mkdirCmd = "mkdir -p `"$remoteDir`""
          $mkdirResult = Invoke-SSHCommand -SessionId $sessionId -Command $mkdirCmd -ErrorAction Stop
          
          if ($mkdirResult.ExitStatus -ne 0)
          {
            Write-Warning "mkdir command output: $($mkdirResult.Output -join "`n")"
          }
        }
        
        # Remove existing file if present
        Write-Host "Removing existing file if present: $remotePath"
        $rmCmd = "rm -f `"$remotePath`""
        Invoke-SSHCommand -SessionId $sessionId -Command $rmCmd -ErrorAction SilentlyContinue | Out-Null

        Write-Host "Uploading '$localPath' to '$remoteDir/' via SCP..."
        
        # Upload to directory, SCP will use the local filename
        Set-SCPItem -ComputerName $ComputerName -Credential $Credential -Path $localPath -Destination $remoteDir -AcceptKey -Force -ErrorAction Stop
        
        # If uploaded filename differs from desired, rename it
        $uploadedPath = "$remoteDir/$([System.IO.Path]::GetFileName($localPath))"
        if ($uploadedPath -ne $remotePath)
        {
          Write-Host "Renaming to desired filename: $remoteFile"
          $mvCmd = "mv -f `"$uploadedPath`" `"$remotePath`""
          $mvResult = Invoke-SSHCommand -SessionId $sessionId -Command $mvCmd -ErrorAction Stop
          
          if ($mvResult.ExitStatus -ne 0)
          {
            throw "Failed to rename file: $($mvResult.Output -join "`n")"
          }
        }
        
        $results += [PSCustomObject]@{
          LocalFile  = $localPath
          RemotePath = $remotePath
          Success    = $true
          Message    = "Uploaded"
        }
      }
      catch
      {
        $errorMsg = $_.Exception.Message
        Write-Error "Failed to upload '$localPath' to '$remotePath': $errorMsg"
        
        $results += [PSCustomObject]@{
          LocalFile  = $localPath
          RemotePath = $remotePath
          Success    = $false
          Message    = $errorMsg
        }
      }
    }
  }
  finally
  {
    if ($sshSession)
    {
      foreach ($s in @($sshSession))
      {
        try { Remove-SSHSession -SessionId $s.SessionId -ErrorAction SilentlyContinue | Out-Null } catch {}
      }
    }
  }

  return $results
}

function UploadFilesViaSftp
{
  param(
    [Parameter(Mandatory)][string]$ComputerName,
    [Parameter(Mandatory)][RemoteFileUPloadInfo[]]$Files,
    [Parameter(Mandatory)][System.Management.Automation.PSCredential]$Credential
  )

  if (-not (Get-Module -ListAvailable -Name Posh-SSH))
  {
    Write-Host "Posh-SSH module not found. Installing..."
    Install-Module -Name Posh-SSH -Force -Scope CurrentUser -ErrorAction Stop
  }
  Import-Module Posh-SSH -ErrorAction Stop

  $sessions = $null
  $results = @()

  try
  {
    Write-Host "Opening SFTP session to $ComputerName..."
    $sessions = New-SFTPSession -ComputerName $ComputerName -Credential $Credential -AcceptKey -ErrorAction Stop

    $sess = if ($sessions -is [System.Array]) { $sessions[0] } else { $sessions }

    # Temporary troubleshooting: List root directory contents
    try
    {
      Write-Host "--- Listing root directory contents for troubleshooting ---"
      $rootContents = Get-SFTPChildItem -SessionId $sess.SessionId -Path "/" -ErrorAction Stop
      foreach ($item in $rootContents)
      {
        Write-Host "  $($item.FullName) [$($item.GetType().Name)]"
      }
      Write-Host "--- End of root directory listing ---"
    }
    catch
    {
      Write-Warning "Could not list root directory: $($_.Exception.Message)"
    }

    foreach ($fileToUpload in $Files)
    {
      # Basic validation / normalization
      $localPath = $null
      $remoteSpec = $null

      try
      {
        $localPath = $fileToUpload.Local
        $remoteSpec = $fileToUpload.Remote
      }
      catch
      {
        $results += [PSCustomObject]@{
          LocalFile  = $null
          RemotePath = $null
          Success    = $false
          Message    = "File entry must have 'Local' property"
        }
        continue
      }

      if (-not $localPath)
      {
        $results += [PSCustomObject]@{
          LocalFile  = $null
          RemotePath = $remoteSpec
          Success    = $false
          Message    = "Local path missing"
        }
        continue
      }

      if (-not (Test-Path -Path $localPath))
      {
        $results += [PSCustomObject]@{
          LocalFile  = $localPath
          RemotePath = $remoteSpec
          Success    = $false
          Message    = "Local fileToUpload not found"
        }
        continue
      }

      $remoteName = [System.IO.Path]::GetFileName($localPath)
      if ([string]::IsNullOrWhiteSpace($remoteSpec))
      {
        $remotePath = "/$remoteName"
      }
      elseif ($remoteSpec.EndsWith("/"))
      {
        $remotePath = "$remoteSpec$remoteName"
      }
      else
      {
        $remotePath = $remoteSpec
      }

      try
      {
        # Ensure the remote directory exists
        $remoteDir = [System.IO.Path]::GetDirectoryName($remotePath).Replace('\', '/')
        $remoteFileName = [System.IO.Path]::GetFileName($remotePath)
        
        Write-Host "Ensuring remote directory '$remoteDir' exists..."
        if (-not [string]::IsNullOrWhiteSpace($remoteDir) -and $remoteDir -ne '/')
        {
          # Create directory path recursively if needed
          $pathParts = $remoteDir.Split('/', [StringSplitOptions]::RemoveEmptyEntries)
          $currentPath = ""
          
          foreach ($part in $pathParts)
          {
            $currentPath += "/$part"
            try
            {
              $pathExists = Test-SFTPPath -SessionId $sess.SessionId -Path $currentPath
              if (-not $pathExists)
              {
                Write-Host "Creating directory '$currentPath'..."
                New-SFTPItem -SessionId $sess.SessionId -Path $currentPath -ItemType Directory -ErrorAction Stop
              }
            }
            catch
            {
              Write-Warning "Could not verify/create directory '$currentPath': $($_.Exception.Message)"
            }
          }
        }

        Write-Host "Uploading '$localPath' to '$remotePath'..."
        
        # Delete existing file if it exists
        try
        {
          if (Test-SFTPPath -SessionId $sess.SessionId -Path $remotePath)
          {
            Write-Host "Deleting existing file at '$remotePath'..."
            Remove-SFTPItem -SessionId $sess.SessionId -Path $remotePath -Force -ErrorAction Stop
          }
        }
        catch
        {
          Write-Warning "Could not delete existing file '$remotePath': $($_.Exception.Message)"
        }
        
        # Upload to the directory, then rename if needed
        if (-not [string]::IsNullOrWhiteSpace($remoteDir) -and $remoteDir -ne '/')
        {
          # Upload to directory and let it use the destination filename
          Set-SFTPItem -SessionId $sess.SessionId -Path $localPath -Destination $remoteDir -ErrorAction Stop
          
          # If the uploaded filename differs from desired, rename it
          $uploadedPath = "$remoteDir/$([System.IO.Path]::GetFileName($localPath))"
          if ($uploadedPath -ne $remotePath)
          {
            Write-Host "Renaming '$uploadedPath' to '$remotePath'..."
            Rename-SFTPFile -SessionId $sess.SessionId -Path $uploadedPath -NewName $remoteFileName -ErrorAction Stop
          }
        }
        else
        {
          # Upload to root
          Set-SFTPItem -SessionId $sess.SessionId -Path $localPath -Destination "/" -ErrorAction Stop
          
          # Rename if needed
          $uploadedPath = "/$([System.IO.Path]::GetFileName($localPath))"
          if ($uploadedPath -ne $remotePath)
          {
            Write-Host "Renaming '$uploadedPath' to '$remotePath'..."
            Rename-SFTPFile -SessionId $sess.SessionId -Path $uploadedPath -NewName $remoteFileName -ErrorAction Stop
          }
        }

        $results += [PSCustomObject]@{
          LocalFile  = $localPath
          RemotePath = $remotePath
          Success    = $true
          Message    = "Uploaded"
        }
      }
      catch
      {
        $errorMsg = $_.Exception.Message
        Write-Error "Failed to upload '$localPath' to '$remotePath': $errorMsg"
        
        $results += [PSCustomObject]@{
          LocalFile  = $localPath
          RemotePath = $remotePath
          Success    = $false
          Message    = $errorMsg
        }
      }
    }

    return $results
  }
  catch
  {
    Write-Host "Error establishing SFTP session: $($_.Exception.Message)"
    return $results
  }
  finally
  {
    if ($sessions)
    {
      foreach ($s in @($sessions))
      {
        try { Remove-SFTPSession -SessionId $s.SessionId -ErrorAction SilentlyContinue } catch {}
      }
    }
  }
}

function ExecuteSshCommands
{
  param(
    [Parameter(Mandatory)][string]$ComputerName,
    [Parameter(Mandatory)][string[]]$Command,
    [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential,
    [Parameter(Mandatory = $false)][string]$KeyFilePath
  )

  if (-not (Get-Module -ListAvailable -Name Posh-SSH))
  {
    Write-Host "Posh-SSH module not found. Installing..."
    Install-Module -Name Posh-SSH -Force -Scope CurrentUser -ErrorAction Stop
  }
  Import-Module Posh-SSH -ErrorAction Stop

  $sessions = $null
  $results = @()

  try
  {
    Write-Host "Opening SSH session to $ComputerName..."
    $sessions = New-SSHSession -ComputerName $ComputerName -Credential $Credential -AcceptKey -ErrorAction Stop

    $sess = if ($sessions -is [System.Array]) { $sessions[0] } else { $sessions }

    foreach ($cmd in $Command)
    {
      Write-Host "Executing SSH command..."
      try
      {
        $result = Invoke-SSHCommand -SessionId $sess.SessionId -Command $cmd -ErrorAction Stop
        
        $outputText = if ($result.Output) { ($result.Output -join "`n") } else { "" }
        
        # Check for "Error" in output (case insensitive)
        $hasError = $outputText -match '(?i)error'
        
        if ($hasError)
        {
          Write-Error "Command failed with error in response:"
          Write-Host $outputText -ForegroundColor Red
        }

        $results += [PSCustomObject]@{
          Command      = $cmd
          Success      = -not $hasError
          Output       = $outputText
          ErrorMessage = if ($hasError) { "Error detected in command output" } else { $null }
        }
      }
      catch
      {
        $results += [PSCustomObject]@{
          Command      = $cmd
          Success      = $false
          Output       = $null
          ErrorMessage = $_.Exception.Message
        }
      }
    }

    return $results
  }
  catch
  {
    Write-Host "Error during SSH session: $($_.Exception.Message)"
    return $results
  }
  finally
  {
    if ($sessions)
    {
      foreach ($s in @($sessions))
      {
        try { Remove-SSHSession -SessionId $s.SessionId -ErrorAction SilentlyContinue | Out-Null } catch {}
      }
    }
  }
}

function ExecuteSshCommand
{
  param(
    [Parameter(Mandatory)][string]$ComputerName,
    [Parameter(Mandatory)][string]$Command,
    [Parameter(Mandatory)][System.Management.Automation.PSCredential]$Credential
  )

  if (-not (Get-Module -ListAvailable -Name Posh-SSH))
  {
    Write-Host "Posh-SSH module not found. Installing..."
    Install-Module -Name Posh-SSH -Force -Scope CurrentUser -ErrorAction Stop
  }
  Import-Module Posh-SSH -ErrorAction Stop

  $sessions = $null
  try
  {
    Write-Host "Opening SSH session to $ComputerName..."
    $sessions = New-SSHSession -ComputerName $ComputerName -Credential $Credential -AcceptKey -ErrorAction Stop

    $sess = if ($sessions -is [System.Array]) { $sessions[0] } else { $sessions }

    Write-Host "Executing SSH command..."
    $result = Invoke-SSHCommand -SessionId $sess.SessionId -Command $Command -ErrorAction Stop
    
    return $result.Output
  }
  catch
  {
    Write-Host "Error during SSH command execution: $($_.Exception.Message)" 
  }
  finally
  {
    if ($sessions)
    {
      foreach ($s in @($sessions))
      {
        try { Remove-SSHSession -SessionId $s.SessionId -ErrorAction SilentlyContinue } catch {}
      }
    }
  }
}

# Compatibility wrapper for existing callers expecting UploadFileViaSftp
function UploadFileViaSftp
{
  param(
    [Parameter(Mandatory)][string]$ComputerName,
    [Parameter(Mandatory)][string]$LocalFile,
    [string]$RemoteDir = "/sys",
    [Parameter(Mandatory)][System.Management.Automation.PSCredential]$Credential
  )

  $result = UploadFilesViaSftp -ComputerName $ComputerName -LocalFiles @($LocalFile) -RemoteDir $RemoteDir -Credential $Credential
  if ($result -and $result.Count -gt 0) { return $result[0].Success } else { return $false }
}

# Prompt for CA private key password once
$caKeyPassword = ""
if ($certificatesToManage.Count -gt 0)
{
  Write-Host "`nEntering certificate generation and deployment process..." -ForegroundColor Cyan
  $caKeyPasswordSecure = Read-Host "Enter CA private key password (press Enter if no password)" -AsSecureString
  $caKeyPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($caKeyPasswordSecure))
}

# Track failures
$failedOperations = @()

# Filter devices if -Devices parameter was provided
if ($Devices -and $Devices.Count -gt 0)
{
  Write-Host "`nFiltering to specific devices: $($Devices -join ', ')" -ForegroundColor Cyan
  $certificatesToManage = $certificatesToManage | Where-Object {
    $Devices -contains $_.FQDN -or $Devices -contains $_.IPAddress
  }
  
  if ($certificatesToManage.Count -eq 0)
  {
    Write-Warning "No devices matched the specified filter: $($Devices -join ', ')"
    Write-Host "Available devices in configuration:"
    foreach ($cert in $certificatesToManage)
    {
      Write-Host "  - $($cert.FQDN) ($($cert.IPAddress))"
    }
    exit 0
  }
  
  Write-Host "Found $($certificatesToManage.Count) matching device(s)" -ForegroundColor Green
}

# Main processing loop
# Loop through each certificate entry and create/update as needed
foreach ($certEntry in $certificatesToManage)
{
  $fqdn = $certEntry.FQDN
  
  # Validate we have at least one IP
  if ($certEntry.IPAddress.Count -eq 0 -or [string]::IsNullOrWhiteSpace($certEntry.IPAddress[0]))
  {
    Write-Error "No valid IP address found for $fqdn. IPAddress must be an array of IP addresses."
    continue
  }
  
  # Use first IP address for connections
  $ipAddress = $certEntry.IPAddress.GetType().Name -eq "String" ? $certEntry.IPAddress : $certEntry.IPAddress[0]
  
  # Use ConnectionAddress if specified, otherwise fall back to default per-type behavior
  $connectionAddress = if ($certEntry.ContainsKey('ConnectionAddress') -and -not [string]::IsNullOrWhiteSpace($certEntry.ConnectionAddress)) { $certEntry.ConnectionAddress } else { $null }
  
  # Create comma-separated list of all IPs for certificate generation
    
  $username = $certEntry.Username ?? ""
  $passwordPlain = $certEntry.Password ?? ""
  $updateType = $certEntry.UpdateType

  Write-Host "`nProcessing certificate for $fqdn" -ForegroundColor Cyan
  
  $certGenerationFailed = $false
  $uploadFailed = $false
  $executionFailed = $false
  
  # Create the client certificate
  $certDir = Join-Path $certsRootDirectory $fqdn
  if (-not (Test-Path -Path $certDir))
  {
    New-Item -ItemType Directory -Path $certDir | Out-Null
  }

  # Call the existing Create-ClientCert.ps1 script to generate the certificate
  if ($certEntry.IPAddress.Count -gt 1)
  {
    Write-Host "Creating client certificate for $fqdn with IPs: $allIpAddresses (using $ipAddress for connection)"
  }
  else
  {
    Write-Host "Creating client certificate for $fqdn with IP $ipAddress"
  }
  # Create a password for the certificate based on whether it is a Crestron device
  $certPassword = $updateType -like "Crestron*" ? $config.CrestronCertPassword : ""
  
  # Determine if legacy encryption should be used (for Crestron Series 3)
  $useLegacy = $updateType -eq "Crestron3"
  
  try
  {
    $createCertScript = Join-Path $PSScriptRoot "Create-ClientCert.ps1"
    
    if ($useLegacy)
    {
      & $createCertScript -TargetDirectory $certsRootDirectory -Fqdn $fqdn -IpAddress $certEntry.IPAddress -CertPassword $certPassword -Country $config.Country -OrgUnit $config.OrgUnit -CaCertName $config.CACertificateName -CaCertificatePath $caCertificateDir -CaKeyPassword $caKeyPassword -UseLegacyEncryption
    }
    else
    {
      & $createCertScript -TargetDirectory $certsRootDirectory -Fqdn $fqdn -IpAddress $certEntry.IPAddress -CertPassword $certPassword -Country $config.Country -OrgUnit $config.OrgUnit -CaCertName $config.CACertificateName -CaCertificatePath $caCertificateDir -CaKeyPassword $caKeyPassword
    }
    
    # Verify certificate was created
    $certPath = Join-Path $certDir "$fqdn.pfx"
    if (-not (Test-Path $certPath))
    {
      throw "Certificate file was not created at $certPath"
    }
  }
  catch
  {
    Write-Error "Failed to create client certificate for $fqdn`: $($_.Exception.Message)"
    $certGenerationFailed = $true
    $failedOperations += [PSCustomObject]@{
      FQDN      = $fqdn
      Operation = "Certificate Generation"
      Error     = $_.Exception.Message
    }
    continue
  }

  # Path to the generated certificate (.pfx)
  $certPath = Join-Path $certDir "$fqdn.pfx"

  # Get SSH key path if specified
  $sshKeyPath = $null
  if ($certEntry.ContainsKey('SshKeyPath'))
  {
    $sshKeyPath = $certEntry.SshKeyPath
    # Expand environment variables and relative paths
    $sshKeyPath = [System.Environment]::ExpandEnvironmentVariables($sshKeyPath)
    if (-not [System.IO.Path]::IsPathRooted($sshKeyPath))
    {
      $sshKeyPath = Join-Path $TargetDirectory $sshKeyPath
    }
    if (-not (Test-Path $sshKeyPath))
    {
      Write-Warning "SSH key file not found: $sshKeyPath. Will attempt password authentication."
      $sshKeyPath = $null
    }
  }

  # Convert plain password to SecureString
  if ($passwordPlain -eq "" -and $updateType -ne "TrueNAS" -and -not $sshKeyPath)
  {
    Write-Error "Password or SshKeyPath is required for device $fqdn of type $updateType"
    $failedOperations += [PSCustomObject]@{
      FQDN      = $fqdn
      Operation = "Device Update"
      Error     = "Password or SshKeyPath is required for device of type $updateType"
    }
    continue
  }
  elseif ($updateType -ne "TrueNAS")
  {
    $securePassword = if ($passwordPlain) { ConvertTo-SecureString -String $passwordPlain -AsPlainText -Force } else { $null }
  }

  # Update the device based on the update type
  Write-Host "Updating device $fqdn of type $updateType"
  $rootCer = Join-Path $caCertificateDir "$($config.CACertificateName).cer"
  try
  {
    switch ($updateType)
    {
      "Crestron4"
      {
        UpdateCrestronCertificate -CaCertPath $rootCer -CertPath $certPath -IPAddress $ipAddress -FQDN $fqdn -Username $username -Password $securePassword -Series "4" -OrgUnit $config.OrgUnit -ConnectionAddress ($connectionAddress ?? "")
      }
      "Crestron3"
      {
        UpdateCrestronCertificate -CaCertPath $rootCer -CertPath $certPath -IPAddress $ipAddress -FQDN $fqdn -Username $username -Password $securePassword -Series "3" -OrgUnit $config.OrgUnit -ConnectionAddress ($connectionAddress ?? "")
      }
      "CrestronTP70Series"
      {
        UpdateCrestronCertificate -CaCertPath $rootCer -CertPath $certPath -IPAddress $ipAddress -FQDN $fqdn -Username $username -Password $securePassword -Series "TP70" -OrgUnit $config.OrgUnit -ConnectionAddress ($connectionAddress ?? "")
      }
      "CrestronTP60Series"
      {
        UpdateCrestronCertificate -CaCertPath $rootCer -CertPath $certPath -IPAddress $ipAddress -FQDN $fqdn -Username $username -Password $securePassword -Series "TP60" -OrgUnit $config.OrgUnit -ConnectionAddress ($connectionAddress ?? "")
      }
      "SCP"
      {
        # Get file mappings from config if specified
        $fileMappings = $null
        if ($certEntry.ContainsKey('FileMappings') -and $certEntry.FileMappings)
        {
          $fileMappings = @{}
          foreach ($prop in $certEntry.FileMappings.PSObject.Properties)
          {
            $fileMappings[$prop.Name] = $prop.Value
          }
        }
        
        $scpAddress = $connectionAddress ?? $ipAddress
        if ($sshKeyPath)
        {
          UploadCertsUsingScp -CaCertPath $rootCer -CertPath $certPath -CertDir $certDir -IPAddress $scpAddress -FQDN $fqdn -Username $username -Password $securePassword -FileMappings $fileMappings -SshKeyPath $sshKeyPath
        }
        else
        {
          UploadCertsUsingScp -CaCertPath $rootCer -CertPath $certPath -CertDir $certDir -IPAddress $scpAddress -FQDN $fqdn -Username $username -Password $securePassword -FileMappings $fileMappings
        }
      }
      "TrueNAS"
      {
        # Get API key from config
        $apiKey = $null
        if ($certEntry.ContainsKey('ApiKey'))
        {
          $apiKey = $certEntry.ApiKey
        }
        else
        {
          Write-Error "ApiKey is required for TrueNAS UpdateType"
          throw "ApiKey is required for TrueNAS UpdateType"
        }
        
        $truenasAddress = $connectionAddress ?? $ipAddress
        UpdateTrueNASCertificate -CaCertPath $rootCer -CertPath $certPath -CertDir $certDir -IPAddress $truenasAddress -FQDN $fqdn -ApiKey $apiKey
      }
      "UniFi"
      {
        $unifiAddress = $connectionAddress ?? $ipAddress
        if ($sshKeyPath)
        {
          UpdateUniFiCertificate -CaCertPath $rootCer -CertPath $certPath -CertDir $certDir -IPAddress $unifiAddress -FQDN $fqdn -Username $username -Password $securePassword -SshKeyPath $sshKeyPath
        }
        else
        {
          UpdateUniFiCertificate -CaCertPath $rootCer -CertPath $certPath -CertDir $certDir -IPAddress $unifiAddress -FQDN $fqdn -Username $username -Password $securePassword
        }
      }
      default
      {
        Write-Warning "Unsupported update type: $updateType for device $fqdn"
        $failedOperations += [PSCustomObject]@{
          FQDN      = $fqdn
          Operation = "Device Update"
          Error     = "Unsupported update type: $updateType"
        }
      }
    }
  }
  catch
  {
    Write-Error "Failed to update device $fqdn`: $($_.Exception.Message)"
    $failedOperations += [PSCustomObject]@{
      FQDN      = $fqdn
      Operation = "Device Update"
      Error     = $_.Exception.Message
    }
  }
}

# Display summary of failures
if ($failedOperations.Count -gt 0)
{
  Write-Host "`n========================================" -ForegroundColor Red
  Write-Host "FAILED OPERATIONS SUMMARY" -ForegroundColor Red
  Write-Host "========================================" -ForegroundColor Red
  $failedOperations | Format-Table -AutoSize
}
else
{
  Write-Host "`n========================================" -ForegroundColor Green
  Write-Host "All operations completed successfully!" -ForegroundColor Green
  Write-Host "========================================" -ForegroundColor Green
}