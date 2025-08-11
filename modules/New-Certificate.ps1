. "$PSScriptRoot\Write-Log.ps1"

# 生成憑證
function New-Certificate {
  [CmdletBinding()]
  param (
    [string]$CertName,
    [string]$IPAddress,
    [string]$SiteName,
    [bool]$CleanupFiles = $false
  )
  Write-Log "正在呼叫憑證生成程式..." -Level Information
  Write-Log "Parameters:" -Level Information
  Write-Log "  CertName: [$CertName]" -Level Information
  Write-Log "  IPAddress: [$IPAddress]" -Level Information
  Write-Log "  SiteName: [$SiteName]" -Level Information
  Write-Log "  CleanupFiles: [$CleanupFiles]" -Level Information

  if ([string]::IsNullOrEmpty($CertName)) {
    Write-Log "Error: Certificate name not provided" -Level Error
    throw "Certificate name cannot be empty"
  }

  $CSRPath = "$PSScriptRoot\Certificate\$($CertName).csr"
  $INFPath = "$PSScriptRoot\Certificate\$($CertName).inf"
  $CERPath = "$PSScriptRoot\Certificate\$($CertName).cer"

  try {
    $testPath = Split-Path $CSRPath -Parent
    Write-Log "Certificate directory path: $testPath" -Level Information
    
    if (!(Test-Path $testPath)) {
      Write-Log "Certificate directory does not exist, attempting to create it..." -Level Warning
      try {
        New-Item -ItemType Directory -Path $testPath -Force -ErrorAction Stop | Out-Null
        Write-Log "Certificate directory created successfully" -Level Success
      }
      catch {
        Write-Log "Failed to create certificate directory: $($_.Exception.Message)" -Level Error
        throw "無法創建憑證目錄: $testPath - $($_.Exception.Message)"
      }
    }

    Write-Log "Starting certificate generation: $CertName" -Level Information
    
    $Signature = '$Windows NT$'
    $FullDomainName = "$CertName.server.com"

    $INF = @"
[Version]
Signature= "$Signature" 

[NewRequest]
Subject = "CN=$CertName, OU=Server, O=SYSTEX, L=Taiwan, S=TW, C=TW"
KeySpec = 1
KeyLength = 2048
Exportable = TRUE
MachineKeySet = TRUE
PrivateKeyArchive = FALSE
UserProtected = FALSE
UseExistingKeySet = FALSE
ProviderName = "Microsoft RSA SChannel Cryptographic Provider"
ProviderType = 12
RequestType = CMC
KeyUsage = 0xa0

[EnhancedKeyUsageExtension]

OID=1.3.6.1.5.5.7.3.1  ; Server Authentication
[Extensions]

; If your client operating system is Windows Server 2008, Windows Server 2008 R2, Windows Vista, or Windows 7

; SANs can be included in the Extensions section by using the following text format. Note 2.5.29.17 is the OID for a SAN extension.
2.5.29.17 = "{text}"
_continue_ = "dns=$CertName&"
_continue_ = "dns=$FullDomainName&"   
"@

    if ($IPAddress -ne "") {
      $INF += "`n_continue_ = `"IPAddress=$IPAddress&`""
    }

    $INF += @"
; If your client operating system is Windows Server 2003, Windows Server 2003 R2, or Windows XP
; SANs can be included in the Extensions section only by adding Base64-encoded text containing the alternative names in ASN.1 format.
; Use the provided script MakeSanExt.vbs to generate a SAN extension in this format.
; RMILNE – the below line is remmed out else we get an error since there are duplicate sections for OID 2.5.29.17
; 2.5.29.17=MCaCEnd3dzAxLmZhYnJpa2FtLmNvbYIQd3d3LmZhYnJpa2FtLmNvbQ

[RequestAttributes]
; If your client operating system is Windows Server 2003, Windows Server 2003 R2, or Windows XP
; and you are using a standalone CA, SANs can be included in the RequestAttributes
; section by using the following text format.
;"SAN="dns=$CertName&dns=$FullDomainName&IPAddress=$IPAddress"

; Multiple alternative names must be separated by an ampersand (&).
CertificateTemplate = WebServer ; Modify for your environment by using the LDAP common name of the template.

;Required only for enterprise CAs.

"@

    Write-Log "Generating INF file at path: $INFPath" -Level Warning
    try {
      $infDir = Split-Path $INFPath -Parent
      if (!(Test-Path $infDir)) {
        New-Item -ItemType Directory -Path $infDir -Force -ErrorAction Stop | Out-Null
        Write-Log "Created directory for INF file: $infDir" -Level Information
      }
      
      # 寫入INF檔案
      # $INF | Out-File -FilePath $INFPath -Encoding utf8NoBOM -Force -ErrorAction Stop # powershell 7.0+ 支援 utf8NoBOM
      [System.IO.File]::WriteAllText($INFPath, $INF, [System.Text.UTF8Encoding]::new($false))
      Write-Log "INF file written successfully" -Level Information
    }
    catch {
      Write-Log "Failed to write INF file: $($_.Exception.Message)" -Level Error
      Write-Log "Error details: $($_.Exception | Out-String)" -Level Error
      throw "無法創建INF檔案: $($_.Exception.Message)"
    }
    
    Write-Log "Generating CSR..." -Level Warning
    $csrResult = certreq -new $INFPath $CSRPath 2>&1
    Write-Log "CSR command result: $csrResult" -Level Information
    if ($LASTEXITCODE -ne 0) { 
      Write-Log "CSR generation failed: $csrResult" -Level Error
      throw "CSR生成失敗: $csrResult" 
    }
    Write-Log "CSR generated successfully" -Level Success

    Write-Log "Submitting CSR to CA..." -Level Warning
    $submitResult = certreq -submit -config "AD-server" $CSRPath $CERPath 2>&1
    if ($LASTEXITCODE -ne 0) { 
      Write-Log "CSR submission failed: $submitResult" -Level Error
      throw "CSR提交失敗: $submitResult" 
    }
    Write-Log "CSR submitted successfully" -Level Success

    Write-Log "Installing certificate..." -Level Warning
    $acceptResult = certreq -accept $CERPath 2>&1
    if ($LASTEXITCODE -ne 0) { 
      Write-Log "Certificate installation failed: $acceptResult" -Level Error
      throw "憑證安裝失敗: $acceptResult" 
    }
    Write-Log "Certificate installed successfully" -Level Success

    Start-Sleep -Seconds 5

    Write-Log "Searching for newly installed certificate..." -Level Information
    $Certificate = Get-ChildItem -Path Cert:\LocalMachine\My | 
    Where-Object { $_.Subject -like "CN=$CertName,*" } | 
    Sort-Object NotBefore -Descending | 
    Select-Object -First 1

    if (-not $Certificate) {
      Write-Log "Cannot find the newly installed certificate" -Level Error
      throw "找不到新安裝的憑證，請檢查憑證存儲區"
    }

    Write-Log "Certificate generated and installed successfully" -Level Success
    Write-Log "Certificate thumbprint: $($Certificate.Thumbprint)" -Level Information
    Write-Log "Certificate subject: $($Certificate.Subject)" -Level Information
    Write-Log "Certificate validity: $($Certificate.NotBefore) to $($Certificate.NotAfter)" -Level Information

    return $Certificate.Thumbprint

  }
  catch {
    Write-Log "Error during certificate generation: $_" -Level Error
    Write-Log "Error type: $($_.Exception.GetType().FullName)" -Level Error
    Write-Log "Error message: $($_.Exception.Message)" -Level Error
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level Error
    throw
  }
  finally {
    if ($CleanupFiles) {
      Write-Log "Cleanup mode enabled, will remove temporary files" -Level Information
      @($CSRPath, $INFPath, $CERPath) | ForEach-Object {
        if (Test-Path $_) {
          try {
            Remove-Item $_ -Force
            Write-Log "Cleaned up file: $_" -Level Information
          }
          catch {
            Write-Log "Failed to clean up file: $_ - $($_.Exception.Message)" -Level Warning
          }
        }
      }
    }
    else {
      Write-Log "Temporary certificate files are preserved at:" -Level Information
      @($CSRPath, $INFPath, $CERPath) | ForEach-Object {
        if (Test-Path $_) {
          Write-Log "  $_" -Level Information
        }
      }
    }
  }
}

if ($MyInvocation.InvocationName -ne ".") {
  try {
    if ([string]::IsNullOrEmpty($CertName)) {
      $CertName = "$env:COMPUTERNAME"
      Write-Log "Using default certificate name: $CertName" -Level Warning
    }

    $params = @{
      CertName     = $CertName
      CleanupFiles = $false  # 預設保留臨時檔案
    }

    if ($IPAddress) { $params['IPAddress'] = $IPAddress }
    if ($SiteName) { $params['SiteName'] = $SiteName }
    if ($PSBoundParameters.ContainsKey('CleanupFiles')) { $params['CleanupFiles'] = $CleanupFiles }

    $result = New-Certificate @params
    Write-Log "Certificate generation completed, thumbprint: $result" -Level Success
  }
  catch {
    Write-Log "Execution failed: $($_.Exception.Message)" -Level Error
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level Error
    exit 1
  }
}