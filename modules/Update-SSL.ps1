. "$PSScriptRoot\Write-Log.ps1"

function Update-SSL {
  [CmdletBinding()]
  param (
    [string]$SiteName,
    [object]$selectedBinding,
    [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate = $null,
    [switch]$SkipConfirmation
  )

  try {
    # 如果沒有提供憑證，嘗試獲取最新的憑證
    if (-not $Certificate) {
      $CertName = "$env:COMPUTERNAME"
      $Certificate = Get-ChildItem -Path Cert:\LocalMachine\My | 
        Where-Object { $_.Subject -like "CN=$CertName,*" } | 
        Sort-Object NotBefore -Descending | 
        Select-Object -First 1
      
      if (-not $Certificate) {
        throw "無法找到最新申請的憑證"
      }
    }

    $Website = Get-Website -Name $SiteName
    if (-not $Website) {
      throw "找不到指定的站台: $SiteName"
    }

    # 更新 SSL 憑證
    Write-Log "準備更新 SSL 憑證" -Level Information

    $CertNameHash = $Certificate.Thumbprint

    Write-Log "開始更新 SSL 憑證..." -Level Warning
    Write-Log "現有繫結資訊: $($selectedBinding.bindingInformation)" -Level Information
    Write-Log "目前憑證指紋: $($selectedBinding.certificateHash)" -Level Information
    Write-Log "準備更新為新憑證指紋: $CertNameHash" -Level Information

    # 如果未跳過確認，則詢問使用者
    if (-not $SkipConfirmation) {
      $confirm = Read-Host "是否要修改此繫結的 SSL 證書？(Y/N)"

      if (-not ([string]::IsNullOrWhiteSpace($confirm) -or $confirm -match '^[Yy]$')) {
        if ($confirm -match '^[Nn]$') {
          Write-Log "使用者取消操作" -Level Warning
          throw "使用者取消操作"
        } else {
          Write-Log "無效的輸入" -Level Error
          throw "請輸入 Y 或 N"
        }
      }
    }
    
    Write-Log "正在更新 SSL 憑證" -Level Information
    
    # 使用原始的繫結資訊
    $bindingInfo = $selectedBinding.bindingInformation -split ':'
    $Port = $bindingInfo[1]
    $HostHeaderName = $bindingInfo[2]
    
    # 構建繫結字串
    $BindingInfo = "*:${Port}:${HostHeaderName}"
    
    # 取得現有繫結
    $ExistingBinding = Get-WebBinding -Name $SiteName -Protocol "https" |
    Where-Object { $_.bindingInformation -eq $BindingInfo }
  
    if ($ExistingBinding) {
      # 修改現有繫結的憑證
      $ExistingBinding.RemoveSslCertificate()
      $ExistingBinding.AddSslCertificate($Certificate.Thumbprint, "my")
  
      # 驗證更新
      Start-Sleep -Seconds 1
      $updatedBinding = Get-WebBinding -Name $SiteName -Protocol "https" |
      Where-Object { $_.bindingInformation -eq $BindingInfo }
  
      if ($updatedBinding.certificateHash -eq $Certificate.Thumbprint) {
        Write-Log "SSL 憑證更新成功" -Level Success
        return $true
      }
      else {
        throw "SSL 憑證更新失敗：更新後的指紋與預期不符"
      }
    }
    else {
      throw "找不到指定的 HTTPS 繫結: $BindingInfo"
    }
  }
  catch {
    Write-Log "更新 SSL 憑證失敗: $($_.Exception.Message)" -Level Error
    throw "SSL 憑證更新失敗"
  }
}