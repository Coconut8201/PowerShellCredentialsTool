# 導入模組
Import-Module "$PSScriptRoot\modules\CertificateTools.psm1" -Force
Import-Module "$PSScriptRoot\modules\Update-SSL.ps1" -Force

function Main {
  [CmdletBinding()]
  # 憑證名稱(電腦名稱)
  $CertName = "$env:COMPUTERNAME"

  # IP地址
  $IPAddress = (Get-NetIPAddress -AddressFamily IPv4 -InterfaceAlias "Ethernet0" | Where-Object { $_.PrefixOrigin -ne "WellKnown" }).IPAddress

  # 檢查管理員權限
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $isAdmin) {
    Write-Log "此程序需要系統管理員權限才能執行!" -Level Error
    Write-Log "請以系統管理員身份重新開啟 PowerShell 再執行此程序" -Level Warning
    Write-Host "按任意鍵結束..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
  }

  Write-Log "系統管理員權限確認完成，開始執行程序..." -Level Success
  Write-Log "------------------------------------------" -Level Information

  try {
    if (-not (Get-Module -Name WebAdministration)) {
      Import-Module WebAdministration
      Write-Log "已導入 WebAdministration 模組" -Level Information
    }

    do {
      # 選擇站台
      if (-not $SiteName) {
        $SiteName = Select-Website
      }
      
      # 選擇繫結
      $selectedBinding = Select-Binding -SiteName $SiteName
      if ($null -eq $selectedBinding) {
        Write-Log "未找到有效的 HTTPS 繫結，返回站台選擇" -Level Warning
        $SiteName = $null
        Clear-Host
      }
      if ($selectedBinding -eq $False) {
        Write-Log "重新選擇繫結" -Level Warning
        Clear-Host
      }
    } while (-not $SiteName -or $null -eq $selectedBinding -or $selectedBinding -eq $False)

    $params = @{
      CertName  = $CertName
      IPAddress = $IPAddress
      SiteName  = $SiteName
    }
    Write-Log "------------------------------------------" -Level Information
    Write-Log "即將使用以下資訊生成憑證：" -Level Information
    Write-Log "  憑證名稱: $CertName" -Level Information
    Write-Log "  IP位址: $IPAddress" -Level Information 
    Write-Log "  站台名稱: $SiteName" -Level Information
    Write-Host ""
    
    Write-Log "開始生成憑證..." -Level Information
    # 生成 SSL 憑證
    $NewCertificate = New-Certificate @params
    
    if (-not $NewCertificate) {
      throw "憑證生成失敗"
    }
    
    Write-Log "憑證生成完成" -Level Success

    Write-Log "已生成憑證 $NewCertificate" -Level Information
    $Certificate = Get-ChildItem -Path Cert:\LocalMachine\My | 
    Where-Object { $_.Subject -like "CN=$CertName,*" } | 
    Sort-Object NotBefore -Descending | 
    Select-Object -First 1

    if (-not $Certificate) {
      throw "無法找到最新申請的憑證"
    }

    # 顯示憑證詳細資訊
    Write-Log "------------------------------------------" -Level Information
    Write-Log "憑證詳細資訊：" -Level Information
    # 從 Subject 中提取 CN 後的名稱
    $cnName = if ($Certificate.Subject -match "CN=([^,]+)") { $Matches[1] } else { "未知" }
    Write-Log "  名稱: $cnName" -Level Information
    Write-Log "  發行者: $($Certificate.Issuer)" -Level Information
    Write-Log "  序號: $($Certificate.SerialNumber)" -Level Information
    Write-Log "  生效日期: $($Certificate.NotBefore)" -Level Information
    Write-Log "  到期日期: $($Certificate.NotAfter)" -Level Information
    Write-Log "  指紋: $($Certificate.Thumbprint)" -Level Information
    Write-Log "------------------------------------------" -Level Information
    
    # 詢問使用者是否確認替換憑證
    $confirmReplace = Read-Host "是否確認使用此憑證替換現有憑證？(Y/N)"
    if (-not ([string]::IsNullOrWhiteSpace($confirmReplace) -or $confirmReplace -match '^[Yy]$')) {
      if ($confirmReplace -match '^[Nn]$') {
        Write-Log "使用者取消替換憑證" -Level Warning
        throw "使用者取消替換憑證"
      }
      else {
        Write-Log "無效的輸入" -Level Error
        throw "請輸入 Y 或 N"
      }
    }

    # 更新 SSL 憑證
    $params = @{
      SiteName         = $SiteName
      selectedBinding  = $selectedBinding
      Certificate      = $Certificate
      SkipConfirmation = $true
    }
    
    $result = Update-SSL @params
    
    if ($result) {
      Write-Log "成功修改 HTTPS 繫結和憑證" -Level Success
      Write-Log "所有操作已完成" -Level Success
      Write-Log "按任意鍵結束..." -Level Information
      $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
  }
  catch {
    Write-Log $_.Exception.Message -Level Error
    Write-Log "堆疊追蹤: $($_.ScriptStackTrace)" -Level Error
    Write-Log "程式執行失敗，按任意鍵結束..." -Level Error
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
  }
}

Main @PSBoundParameters