. "$PSScriptRoot\Write-Log.ps1"

# 選擇繫結
function Select-Binding {
  param (
    [string]$SiteName
  )

  function Show-Menu {
    param (
      [array]$bindings,
      [int]$selectedIndex
    )

    foreach ($i in 0..($bindings.Count - 1)) {
      $binding = $bindings[$i]
      $info = $binding.bindingInformation -split ':'
      $bindingIp = if ($info[0] -eq '*') { '未指定的 IP 位址' }else { $info[0] }
      $bindingPort = $info[1]
      $bindingHost = if ([string]::IsNullOrEmpty($info[2])) { '*' }else { $info[2] }

      if ($i -eq $selectedIndex) {
        Write-Host ">" -NoNewline -ForegroundColor Green
        Write-Host " [$i] HTTPS 繫結" -ForegroundColor Cyan
        Write-Host "     IP 位址: $bindingIp" -ForegroundColor Green
        Write-Host "     通訊埠: $bindingPort" -ForegroundColor Green
        Write-Host "     主機名稱: $bindingHost" -ForegroundColor Green
      }
      else {
        Write-Host "  [$i] HTTPS 繫結" -ForegroundColor Gray
        Write-Host "     IP 位址: $bindingIp" -ForegroundColor DarkGray
        Write-Host "     通訊埠: $bindingPort" -ForegroundColor DarkGray
        Write-Host "     主機名稱: $bindingHost" -ForegroundColor DarkGray
      }
      Write-Host ""
    }
  }

  function Get-CertificateInfo {
    param (
      [string]$certificateStoreName,
      [string]$certificateHash
    )

    $certInfo = Get-ChildItem -Path Cert:\LocalMachine\$certificateStoreName\$certificateHash -ErrorAction SilentlyContinue
    if ($certInfo) {
      $certName = if ($certInfo.Subject) {
        if ($certInfo.Subject -match "CN=([^,]+)") {
          $matches[1]
        }
        else {
          $certInfo.Subject
        }
      }
      else {
        "未知"
      }
      
      return @{
        Name           = $certName
        ExpirationDate = $certInfo.NotAfter
        Thumbprint     = $certInfo.Thumbprint
      }
    }
    else {
      return @{
        Name           = "未知"
        ExpirationDate = $null
        Thumbprint     = $certificateHash
      }
    }
  }

  Write-Log "正在檢查站台 '$SiteName' 的 HTTPS 繫結..." -Level Information
  Write-Log "請使用上下方向鍵選擇站台繫結，按 Enter 確認，按 Esc 取消退出：" -Level Information
  Write-Host ""
  Write-Host "----------------------------------------" -ForegroundColor DarkGray

  $allBindings = Get-WebBinding -Name $SiteName -Protocol "https"
  if (-not $allBindings -or $allBindings.Count -eq 0) {
    Write-Log "未找到 HTTPS 繫結" -Level Warning
    $retry = Read-Host "是否要選擇其他站臺？(Y/N)"
    return $null
  }

  # 過濾特定 Port 的繫結項目
  $bindings = $allBindings
  # 確保 $bindings 為陣列
  if ($bindings -and -not ($bindings -is [Array])) {
    $bindings = @($bindings)
  }
  Write-Log "找到 $($bindings.Count) 個符合的繫結" -Level Information
  
  if (-not $bindings -or $bindings.Count -eq 0) {
    Write-Log "未找到符合通訊埠 $Port 的 HTTPS 繫結" -Level Warning
    $retry = Read-Host "是否要選擇其他站臺？(Y/N)"
    if ([string]::IsNullOrWhiteSpace($retry) -or $retry -match '^[Yy]$') {
      return $null
    }
    else {
      return $null
    }
  }

  $selectedIndex = 0
  $selectionBinding = $null

  $menuStart = $host.UI.RawUI.CursorPosition
  
  do {
    $host.UI.RawUI.CursorPosition = $menuStart

    $menuHeight = ($bindings.Count * 5) + 2
    1..$menuHeight | ForEach-Object {
      Write-Host (" " * $host.UI.RawUI.BufferSize.Width)
    }
    
    $host.UI.RawUI.CursorPosition = $menuStart
    Show-Menu $bindings $selectedIndex
    
    $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

    switch ($key.VirtualKeyCode) {
      38 {
        # 上
        if ($selectedIndex -gt 0) { $selectedIndex-- }
      }
      40 {
        # 下
        if ($selectedIndex -lt ($bindings.Count - 1)) { $selectedIndex++ }
      }
      13 {
        # Enter
        if ($bindings.Count -eq 1) {
          $selectionBinding = $bindings
        }
        else {
          $selectionBinding = $bindings[$selectedIndex]
        }
        
        if ($selectionBinding) {      
          $bindingInformation = $selectionBinding.bindingInformation
          $bindingInfo = $bindingInformation -split ':'
          
          # 使用新函數獲取證書信息
          $certInfo = Get-CertificateInfo -certificateStoreName $selectionBinding.certificateStoreName -certificateHash $selectionBinding.certificateHash

          # Write-Log "選擇的繫結資訊: $($bindings[0] | ConvertTo-Json)" -Level Information
          Write-Log "已選擇修改站台 $SiteName 的 https 繫結：" -Level Success
          Write-Log "  IP 位址: $(if($bindingInfo[0] -eq '*'){'未指定 IP 位址'}else{$bindingInfo[0]})" -Level Information
          Write-Log "  通訊埠(Port): $($bindingInfo[1])" -Level Information
          Write-Log "  主機名稱(Host): $(if([string]::IsNullOrEmpty($bindingInfo[2])){'所有主機名稱'}else{$bindingInfo[2]})" -Level Information
          Write-Log "  完整繫結資訊: $bindingInformation" -Level Information
          Write-Log "  憑證名稱: $($certInfo.Name)" -Level Information
          Write-Log "  憑證到期日: $($certInfo.ExpirationDate)" -Level Information
          Write-Log "  憑證指紋: $($certInfo.Thumbprint)" -Level Information
        }
        else {
          Write-Log "無法獲取選擇的繫結資訊，請重試" -Level Warning
        }
        $confirm = Read-Host "是否確定要編輯此繫結 SSL 證書？(Y/N)"
        if (-not ([string]::IsNullOrWhiteSpace($confirm) -or $confirm -match '^[Yy]$')) {
          if ($confirm -match '^[Nn]$') {
            Write-Log "使用者取消操作" -Level Warning
            return $False
          }
          else {
            Write-Log "無效的輸入" -Level Error
            return $False
          }
        }
        return $selectionBinding
      }
      27 {
        # ESC
        throw "使用者取消操作"
      }
    }
  } while ($null -eq $selectionBinding)
}