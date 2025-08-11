. "$PSScriptRoot\Write-Log.ps1"

# 選擇站台
function Select-Website {
  function Show-Menu { # 顯示站台清單
    param (
      [array]$sites,
      [int]$selectedIndex
    )
    
    foreach ($i in 0..($sites.Count-1)) {
      $site = $sites[$i]
      if ($i -eq $selectedIndex) {
        Write-Host ">" -NoNewline -ForegroundColor Green
        Write-Host " [$i] $($site.Name)" -ForegroundColor Cyan
        Write-Host "     狀態: $($site.State)" -ForegroundColor Green
        Write-Host "     實體路徑: $($site.PhysicalPath)" -ForegroundColor Green
      } else {
        Write-Host "  [$i] $($site.Name)" -ForegroundColor Gray
        Write-Host "     狀態: $($site.State)" -ForegroundColor DarkGray
        Write-Host "     實體路徑: $($site.PhysicalPath)" -ForegroundColor DarkGray
      }
      Write-Host ""
    }
  }

  do {
    Write-Log "請使用上下方向鍵選擇站台，按 Enter 確認，按 Esc 取消退出：" -Level Information
    Write-Host ""
    Write-Host "----------------------------------------" -ForegroundColor DarkGray

    $sites = Get-Website
    $selectedIndex = 0
    $selectionSite = $null
    $continueSelection = $true
    $confirmSite = $false
    
    $menuStart = $host.UI.RawUI.CursorPosition
    
    while ($continueSelection) {
      $host.UI.RawUI.CursorPosition = $menuStart
      
      $menuHeight = ($sites.Count * 4) + 1
      1..$menuHeight | ForEach-Object {
        Write-Host (" " * $host.UI.RawUI.BufferSize.Width)
      }
      
      $host.UI.RawUI.CursorPosition = $menuStart
      Show-Menu $sites $selectedIndex

      $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

      switch ($key.VirtualKeyCode) {
        38 { # 上箭頭
          if ($selectedIndex -gt 0) { $selectedIndex-- }
        }
        40 { # 下箭頭
          if ($selectedIndex -lt ($sites.Count - 1)) { $selectedIndex++ }
        }
        13 { # Enter
          $selectedSite = $sites[$selectedIndex]
          $selectionSite = $selectedSite.Name
          $continueSelection = $false
        }
        27 { # Esc
          throw "使用者取消操作"
        }
      }
    }
    

    $confirm = Read-Host "是否確定要為此站台設定 HTTPS 繫結？(Y/N)"
    if ([string]::IsNullOrWhiteSpace($confirm) -or $confirm -match '^[Yy]$') {
      Write-Log "已選擇站台: $selectionSite" -Level Success
      $confirmSite = $true
      return $selectionSite
    }
    elseif ($confirm -match '^[Nn]$') {
      Write-Log "返回站台選擇..." -Level Information
      Clear-Host
      $selectedIndex = 0
      $selectionSite = $null
      $continueSelection = $true
    }
    else {
      Write-Log "請輸入 Y 或 N" -Level Warning
      $selectedIndex = 0
      $selectionSite = $null
      $continueSelection = $true
    }
  } while (-not $confirmSite)
}