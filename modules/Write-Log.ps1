# log輸出設定
function Write-Log {
  param(
    [string]$Message,
    [ValidateSet('Information','Warning','Error','Success')]
    [string]$Level = 'Information'
  )
  
  $ColorMap = @{
    'Information' = 'White'
    'Warning' = 'Yellow'
    'Error' = 'Red'
    'Success' = 'Green'
  }
  
  $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Write-Host "[$TimeStamp] [$Level] $Message" -ForegroundColor $ColorMap[$Level]
}