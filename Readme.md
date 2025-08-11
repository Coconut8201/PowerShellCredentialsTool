# PowerShell SSL 憑證替換工具

這是一個用於自動替換 IIS 網站 SSL 憑證的 PowerShell 工具，可從 AD Server 取得新憑證。

## 概述

此工具自動化從 AD Server 申請新 SSL 憑證並更新 IIS 網站繫結的流程。它可延長憑證有效期 2 年，並將憑證儲存在本地 Certificate 資料夾中。

## 程式流程圖

![流程圖](/process.png)

## 系統需求

- **需要系統管理員權限** - 必須以系統管理員身分執行 PowerShell
- 具有現有 HTTPS 繫結的 IIS 伺服器
- 可存取 AD Server 進行憑證申請
- 7-Zip 用於解壓縮工具檔案

## 安裝說明

1. 下載 `PowerShellCredentialsTool.7z` 檔案
2. 使用 7-Zip 解壓縮，取得 `PowerShellCredentialsTool` 資料夾
3. 記下解壓縮路徑（例如：`C:\Users\使用者名稱\Desktop\PowerShellCredentialsTool`）

## 使用方式

本工具提供兩種啟動方式，請根據您的需求選擇適合的方式：

### 方式一：直接執行（建議一般使用者）

適合直接使用，無需修改程式碼的情況。

#### 步驟 1：以系統管理員身分啟動 PowerShell
> [!IMPORTANT]
> **必須以系統管理員權限執行 PowerShell**

以系統管理員身分開啟 PowerShell，並導航至 PowerShellCredentialsTool 目錄：

```powershell
cd "C:\路徑\至\PowerShellCredentialsTool"
```

#### 步驟 2：執行主要腳本

執行單一主要檔案：

```powershell
.\Main.ps1
```

### 方式二：模組化執行（建議開發者或需要維護）

適合需要修改程式內容、進行維護或自訂功能的情況。

#### 步驟 1：以系統管理員身分啟動 PowerShell
> [!IMPORTANT]
> **必須以系統管理員權限執行 PowerShell**

以系統管理員身分開啟 PowerShell，並導航至 PowerShellCredentialsTool 目錄：

```powershell
cd "C:\路徑\至\PowerShellCredentialsTool"
```

#### 步驟 2：載入模組並執行

載入模組化檔案並執行：

```powershell
# 導入主要模組
Import-Module ".\modules\CertificateTools.psm1" -Force

# 執行主要功能（可根據需要自訂參數）
# 這裡可以個別調用各功能模組進行自訂操作
```

## 操作流程

### 步驟 1：選擇目標網站

從可用選項中選擇您要修改的 IIS 網站。

> [!NOTE]
> 如果所選網站沒有 HTTPS 繫結，工具將顯示「**未找到 HTTPS 繫結**」並提示您重新選擇網站。

### 步驟 2：選擇 HTTPS 繫結

選擇網站後，選擇您要更新的特定 HTTPS 繫結。這相當於在 IIS 介面中編輯繫結。

工具會要求確認後才會繼續進行 SSL 憑證修改。

### 步驟 3：憑證申請

工具將會：
- 使用主機資訊（主機名稱、IP 位址）向 AD Server 申請新的 SSL 憑證
- 將申請的憑證儲存在 PowerShellCredentialsTool 目錄內的 `Certificate` 資料夾中
- 顯示憑證資訊並詢問是否確認更新 SSL 憑證

### 步驟 4：憑證安裝

確認後，工具將會：
- 替換所選繫結的 SSL 憑證
- 使用新的 2 年有效期更新憑證
- 成功完成程序

## 執行結果

成功執行後，您將看到所選繫結現在使用新申請的憑證，並具有延長的 2 年到期日期。

## 資料夾結構

```
PowerShellCredentialsTool/
├── Main.ps1                    # 主要執行檔案（方式一使用）
├── modules/                    # 模組檔案資料夾（方式二使用）
│   ├── CertificateTools.psm1   # 主要模組檔案
│   ├── New-Certificate.ps1     # 憑證申請功能
│   ├── Select-Binding.ps1      # 繫結選擇功能
│   ├── Select-Website.ps1      # 網站選擇功能
│   ├── Update-SSL.ps1          # SSL 更新功能
│   └── Write-Log.ps1           # 日誌記錄功能
├── Certificate/                # 儲存申請的憑證
└── process.png                 # 程式流程圖
```

## 使用場景建議

- **方式一（Main.ps1）**：適合系統管理員日常使用，快速執行憑證更新作業
- **方式二（modules）**：適合開發人員進行程式維護、功能修改或整合其他系統

## 技術支援

如有關於此工具的問題或疑問，請聯繫您的系統管理員或 IT 支援團隊。