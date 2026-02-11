# 伊蓮娜的煩惱 (ResumeMailer)

Windows 履歷寄送工具，透過本機 Outlook 自動寄出兩封信：
1. 第一封：加密壓縮檔附件
2. 第二封：壓縮檔密碼

## 下載
- [下載 Release/ResumeMailer.zip](./release/ResumeMailer.zip)

## 快速使用
```powershell
.\bootstrap.ps1
.\run.ps1
```

## 打包
```powershell
.\build.ps1
```

輸出：
- `release\ResumeMailer\ResumeMailer.exe`
- `release\ResumeMailer.exe`
- `release\ResumeMailer.zip`

## 版本號規則
- 版本檔：`VERSION`
- 起始版本：`v0.9.0`
- 每次執行 `.\build.ps1` 會自動遞增 patch
- 例如：`v0.9.0 -> v0.9.1 -> v0.9.2`

## 設定與暫存位置（與 EXE 同目錄）
- 設定檔：`config.json`
- 紀錄檔：`logs\audit.jsonl`
- 上傳暫存：`uploads\`
- 壓縮暫存：`outbox\`

每次送件後會刪除當次的上傳暫存與壓縮暫存檔案。

## 必要條件
- Windows 10/11
- 已安裝且可登入 Outlook 桌面版

## 注意
- 收件端建議使用 7-Zip 或 WinRAR 解 AES ZIP。
- 若打包失敗且提示檔案被占用，先關閉 `ResumeMailer.exe` 再重試。
