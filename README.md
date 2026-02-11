# 伊蓮娜的煩惱 (ResumeMailer)

Windows 履歷寄送工具。透過本機 Outlook 自動寄出兩封信：
1. 第一封：加密壓縮檔附件
2. 第二封：壓縮檔密碼

## 功能
- Windows GUI (Tkinter)
- Outlook COM 寄信
- 多帳號寄件，並記住預設寄件帳號
- 8 位數字密碼 + AES ZIP
- 寄送前先驗證壓縮檔可解
- 啟動畫面 (splash): `photo/elena.png`

## 快速使用
```powershell
.\bootstrap.ps1
.\run.ps1
```

## 打包 EXE
```powershell
.\build.ps1
```

輸出：
- `release\ResumeMailer\ResumeMailer.exe`

## 設定與紀錄存放位置
程式會存到 EXE 同目錄：
- `config.json`: `<exe-folder>\config.json`
- `audit.jsonl`: `<exe-folder>\logs\audit.jsonl`
- 寄送用 zip 暫存：`<exe-folder>\outbox\`

## 必要條件
- Windows 10/11
- 已安裝並登入 Outlook 桌面版

## 注意
- 收件端建議用 7-Zip 或 WinRAR 解 AES ZIP。
- 若重打包失敗且提示檔案被占用，先關閉 `ResumeMailer.exe` 再執行 `.\build.ps1`。
