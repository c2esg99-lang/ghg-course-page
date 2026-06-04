# GHG Protocol 課程頁面

## 啟動頁面與後端

在這個資料夾執行：

```powershell
node server.js
```

開啟：

```text
http://localhost:3000
```

報名表會送到後端 `/api/register`，並保存到：

```text
data/registrations.jsonl
```

## 設定自動寄信

若要讓後端自動把報名資料寄到 `c2esg99@gmail.com`，請先設定 SMTP 環境變數。

Gmail 範例：

```powershell
$env:SMTP_HOST="smtp.gmail.com"
$env:SMTP_PORT="587"
$env:SMTP_USER="寄件用 Gmail 帳號"
$env:SMTP_PASS="Gmail App Password"
$env:MAIL_FROM="寄件用 Gmail 帳號"
$env:MAIL_TO="c2esg99@gmail.com"
node server.js
```

Gmail 需要使用 App Password，不能使用一般 Gmail 登入密碼。
