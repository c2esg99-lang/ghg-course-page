const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const net = require("node:net");
const tls = require("node:tls");

const root = __dirname;
const port = Number(process.env.PORT || 3000);
const recipient = process.env.MAIL_TO || "c2esg99@gmail.com";
const dataDir = path.join(root, "報名資料");
const submissionsFile = path.join(dataDir, "registrations.jsonl");
const csvFile = path.join(dataDir, "報名資料.csv");
const xlsFile = path.join(dataDir, "報名資料.xls");
const csvFields = [
  "報名時間",
  "報名編號",
  "報名梯次",
  "報名優惠方案",
  "姓名",
  "手機",
  "Email",
  "LINE ID",
  "公司 / 單位",
  "職稱",
  "收據 / 發票需求",
  "發票公司名稱",
  "統一編號",
  "發票地址",
  "飲食需求",
  "繳費狀態",
  "轉帳後五碼",
  "備註"
];

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon"
};

function sendJson(response, statusCode, body) {
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  });
  response.end(JSON.stringify(body));
}

function readRequestBody(request) {
  return new Promise((resolve, reject) => {
    let body = "";
    request.on("data", (chunk) => {
      body += chunk;
      if (body.length > 1024 * 1024) {
        request.destroy();
        reject(new Error("資料量過大"));
      }
    });
    request.on("end", () => resolve(body));
    request.on("error", reject);
  });
}

function sanitizeText(value) {
  return String(value || "").replace(/\r?\n/g, " ").trim();
}

function validateRegistration(data) {
  const requiredFields = ["報名梯次", "報名優惠方案", "姓名", "手機", "Email", "繳費狀態"];
  const missing = requiredFields.filter((field) => !sanitizeText(data[field]));
  if (data["收據 / 發票需求"] === "需要開發票") {
    ["發票公司名稱", "統一編號", "發票地址"].forEach((field) => {
      if (!sanitizeText(data[field])) missing.push(field);
    });
  }
  if (data["繳費狀態"] === "已繳費" && !/^\d{5}$/.test(sanitizeText(data["轉帳後五碼"]))) {
    missing.push("轉帳後五碼");
  }
  return missing;
}

function formatRegistrationEmail(record) {
  const fields = [
    "報名梯次",
    "報名優惠方案",
    "姓名",
    "手機",
    "Email",
    "LINE ID",
    "公司 / 單位",
    "職稱",
    "收據 / 發票需求",
    "發票公司名稱",
    "統一編號",
    "發票地址",
    "飲食需求",
    "繳費狀態",
    "轉帳後五碼",
    "備註"
  ];

  const lines = [
    "GHG Protocol 溫室氣體盤查實算工作坊報名資料",
    `報名時間：${record.submittedAt}`,
    `報名編號：${record.id}`,
    ""
  ];

  fields.forEach((field) => {
    lines.push(`${field}：${record.data[field] || "未填寫"}`);
  });

  lines.push("");
  lines.push("繳費資訊");
  lines.push("帳戶：熹堍永續股份有限公司");
  lines.push("銀行：華南銀行（008）新市分行");
  lines.push("帳號：648-10-011120-1");

  return lines.join("\n");
}

function csvEscape(value) {
  return `"${String(value || "").replace(/"/g, '""')}"`;
}

function appendRegistrationCsv(record) {
  const row = csvFields.map((field) => {
    if (field === "報名時間") return csvEscape(record.submittedAt);
    if (field === "報名編號") return csvEscape(record.id);
    return csvEscape(record.data[field]);
  });

  if (!fs.existsSync(csvFile)) {
    fs.writeFileSync(csvFile, `\uFEFF${csvFields.map(csvEscape).join(",")}\n`, "utf8");
  }
  fs.appendFileSync(csvFile, `${row.join(",")}\n`, "utf8");
}

function htmlEscape(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function readSavedRegistrations() {
  if (!fs.existsSync(submissionsFile)) return [];
  return fs.readFileSync(submissionsFile, "utf8")
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => {
      try {
        return JSON.parse(line);
      } catch (error) {
        return null;
      }
    })
    .filter(Boolean);
}

function writeRegistrationXls(records) {
  const widths = {
    "報名時間": 210,
    "報名編號": 300,
    "報名梯次": 420,
    "報名方案": 260,
    "姓名": 140,
    "手機": 150,
    "Email": 260,
    "LINE ID": 160,
    "公司 / 單位": 220,
    "職稱": 160,
    "收據 / 發票需求": 180,
    "發票公司名稱": 240,
    "統一編號": 140,
    "發票地址": 320,
    "飲食需求": 140,
    "繳費狀態": 130,
    "轉帳後五碼": 130,
    "備註": 360
  };

  const headerCells = csvFields.map((field) => {
    const width = widths[field] || 160;
    return `<th style="width:${width}px">${htmlEscape(field)}</th>`;
  }).join("");

  const bodyRows = records.map((record) => {
    const rowCells = csvFields.map((field) => {
      let value = record.data[field] || "";
      if (field === "報名時間") value = record.submittedAt;
      if (field === "報名編號") value = record.id;
      return `<td>${htmlEscape(value)}</td>`;
    }).join("");
    return `<tr>${rowCells}</tr>`;
  }).join("\n");

  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: "Microsoft JhengHei", Arial, sans-serif; }
    table { border-collapse: collapse; table-layout: fixed; }
    th {
      background: #185434;
      color: #ffffff;
      font-weight: 700;
      text-align: left;
    }
    th, td {
      border: 1px solid #b9cdb5;
      padding: 8px 10px;
      vertical-align: top;
      mso-number-format: "\\@";
      white-space: normal;
      word-break: break-word;
    }
    td { background: #fffdf8; }
  </style>
</head>
<body>
  <table>
    <thead><tr>${headerCells}</tr></thead>
    <tbody>${bodyRows}</tbody>
  </table>
</body>
</html>`;

  fs.writeFileSync(xlsFile, html, "utf8");
}

function smtpRead(socket) {
  return new Promise((resolve, reject) => {
    let buffer = "";
    const onData = (chunk) => {
      buffer += chunk.toString("utf8");
      const lines = buffer.split(/\r?\n/).filter(Boolean);
      const lastLine = lines[lines.length - 1] || "";
      if (/^\d{3} /.test(lastLine)) {
        socket.off("data", onData);
        socket.off("error", onError);
        resolve(buffer);
      }
    };
    const onError = (error) => {
      socket.off("data", onData);
      reject(error);
    };
    socket.on("data", onData);
    socket.on("error", onError);
  });
}

async function smtpCommand(socket, command, expectedCode) {
  socket.write(`${command}\r\n`);
  const response = await smtpRead(socket);
  if (!response.startsWith(String(expectedCode))) {
    throw new Error(`SMTP 指令失敗：${command}`);
  }
  return response;
}

function connectSmtp(host, portNumber, secure) {
  return new Promise((resolve, reject) => {
    const socket = secure ? tls.connect(portNumber, host) : net.connect(portNumber, host);
    socket.once("connect", () => resolve(socket));
    socket.once("secureConnect", () => resolve(socket));
    socket.once("error", reject);
  });
}

async function sendMail({ subject, text }) {
  const host = process.env.SMTP_HOST;
  const smtpPort = Number(process.env.SMTP_PORT || 587);
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS;
  const from = process.env.MAIL_FROM || user;

  if (!host || !user || !pass || !from) {
    return false;
  }

  let socket = await connectSmtp(host, smtpPort, smtpPort === 465);
  await smtpRead(socket);
  await smtpCommand(socket, `EHLO ${process.env.SMTP_DOMAIN || "localhost"}`, 250);

  if (smtpPort !== 465) {
    await smtpCommand(socket, "STARTTLS", 220);
    socket = tls.connect({ socket, servername: host });
    await smtpCommand(socket, `EHLO ${process.env.SMTP_DOMAIN || "localhost"}`, 250);
  }

  await smtpCommand(socket, "AUTH LOGIN", 334);
  await smtpCommand(socket, Buffer.from(user).toString("base64"), 334);
  await smtpCommand(socket, Buffer.from(pass).toString("base64"), 235);
  await smtpCommand(socket, `MAIL FROM:<${from}>`, 250);
  await smtpCommand(socket, `RCPT TO:<${recipient}>`, 250);
  await smtpCommand(socket, "DATA", 354);

  const encodedSubject = `=?UTF-8?B?${Buffer.from(subject).toString("base64")}?=`;
  const message = [
    `From: ${from}`,
    `To: ${recipient}`,
    `Subject: ${encodedSubject}`,
    "MIME-Version: 1.0",
    "Content-Type: text/plain; charset=UTF-8",
    "Content-Transfer-Encoding: 8bit",
    "",
    text.replace(/\n/g, "\r\n"),
    "."
  ].join("\r\n");

  await smtpCommand(socket, message, 250);
  socket.write("QUIT\r\n");
  socket.end();
  return true;
}

async function handleRegistration(request, response) {
  try {
    const rawBody = await readRequestBody(request);
    const data = JSON.parse(rawBody || "{}");
    const cleaned = Object.fromEntries(
      Object.entries(data).map(([key, value]) => [sanitizeText(key), sanitizeText(value)])
    );
    const missing = validateRegistration(cleaned);
    if (missing.length) {
      sendJson(response, 400, { message: `請補齊欄位：${missing.join("、")}` });
      return;
    }

    fs.mkdirSync(dataDir, { recursive: true });
    const record = {
      id: crypto.randomUUID(),
      submittedAt: new Date().toISOString(),
      data: cleaned
    };
    fs.appendFileSync(submissionsFile, `${JSON.stringify(record)}\n`, "utf8");
    appendRegistrationCsv(record);
    writeRegistrationXls(readSavedRegistrations());

    let emailSent = false;
    try {
      emailSent = await sendMail({
        subject: `課程報名｜${cleaned["姓名"]}｜${cleaned["報名梯次"]}`,
        text: formatRegistrationEmail(record)
      });
    } catch (error) {
      console.error("Email delivery failed:", error.message);
    }

    sendJson(response, 200, {
      message: "報名資料已送出",
      emailSent
    });
  } catch (error) {
    sendJson(response, 500, { message: "伺服器暫時無法處理報名資料" });
  }
}

function serveStatic(request, response) {
  const requestUrl = new URL(request.url, `http://${request.headers.host}`);
  const decodedPath = decodeURIComponent(requestUrl.pathname);
  const requestedFile = decodedPath === "/" ? "index.html" : decodedPath.replace(/^\/+/, "");
  const safePath = path.normalize(requestedFile).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(root, safePath);
  const publicRoots = [
    path.join(root, "index.html"),
    path.join(root, "課程詳情.png"),
    path.join(root, "講師介紹.png")
  ];

  if (!publicRoots.some((publicPath) => filePath === publicPath || filePath.startsWith(`${publicPath}${path.sep}`))) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (error, content) => {
    if (error) {
      response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("Not found");
      return;
    }
    response.writeHead(200, {
      "Content-Type": mimeTypes[path.extname(filePath).toLowerCase()] || "application/octet-stream"
    });
    response.end(content);
  });
}

const server = http.createServer((request, response) => {
  if (request.method === "POST" && request.url === "/api/register") {
    handleRegistration(request, response);
    return;
  }
  if (request.method === "GET" || request.method === "HEAD") {
    serveStatic(request, response);
    return;
  }
  response.writeHead(405, { "Content-Type": "text/plain; charset=utf-8" });
  response.end("Method not allowed");
});

fs.mkdirSync(dataDir, { recursive: true });
writeRegistrationXls(readSavedRegistrations());

server.listen(port, () => {
  console.log(`Course page server is running at http://localhost:${port}`);
});
