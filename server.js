const http = require("http");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const PORT = Number(process.env.PORT || 3000);
const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, "data");
const DATA_FILE = path.join(DATA_DIR, "reservations.json");
const DATABASE_URL = process.env.DATABASE_URL || "";
const { Pool } = DATABASE_URL ? require("pg") : { Pool: null };
const pool = DATABASE_URL
  ? new Pool({
      connectionString: DATABASE_URL,
      ssl: process.env.NODE_ENV === "production" ? { rejectUnauthorized: false } : undefined
    })
  : null;
let dbReady = false;

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".ico": "image/x-icon"
};

function ensureDataFile() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(DATA_FILE)) fs.writeFileSync(DATA_FILE, "[]", "utf8");
}

function readReservations() {
  ensureDataFile();
  try {
    const data = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
    return Array.isArray(data) ? data : [];
  } catch (error) {
    return [];
  }
}

function writeReservations(reservations) {
  ensureDataFile();
  const tmpFile = `${DATA_FILE}.tmp`;
  fs.writeFileSync(tmpFile, JSON.stringify(reservations, null, 2), "utf8");
  fs.renameSync(tmpFile, DATA_FILE);
}

async function initDatabase() {
  if (!pool) return;
  await pool.query(`
    CREATE TABLE IF NOT EXISTS reservations (
      id TEXT PRIMARY KEY,
      date TEXT NOT NULL,
      start_time TEXT NOT NULL,
      end_time TEXT NOT NULL,
      people JSONB NOT NULL,
      created_at TEXT NOT NULL
    )
  `);
  dbReady = true;
}

async function readStoredReservations() {
  if (!pool) return readReservations();
  const result = await pool.query(`
    SELECT id, date, start_time AS start, end_time AS end, people, created_at AS "createdAt"
    FROM reservations
    ORDER BY date ASC, start_time ASC, end_time ASC, created_at ASC
  `);
  return result.rows.map((row) => ({
    ...row,
    people: Array.isArray(row.people) ? row.people : []
  }));
}

async function writeStoredReservations(reservations) {
  if (!pool) {
    writeReservations(reservations);
    return;
  }

  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    await client.query("DELETE FROM reservations");
    for (const item of reservations) {
      await client.query(
        `INSERT INTO reservations (id, date, start_time, end_time, people, created_at)
         VALUES ($1, $2, $3, $4, $5::jsonb, $6)`,
        [
          item.id,
          item.date,
          item.start,
          item.end,
          JSON.stringify(item.people),
          item.createdAt || new Date().toISOString()
        ]
      );
    }
    await client.query("COMMIT");
  } catch (error) {
    await client.query("ROLLBACK");
    throw error;
  } finally {
    client.release();
  }
}

function send(res, status, body, contentType = "text/plain; charset=utf-8") {
  res.writeHead(status, {
    "Content-Type": contentType,
    "Cache-Control": "no-store"
  });
  res.end(body);
}

function sendFile(res, buffer, fileName) {
  res.writeHead(200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`,
    "Content-Length": buffer.length,
    "Cache-Control": "no-store"
  });
  res.end(buffer);
}

function parseDateKey(key) {
  const parts = key.split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function dateKey(date) {
  const pad = (value) => String(value).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function formatDateLabel(key) {
  const date = parseDateKey(key);
  return `${String(date.getMonth() + 1).padStart(2, "0")}月${String(date.getDate()).padStart(2, "0")}日`;
}

function formatWeekday(key) {
  return ["周日", "周一", "周二", "周三", "周四", "周五", "周六"][parseDateKey(key).getDay()];
}

function buildExportDates(reservations, mode, start, end) {
  if (mode === "history") {
    return [...new Set(reservations.map((item) => item.date))].sort();
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(start || "") || !/^\d{4}-\d{2}-\d{2}$/.test(end || "")) {
    return [];
  }

  const dates = [];
  const current = parseDateKey(start);
  const endDate = parseDateKey(end);
  while (current <= endDate) {
    const day = current.getDay();
    if (day >= 3 && day <= 5) dates.push(dateKey(current));
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

function reservationCellText(reservation) {
  const people = Array.isArray(reservation.people) ? reservation.people : [];
  const teachers = people.filter((person) => person.role === "教员").map((person) => person.name).join("、") || "无";
  const students = people.filter((person) => person.role === "学员").map((person) => person.name);
  return [
    `${reservation.start.replace(":", "")}-${reservation.end.replace(":", "")}`,
    `教员：${teachers}`,
    `学员：${students[0] || ""}`,
    `学员：${students.slice(1).join("、")}`
  ].join("\n");
}

async function buildExcelWorkbook(reservations, mode, start, end) {
  const dates = buildExportDates(reservations, mode, start, end);
  const grouped = dates.map((date) => ({
    date,
    reservations: reservations
      .filter((item) => item.date === date)
      .sort((a, b) => `${a.start}-${a.end}`.localeCompare(`${b.start}-${b.end}`))
  }));
  const maxBookings = Math.max(1, ...grouped.map((day) => day.reservations.length));
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("FTD使用记录", {
    pageSetup: {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0
    }
  });

  sheet.getColumn(1).width = 18;
  for (let index = 2; index <= maxBookings + 1; index += 1) {
    sheet.getColumn(index).width = 30;
  }

  const border = {
    top: { style: "thin", color: { argb: "FFD4DEDC" } },
    left: { style: "thin", color: { argb: "FFD4DEDC" } },
    bottom: { style: "thin", color: { argb: "FFD4DEDC" } },
    right: { style: "thin", color: { argb: "FFD4DEDC" } }
  };

  grouped.forEach((day, rowIndex) => {
    const row = sheet.getRow(rowIndex + 1);
    row.height = 112;

    const dateCell = row.getCell(1);
    dateCell.value = `${formatDateLabel(day.date)}\n${formatWeekday(day.date)}`;
    dateCell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FF172423" } };
    dateCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    dateCell.border = border;

    for (let index = 0; index < maxBookings; index += 1) {
      const cell = row.getCell(index + 2);
      const reservation = day.reservations[index];
      cell.value = reservation ? reservationCellText(reservation) : "";
      cell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FF172423" } };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.border = border;
      if (reservation) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFBF1E2" } };
      }
    }
  });

  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => {
      body += chunk;
      if (body.length > 2_000_000) {
        reject(new Error("body too large"));
        req.destroy();
      }
    });
    req.on("end", () => resolve(body));
    req.on("error", reject);
  });
}

function isValidReservation(item) {
  return item
    && typeof item.id === "string"
    && /^\d{4}-\d{2}-\d{2}$/.test(item.date)
    && /^\d{2}:\d{2}$/.test(item.start)
    && /^\d{2}:\d{2}$/.test(item.end)
    && Array.isArray(item.people)
    && item.people.length > 0
    && item.people.length <= 5;
}

async function handleApi(req, res) {
  const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);

  if (req.method === "GET" && url.pathname === "/api/storage-status") {
    send(res, 200, JSON.stringify({
      storage: dbReady ? "database" : "temporary-file",
      persistent: dbReady
    }), "application/json; charset=utf-8");
    return true;
  }

  if (req.method === "GET" && url.pathname === "/api/reservations") {
    try {
      send(res, 200, JSON.stringify(await readStoredReservations()), "application/json; charset=utf-8");
    } catch (error) {
      send(res, 500, JSON.stringify({ error: "storage unavailable" }), "application/json; charset=utf-8");
    }
    return true;
  }

  if (req.method === "POST" && url.pathname === "/api/reservations") {
    try {
      const parsed = JSON.parse(await readBody(req));
      if (!Array.isArray(parsed) || !parsed.every(isValidReservation)) {
        send(res, 400, JSON.stringify({ error: "invalid reservations" }), "application/json; charset=utf-8");
        return true;
      }
      await writeStoredReservations(parsed);
      send(res, 200, JSON.stringify(await readStoredReservations()), "application/json; charset=utf-8");
    } catch (error) {
      send(res, 400, JSON.stringify({ error: "bad request" }), "application/json; charset=utf-8");
    }
    return true;
  }

  if (req.method === "GET" && url.pathname === "/api/export-reservations") {
    try {
      const mode = url.searchParams.get("mode") === "history" ? "history" : "future";
      const start = url.searchParams.get("start") || "";
      const end = url.searchParams.get("end") || "";
      const reservations = await readStoredReservations();
      const buffer = await buildExcelWorkbook(reservations, mode, start, end);
      const range = mode === "history" ? "历史预约记录" : `${start}_${end}`;
      sendFile(res, buffer, `飞行部FTD使用记录_${range}.xlsx`);
    } catch (error) {
      send(res, 500, JSON.stringify({ error: "export failed" }), "application/json; charset=utf-8");
    }
    return true;
  }

  return false;
}

function serveStatic(req, res) {
  const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);
  const pathname = decodeURIComponent(url.pathname === "/" ? "/index.html" : url.pathname);
  const requestedPath = path.normalize(path.join(ROOT, pathname));

  if (!requestedPath.startsWith(ROOT)) {
    send(res, 403, "Forbidden");
    return;
  }

  fs.readFile(requestedPath, (error, data) => {
    if (error) {
      send(res, 404, "Not found");
      return;
    }
    const contentType = mimeTypes[path.extname(requestedPath)] || "application/octet-stream";
    send(res, 200, data, contentType);
  });
}

const server = http.createServer(async (req, res) => {
  if (await handleApi(req, res)) return;
  serveStatic(req, res);
});

ensureDataFile();
initDatabase()
  .then(() => {
    server.listen(PORT, () => {
      const storage = dbReady ? "PostgreSQL" : "local file";
      console.log(`FTD booking site running: http://localhost:${PORT} (${storage} storage)`);
    });
  })
  .catch((error) => {
    console.error("Database initialization failed:", error);
    process.exit(1);
  });
