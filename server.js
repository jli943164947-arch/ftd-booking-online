const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = Number(process.env.PORT || 3000);
const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, "data");
const DATA_FILE = path.join(DATA_DIR, "reservations.json");

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

function send(res, status, body, contentType = "text/plain; charset=utf-8") {
  res.writeHead(status, {
    "Content-Type": contentType,
    "Cache-Control": "no-store"
  });
  res.end(body);
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
  if (req.method === "GET" && req.url === "/api/reservations") {
    send(res, 200, JSON.stringify(readReservations()), "application/json; charset=utf-8");
    return true;
  }

  if (req.method === "POST" && req.url === "/api/reservations") {
    try {
      const parsed = JSON.parse(await readBody(req));
      if (!Array.isArray(parsed) || !parsed.every(isValidReservation)) {
        send(res, 400, JSON.stringify({ error: "invalid reservations" }), "application/json; charset=utf-8");
        return true;
      }
      writeReservations(parsed);
      send(res, 200, JSON.stringify(readReservations()), "application/json; charset=utf-8");
    } catch (error) {
      send(res, 400, JSON.stringify({ error: "bad request" }), "application/json; charset=utf-8");
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
server.listen(PORT, () => {
  console.log(`FTD booking site running: http://localhost:${PORT}`);
});
