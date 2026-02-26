const https = require("https");
const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = 3000;
const ROOT = __dirname;

const MIME = {
  ".html": "text/html",
  ".js":   "application/javascript",
  ".css":  "text/css",
  ".xml":  "application/xml",
  ".png":  "image/png",
  ".ico":  "image/x-icon",
  ".json": "application/json",
};

const sslOptions = {
  key:  fs.readFileSync(path.join(ROOT, "localhost+1-key.pem")),
  cert: fs.readFileSync(path.join(ROOT, "localhost+1.pem")),
};

const server = https.createServer(sslOptions, (req, res) => {
  // CORS headers
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "*");

  if (req.method === "OPTIONS") {
    res.writeHead(204);
    res.end();
    return;
  }

  // Proxy /proxy/* → http://127.0.0.1:8402/*
  if (req.url.startsWith("/proxy/")) {
    const targetPath = req.url.replace("/proxy", "");
    let body = [];

    req.on("data", chunk => body.push(chunk));
    req.on("end", () => {
      body = Buffer.concat(body);

      const options = {
        hostname: "127.0.0.1",
        port: 8402,
        path: targetPath,
        method: req.method,
        headers: {
          "content-type": req.headers["content-type"] || "application/json",
          "content-length": body.length
        }
      };

      const proxyReq = http.request(options, (proxyRes) => {
        res.writeHead(proxyRes.statusCode, {
          "content-type": proxyRes.headers["content-type"] || "application/json",
          "access-control-allow-origin": "*"
        });
        proxyRes.pipe(res);
      });

      proxyReq.on("error", (err) => {
        res.writeHead(502);
        res.end(JSON.stringify({ error: { message: "OpenClaw proxy error: " + err.message } }));
      });

      proxyReq.write(body);
      proxyReq.end();
    });

    return;
  }

  // Static file serving
  const urlPath = req.url.split("?")[0];
  const filePath = path.join(ROOT, urlPath === "/" ? "taskpane.html" : urlPath);
  const ext = path.extname(filePath);

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end("Not found: " + urlPath);
      return;
    }
    res.writeHead(200, { "Content-Type": MIME[ext] || "text/plain" });
    res.end(data);
  });
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`[${new Date().toISOString()}] Outlook Add-in server running at https://localhost:${PORT}`);
});

server.on("error", (err) => {
  console.error(`[${new Date().toISOString()}] Server error:`, err.message);
});
