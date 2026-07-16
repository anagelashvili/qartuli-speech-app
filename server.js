const http = require("http");
const fs = require("fs");
const path = require("path");
const ttsHandler = require("./api/tts");

const root = process.cwd();
const port = Number(process.argv[2]) || 7422;
const types = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
};

function sendJson(response, status, data) {
  response.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  response.end(JSON.stringify(data));
}

function parseBody(request) {
  return new Promise((resolve, reject) => {
    let raw = "";
    request.on("data", (chunk) => {
      raw += chunk;
    });
    request.on("end", () => {
      try {
        resolve(raw ? JSON.parse(raw) : {});
      } catch (error) {
        reject(error);
      }
    });
  });
}

async function handleApi(request, response) {
  try {
    request.body = await parseBody(request);
    response.status = (statusCode) => ({
      json: (data) => sendJson(response, statusCode, data),
      send: (data) => {
        response.statusCode = statusCode;
        response.end(data);
      },
    });
    response.setHeader = response.setHeader.bind(response);
    await ttsHandler(request, response);
  } catch (error) {
    sendJson(response, 500, { error: error.message || "Server error" });
  }
}

const server = http.createServer((request, response) => {
  const url = new URL(request.url, `http://127.0.0.1:${port}`);

  if (url.pathname === "/api/tts") {
    handleApi(request, response);
    return;
  }

  const name = url.pathname === "/" ? "index.html" : decodeURIComponent(url.pathname.slice(1));
  const filePath = path.normalize(path.join(root, name));

  if (!filePath.startsWith(root)) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (error, data) => {
    if (error) {
      response.writeHead(404);
      response.end("Not found");
      return;
    }

    response.writeHead(200, {
      "Content-Type": types[path.extname(filePath)] || "application/octet-stream",
    });
    response.end(data);
  });
});

server.listen(port, "127.0.0.1", () => {
  console.log(`Georgian presentation reader: http://127.0.0.1:${port}/`);
});
