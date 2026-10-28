const http = require("http");
const fs = require("fs");
const path = require("path");

const root = __dirname;
const buildRoot = path.join(root, ".next");
const publicRoot = path.join(root, "public");
const indexFile = path.join(buildRoot, "server", "app", "index.html");

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".ico": "image/x-icon",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
};

function resolveFile(urlPath) {
  const cleanPath = decodeURIComponent(urlPath.split("?")[0]);
  if (cleanPath === "/" || cleanPath === "/index.html") return indexFile;
  if (cleanPath.startsWith("/_next/")) {
    return path.join(buildRoot, cleanPath.slice("/_next/".length));
  }
  return path.join(publicRoot, cleanPath.replace(/^\/+/, ""));
}

http
  .createServer((request, response) => {
    const file = resolveFile(request.url || "/");
    const allowedRoots = [buildRoot, publicRoot];
    const isAllowed = file === indexFile || allowedRoots.some((base) => file.startsWith(base + path.sep));

    if (request.method === "OPTIONS") {
      response.writeHead(204, {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, HEAD, OPTIONS",
        "Access-Control-Allow-Headers": "*",
      });
      response.end();
      return;
    }

    if (!isAllowed || !fs.existsSync(file) || !fs.statSync(file).isFile()) {
      response.writeHead(404, {
        "Content-Type": "text/plain; charset=utf-8",
        "Access-Control-Allow-Origin": "*",
      });
      response.end("Not found");
      return;
    }

    response.writeHead(200, {
      "Content-Type": mimeTypes[path.extname(file).toLowerCase()] || "application/octet-stream",
      "Cache-Control": file === indexFile ? "no-cache" : "public, max-age=3600",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, HEAD, OPTIONS",
    });
    fs.createReadStream(file).pipe(response);
  })
  .listen(3000, "127.0.0.1", () => {
    console.log("Frontend ready at http://localhost:3000/");
  });
