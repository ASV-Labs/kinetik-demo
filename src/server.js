import http from "node:http";
import { summarizeReadiness } from "./readiness.js";
import { listReleases } from "./releases.js";

const port = Number(process.env.PORT ?? 3000);

export function createServer() {
  return http.createServer((request, response) => {
    if (request.url === "/health") {
      return sendJson(response, 200, { ok: true });
    }

    if (request.url === "/api/releases") {
      return sendJson(response, 200, { releases: listReleases() });
    }

    if (request.url === "/api/readiness") {
      return sendJson(response, 200, { readiness: summarizeReadiness(listReleases()) });
    }

    return sendJson(response, 404, { error: "not_found" });
  });
}

function sendJson(response, statusCode, body) {
  response.writeHead(statusCode, { "content-type": "application/json" });
  response.end(JSON.stringify(body));
}

if (process.argv[1] && process.argv[1].endsWith("server.js")) {
  createServer().listen(port, () => {
    console.log(`Release Radar listening on ${port}`);
  });
}
