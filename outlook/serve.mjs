// Serves this folder over HTTPS on localhost:3000, which Outlook requires for add-ins.
// Certificates come from: npx office-addin-dev-certs install
import { createServer } from "node:https";
import { readFileSync, existsSync } from "node:fs";
import { homedir } from "node:os";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL(".", import.meta.url));
const certDir = join(homedir(), ".office-addin-dev-certs");
const types = { ".html": "text/html", ".js": "text/javascript", ".png": "image/png", ".xml": "text/xml" };

createServer(
  { key: readFileSync(join(certDir, "localhost.key")), cert: readFileSync(join(certDir, "localhost.crt")) },
  (req, res) => {
    const name = normalize(new URL(req.url, "https://localhost").pathname).replace(/^(\.\.[/\\])+/, "");
    const file = join(root, name === "/" ? "taskpane.html" : name);
    if (!file.startsWith(root) || !existsSync(file) || !types[extname(file)]) {
      res.writeHead(404).end("Not found");
      return;
    }
    res.writeHead(200, { "content-type": types[extname(file)] }).end(readFileSync(file));
  }
).listen(3000, () => console.log("Serving https://localhost:3000"));
