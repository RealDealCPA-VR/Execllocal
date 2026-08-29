/*
 * Generate a manifest that points at a remotely hosted pane (headless mode).
 *   node tools/make-remote-manifest.js https://your-box.your-tailnet.ts.net
 * Writes manifest-remote.xml next to manifest.xml.
 */
const fs = require("fs");
const path = require("path");

const base = (process.argv[2] || "").trim().replace(/\/+$/, "");
if (!/^https:\/\/[a-z0-9.-]+(:\d+)?$/i.test(base)) {
  console.error("Usage: node tools/make-remote-manifest.js https://<host>[:port]");
  console.error("The base URL must be HTTPS (Office requirement) - e.g. the address tailscale serve gives you.");
  process.exit(1);
}

const manifest = fs.readFileSync(path.join(__dirname, "..", "manifest.xml"), "utf8");
const remote = manifest.split("https://localhost:3000").join(base);

// AppDomains: keep localhost entries, add the remote origin
const origin = base;
if (!remote.includes("<AppDomain>" + origin + "</AppDomain>")) {
  const out = remote.replace(
    /(  <AppDomains>)/,
    "$1\n    <AppDomain>" + origin + "</AppDomain>"
  );
  fs.writeFileSync(path.join(__dirname, "..", "manifest-remote.xml"), out);
} else {
  fs.writeFileSync(path.join(__dirname, "..", "manifest-remote.xml"), remote);
}
console.log("Wrote manifest-remote.xml pointing at " + base);
console.log("Next: sideload it ->  npm run sideload:remote");
