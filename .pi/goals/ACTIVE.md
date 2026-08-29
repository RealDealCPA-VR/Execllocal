# Active Goal

Status: COMPLETE
Goal: Support models over a tailnet via HTTP (user types the endpoint in the app, bridged safely); push all work to GitHub.
Started: 2026-08-28 20:20
Last updated: 2026-08-28 20:50

## Scope and acceptance criteria
- [x] A1: /bridge dynamic forwarding in dev server + standalone proxy: pane sends x-llm-target header; only private/tailnet hosts allowed (SSRF guard); https targets tolerate private certs
- [x] A2: App: custom http:// URLs are bridged automatically (no more refusal), settings hint shown; error messages show the real target
- [x] A3: Verified live: /bridge with allowed target returns models JSON; disallowed public target rejected 403; full regression green
- [x] A4: Docs: README + TUTORIAL tailnet/LAN section
- [x] A5: Pushed to GitHub (incl. previous uncommitted plug-and-play + sideload-evidence work); screenshot excluded from git

## Constraints and decisions
- Excel sideload already user-confirmed visually ("i see it loaded in"); keep Excel + dev server RUNNING (user is using it) - tests use dedicated ports (8123/4010) so no conflict.
- Bridge guardrail: allow localhost/127/10/172.16-31/192.168/100.64-127 (tailscale CGNAT)/.ts.net/.local only. Public hosts rejected 403.
- Bridge https targets use rejectUnauthorized:false (personal private certs).

## Todo
- [x] P301: llm-forward.js dynamic target mode | Verify: node --check
- [x] P302: webpack + standalone proxy /bridge routes | Verify: build
- [x] P303: transport.ts extraHeaders + App.tsx bridge logic | Verify: npm run build
- [x] P304: live guard tests (allowed + 403) + regression | Verify: curl outputs + npm run check
- [x] P305: docs | Verify: read
- [x] P306: push | Verify: git push

## Progress and evidence
- 21:30 TAILNET FIX + HEADLESS MODE:
- Root cause of user tailnet failure: guard v1 rejected Tailscale MagicDNS SHORT names (http://gpu:8000 -> 403) and the pane discarded the 403 body (banner just said HTTP 403).
- Fixes: guard v2 (any hostname allowed - MagicDNS short names, .ts.net, LAN names; public IP literals still refused; IPv6 ULA fc/fd allowed; 16/16 unit cases pass), pane error banner now includes response body (bridge refusals self-explain).
- Headless mode: server/serve.js (serves dist + all bridges, zero deps, path-traversal guarded), tools/make-remote-manifest.js + manifest:remote (validated: manifest-remote.xml valid), sideload:remote via office-addin-dev-settings. README headless section + TUTORIAL Part 10 (mental model renumbered to 11).
- Verified: build PASS; serve.js smoke (static pane, /vllm, /bridge dynamic, public-IP 403); manifest-remote validate PASS; npm run check + test:integration green.
- USER ACTION NEEDED: restart npm start (running server predates guard v2); then Custom URL http://<short-magicdns-name>:8000 works.