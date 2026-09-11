# Forward Guard — Outlook Web add-in

Office add-in for Outlook on the web. It runs when the user clicks **Send**, detects a **forward**, extracts the original letter details, posts them to the capture API (which calls submit-letter), and **lets send continue**. There is no classification form.

## How it works

1. Outlook raises `OnMessageSend` when Send is clicked.
2. `public/launchevent.js` treats the item as forwarded when:
   - `getComposeTypeAsync` returns `Forward`, or
   - the subject starts with `FW:` / `Fwd:`
3. If it is not a forward, send continues with no capture.
4. If it is a forward, the add-in reads sender, date, subject, body, and To/Cc, `POST`s `/api/captures`, then allows send.

The home page embeds an **Outlook-style simulator** (`/simulator.html`) so you can try the same flow in a browser without sideloading.

## Run locally

## Run locally (WSL)

Put the repo on the **Linux filesystem**, not under `/mnt/c`. Next.js often hangs with no extra output when the project is on the Windows drive.

```bash
# Copy off /mnt/c (adjust the source if your clone is elsewhere)
mkdir -p ~/src
cp -a /mnt/c/Users/HP/fwd-assist ~/src/fwd-assist
cd ~/src/fwd-assist

git pull
npm install
npm start
```

Wait until you see **Ready** or **Local:** then open http://127.0.0.1:43123

The first compile can take a minute. If it still sits forever on `next dev --hostname ...`, you are almost certainly on `/mnt/c` — use `~/src/fwd-assist` as above.

If port 43123 is already taken:

```bash
ss -ltnp | grep 43123
# or
npx kill-port 43123
npm start
```

Open http://localhost:43123

`npm start` runs the **dev** server if you have not built yet. After `npm run build`, it runs **production** (`next start`).

You do not need to type `next start --hostname 0.0.0.0 --port 43123` yourself. That line is Next.js starting. If it then says it cannot find a production build, use:

```powershell
npm install
npm start
```

or:

```powershell
npm run dev
```

On Windows, use **PowerShell** or **Command Prompt** in that folder. If `npm` is not found, install [Node.js LTS](https://nodejs.org/) and reopen the terminal.

| Error | Cause | Fix |
| --- | --- | --- |
| `next` / `Cannot find module` | Dependencies missing | `npm install` |
| `Could not find a production build` | `next start` without `.next` | Pull latest, then `npm install` and `npm start` (falls back to `npm run dev`) |
| `EADDRINUSE` / port in use | Another process on 43123 | Stop the other server, or `npm run dev` |
| `${PORT:-43123}` as a port | Old start script on Windows | Pull latest `main` (start is now `node scripts/start.mjs`) |

- Simulator: `/` or `/simulator.html`
- Stored captures: `/captures` and `GET /api/captures`
- Logs: `/logs`, `GET /api/logs`, `GET /api/logs?format=text`
- Outlook task pane: `/taskpane.html`
- Manifest (dynamic): `/manifest.xml` (generated from `public/manifest.template.xml`)

## Capture API

When a forwarded message is captured, the add-in `POST`s **`/api/captures`**. That route stores a text record and Node `POST`s JSON to:

`https://eloan.cgbankmobile.in/pensioner_api/auth/api/submit-letter`

as **POST** `Content-Type: application/json`:

```json
{
  "receivingDate": "2026-09-10",
  "senderOffice": "NA",
  "sendName": "NA",
  "letterNo": "NA",
  "letterDate": "2026-09-10",
  "letterDesc": "NA",
  "department": "TEST",
  "priority": "High",
  "EntryBy": "BOD"
}
```

| JSON field | Source |
| --- | --- |
| `receivingDate` | Original email date as `yyyy-mm-dd` |
| `letterDate` | Same as `receivingDate` |
| `senderOffice` | Original sender email, or `NA` |
| `sendName` | Original sender name, or `NA` |
| `letterNo` | Always `NA` |
| `letterDesc` | Subject line |
| `priority` | `NA` (no form). Override later via lookup if needed |
| `department` | Lookup of the forward **To** address in `data/letter-lookup.json` |
| `EntryBy` | Lookup of the forward **From** address (mailbox user) in `data/letter-lookup.json` |

Edit **`data/letter-lookup.json`** to add live bank addresses:

```json
{
  "departmentByToEmail": { "it@yourbank.in": "IT" },
  "entryByFromEmail": { "officer@yourbank.in": "BOD" },
  "defaults": { "department": "NA", "entryBy": "NA" }
}
```

| Variable | Purpose |
| --- | --- |
| `LETTER_SUBMIT_URL` or `CAPTURE_PUSH_URL` | Override the letter API base URL |
| `LETTER_SUBMIT_DISABLED` | Set to `1` to skip the remote call |
| `LETTER_SUBMIT_TLS_INSECURE` | Set to `1` if the bank HTTPS certificate is untrusted |
| `LETTER_SUBMIT_TIMEOUT_MS` | Outbound timeout, default `30000` |
| `CAPTURE_FILE_PATH` | Optional override for the local text file path |
| `APP_LOG_PATH` | Optional override for `data/app-logs.txt` |

```bash
curl -s http://127.0.0.1:43123/api/captures
curl -s http://127.0.0.1:43123/api/letter-health
curl -sS -m 45 http://127.0.0.1:43123/api/letter-dummy
```

`GET` or `POST /api/letter-dummy` sends a dummy JSON body to submit-letter from **this Node process** (same path as a real capture). The response includes `tcpLogs` (DNS, TCP connect, TLS handshake, peer certificate, HTTP socket events) and `certificateHint`. Events are also written to `/logs` with source `letter-dummy-tcp`.

```bash
curl -sS -m 45 https://YOUR-APP.up.railway.app/api/letter-dummy
curl -sS -m 45 -X POST https://YOUR-APP.up.railway.app/api/letter-dummy
```

## Logs API

Capture handling and letter-API calls are appended to **`data/app-logs.txt`**.

| Endpoint | Response |
| --- | --- |
| `GET /api/logs` | JSON `{ count, logs }` (newest first) |
| `GET /api/logs?format=text` | Raw text file |
| `GET /api/logs?level=error` | Errors only |
| `GET /api/logs?source=letter-submit` | Letter API logs only |
| `/logs` | Browser view of the same data |

```bash
curl -s http://127.0.0.1:43123/api/logs
curl -s http://127.0.0.1:43123/api/logs?format=text
```

## Deploy on Railway (recommended — no ngrok)

Use Railway for a permanent **HTTPS** URL so Outlook can load the add-in without a tunnel.

**Full guide:** [`docs/RAILWAY.md`](docs/RAILWAY.md)

Quick steps:

1. Push this repo to **GitHub**.
2. In [Railway](https://railway.com): **New Project** → **Deploy from GitHub repo** → select `fwd-assist`.
3. **Settings → Networking** → **Generate Domain**.
4. **Variables** → set `PUBLIC_BASE_URL` = `https://YOUR-APP.up.railway.app`.
5. **Redeploy** — open `https://YOUR-APP.up.railway.app/api/deploy-info` and confirm `configuredBaseUrl` matches your Railway domain.
6. Download `https://YOUR-APP.up.railway.app/manifest.xml` and sideload in Outlook Web.

### Sideload after this change

After a version bump (now **1.0.10.0**), **remove** Forward Guard and sideload `manifest.xml` again. Forward a message and click **Send**. The mail should go out; check `/captures` and `/logs` for the letter payload and the Node submit-letter result.

Every push to `main` triggers a new Railway deploy automatically.

## Test in Outlook on the web (local + ngrok)

Outlook loads add-ins only over **HTTPS**. Your machine serves the app on port **43123**; expose it with a tunnel (ngrok is the simplest), point the manifest at that URL, then sideload.

### 1. Clone and run (WSL)

```bash
cd ~/fwd-assist   # or wherever you cloned the repo
npm install
npm run build
npm start
```

Leave that terminal running. Confirm locally: `curl -I http://127.0.0.1:43123/taskpane.html`

### 2. HTTPS tunnel

In a **second WSL terminal**, install and start ngrok (one-time signup at [ngrok.com](https://ngrok.com)):

```bash
# Example with ngrok — any HTTPS reverse proxy works
ngrok http 43123
```

Copy the **https** forwarding URL (e.g. `https://abc123.ngrok-free.app`). Do not use the `http://` URL.

### 3. Point the manifest at your tunnel

```bash
npm run manifest:url -- https://abc123.ngrok-free.app
```

This rewrites every `https://localhost:43123` entry in `public/manifest.xml`.

Quick sanity check in a browser:

- `https://YOUR-TUNNEL/taskpane.html` — classification form
- `https://YOUR-TUNNEL/manifest.xml` — manifest downloads

### 4. Sideload in Outlook on the web

1. Open [Outlook on the web](https://outlook.office.com) and sign in with a **Microsoft 365** work or school account (personal Outlook.com has limited add-in support).
2. Open any message → **Forward** (or start a new compose and forward an existing mail).
3. On the ribbon: **Apps** (or **Get Add-ins**) → **My add-ins**.
4. Under **Custom add-ins** → **Add a custom add-in** → **Add from file**.
5. Upload `public/manifest.xml` from your repo (the file you just updated with the tunnel URL).
6. Accept the prompt. Forward Guard is registered for Send on compose.

### 5. Exercise the Send intercept

| Step | Action | Expected result |
| --- | --- | --- |
| 1 | Forward a message, click **Send** | Message sends; capture appears at `/captures` and `/logs` |
| 2 | Send a **new** (non-forward) message | Sends with no capture |

### Requirements and troubleshooting

| Issue | What to check |
| --- | --- |
| **Add-in installation fails** | See [Installation failures](#installation-failures) below. |
| Add-in fails to install | Manifest URLs must be **https** and reachable from the public internet (tunnel running, `npm start` still up). |
| Blank task pane | Open `https://YOUR-TUNNEL/taskpane.html` in a normal browser tab first. |
| Send never intercepted | `OnMessageSend` needs Mailbox **1.12+** and a Microsoft 365 mailbox. Your tenant admin may need to allow **on-send** / Smart Alerts add-ins. |
| “We can’t load this add-in” | Tunnel URL changed — re-run `npm run manifest:url` and remove/re-add the add-in. |
| ngrok browser warning | Click through the ngrok interstitial once, or use a paid/static domain. |

### Installation failures

Outlook rejects manifests for a few common reasons. Run these checks **before** uploading:

```bash
# 1. Validate XML schema (must pass)
npm run manifest:validate

# 2. Point manifest at your live HTTPS tunnel and verify URLs respond
npm run manifest:url -- https://YOUR-SUBDOMAIN.ngrok-free.app
```

**Checklist if install still fails:**

1. **Do not upload a manifest with `localhost` URLs.** Outlook’s servers cannot fetch icons or pages from your machine. Always run `npm run manifest:url` with your ngrok **https** URL first.
2. **`AppDomain` must be the hostname only** (e.g. `abc123.ngrok-free.app`), not `https://…`. The helper script sets this automatically.
3. **`IconUrl` must be a PNG file**, not the site root. It should end with `/icons/icon-64.png`.
4. **Keep `npm start` and ngrok running** while installing. If any URL returns 404, install fails.
5. **Remove a broken sideload** before retrying: My add-ins → Custom add-ins → ⋯ → Remove, then upload the new manifest.
6. **Use a Microsoft 365 work/school account.** Personal Outlook.com accounts have limited custom add-in support.
7. **“Sideloading rejected by Exchange”** — your tenant may block custom add-ins; ask an admin to allow user-installed add-ins for testing.

If install fails, open browser DevTools (F12) → **Network** while uploading the manifest and look for a `400` response — the body often says `Sideloading rejected by Exchange`.

### Remove the test add-in

**My add-ins** → **Custom add-ins** → **⋯** next to Forward Guard → **Remove**.

## Sideload in Outlook on the web (short)

Outlook add-ins must be served over **HTTPS**. Point every `https://localhost:43123` URL in `public/manifest.xml` at your HTTPS host (tunnel or trusted local certs), then:

1. In Outlook on the web, open a message compose window.
2. Go to **Get add-ins** → **My add-ins** → **Add a custom add-in** → **Add from file**.
3. Upload `public/manifest.xml`.
4. Compose a **Forward** and click **Send**. The add-in captures the letter and allows send.

### Tenant notes

- Smart Alerts (`OnMessageSend`) needs Mailbox 1.12 and a Microsoft 365 mailbox.
- The older Outlook on the web **ItemSend** event in the same manifest may require an admin to allow on-send add-ins.

## Project layout

| Path | Role |
| --- | --- |
| `public/manifest.template.xml` | Manifest template (URLs injected at request time) |
| `public/launchevent.js` | Send intercept (no DOM; Office event runtime) |
| `public/commands.html` | Command / runtime HTML host |
| `public/taskpane.html` | Classification form hosted in Outlook |
| `public/capture.js` | Collects sender/date/body/TO/CC and POSTs `/api/captures` |
| `public/simulator.html` | Browser simulator of the Send flow |
| `lib/letter-submit.ts` | Maps capture fields and calls submit-letter |
| `lib/app-log.ts` | Appends application logs to `data/app-logs.txt` |
| `app/api/logs/route.ts` | `GET /api/logs` JSON and text log access |
| `data/letter-lookup.json` | TO → department and From → EntryBy dictionaries |
| `scripts/set-manifest-url.mjs` | Rewrite manifest URLs for your HTTPS tunnel |
| `scripts/inject-manifest-url.mjs` | Inject Railway/production URLs at build time |
| `docs/RAILWAY.md` | Deploy from GitHub to Railway (no ngrok) |
