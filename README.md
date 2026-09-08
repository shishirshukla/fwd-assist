# Forward Guard — Outlook Web add-in

Office add-in for Outlook on the web. It runs when the user clicks **Send**, checks whether the message is a **forward**, and if it is, **stops send** until Priority, End Date, and Category are filled in.

## How it works

1. Outlook raises `OnMessageSend` (Smart Alerts) when Send is clicked.
2. `public/launchevent.js` treats the item as forwarded when:
   - `getComposeTypeAsync` returns `Forward`, or
   - the subject starts with `FW:` / `Fwd:`
3. If it is not a forward, send continues.
4. If it is a forward and classification is missing, send is cancelled and Outlook can open the **Classify forward** task pane (`/taskpane`).
5. Saving the form writes custom properties (and optional internet headers) on the item. The next Send is allowed.

The home page embeds an **Outlook-style simulator** (`/simulator.html`) so you can try the same flow in a browser without sideloading.

## Run locally

```bash
npm install
npm run dev
```

Or a production build:

```bash
npm run build
npm start
```

Open [http://localhost:43123](http://localhost:43123).

- Simulator: `/` or `/simulator.html`
- Outlook task pane: `/taskpane.html`
- Manifest (dynamic): `/manifest.xml` (generated from `public/manifest.template.xml`)

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
6. Accept the prompt. You should see **Forward Guard** on the compose ribbon and a **Classify forward** button.

### 5. Exercise the Send intercept

| Step | Action | Expected result |
| --- | --- | --- |
| 1 | Forward a message, click **Send** | Send is blocked; Outlook may show a notification and open the task pane |
| 2 | Fill **Priority**, **End Date**, **Category** → **Save classification** | Success message in the pane |
| 3 | Click **Send** again | Message sends |
| 4 | Send a **new** (non-forward) message | Sends immediately — no form |

If Send is not blocked on a forward, open **Classify forward** manually from the ribbon, save the form, then try Send again.

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
4. Compose a **Forward**, click **Send**, complete the three fields, then send again.

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
| `public/simulator.html` | Browser simulator of the Send flow |
| `scripts/set-manifest-url.mjs` | Rewrite manifest URLs for your HTTPS tunnel |
| `scripts/inject-manifest-url.mjs` | Inject Railway/production URLs at build time |
| `docs/RAILWAY.md` | Deploy from GitHub to Railway (no ngrok) |
