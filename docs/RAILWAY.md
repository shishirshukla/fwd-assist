# Deploy Forward Guard on Railway (no ngrok)

Railway gives you a stable **HTTPS** URL. Outlook can load the add-in from that URL permanently.

## Overview

```
GitHub (push to main) → Railway auto-build → inject manifest URLs → Next.js server
                                                              ↓
                                              https://your-app.up.railway.app
                                                              ↓
                                              Outlook Web sideload manifest.xml
```

## 1. Push the repo to GitHub

If the code is only on Cursor Origin today, mirror it to GitHub first:

```bash
# In WSL, from your local clone
git remote add github git@github.com:YOUR_USER/fwd-assist.git
git push -u github main
```

Or create a new repo on GitHub and push:

```bash
gh repo create fwd-assist --private --source=. --remote=github --push
```

## 2. Create a Railway project from GitHub

1. Open [Railway](https://railway.com) → **New Project**.
2. Choose **Deploy from GitHub repo**.
3. Authorize Railway to access your GitHub account if prompted.
4. Select **`fwd-assist`** (or your repo name) and the **`main`** branch.
5. Railway detects Next.js via Nixpacks and reads [`railway.toml`](../railway.toml).

No separate GitHub Actions deploy step is required — Railway watches `main` and redeploys on every push.

## 3. Enable public HTTPS networking

1. In Railway, open your service → **Settings** → **Networking**.
2. Click **Generate Domain** (or attach a custom domain).
3. Copy the public URL, e.g. `https://fwd-assist-production.up.railway.app`.

## 4. Set environment variables

In Railway → your service → **Variables**, add:

| Variable | Value | Required |
| --- | --- | --- |
| `PUBLIC_BASE_URL` | `https://fwd-assist-production.up.railway.app` | Recommended — use your exact Railway HTTPS URL |
| `NODE_ENV` | `production` | Set automatically by Railway |

`PUBLIC_BASE_URL` is used at **build time** to rewrite every URL in `public/manifest.xml`.  
Railway also exposes `RAILWAY_PUBLIC_DOMAIN`; the build uses that as a fallback if `PUBLIC_BASE_URL` is not set.

**Important:** After you generate or change your Railway domain, set `PUBLIC_BASE_URL` and trigger a **Redeploy** so the manifest is rebuilt with the correct host.

## 5. Deploy

Railway runs on each push to `main`:

```bash
npm ci
npm run build    # injects manifest URLs, then next build
npm start        # listens on Railway's $PORT
```

Watch the deploy logs for:

```
Injected manifest URLs for https://your-app.up.railway.app (AppDomain: your-app.up.railway.app)
```

## 6. Verify the deployment

Open these URLs in a browser (replace with your Railway domain):

| URL | Expected |
| --- | --- |
| `https://YOUR-APP.up.railway.app/simulator.html` | Outlook simulator |
| `https://YOUR-APP.up.railway.app/taskpane.html` | Classification form |
| `https://YOUR-APP.up.railway.app/icons/icon-64.png` | Icon image |
| `https://YOUR-APP.up.railway.app/manifest.xml` | Manifest XML (URLs should match your Railway domain) |

## 7. Sideload in Outlook on the web

1. Download the live manifest:  
   `https://YOUR-APP.up.railway.app/manifest.xml`  
   (Save as `manifest.xml` on your machine.)
2. Outlook Web → **Apps** → **My add-ins** → **Add a custom add-in** → **Add from file**.
3. Upload the downloaded `manifest.xml`.
4. Forward a message → **Send** → complete the form → **Send** again.

You do **not** need ngrok when using Railway.

## 8. Updates

Every `git push` to `main`:

1. GitHub notifies Railway.
2. Railway rebuilds with the current `PUBLIC_BASE_URL`.
3. Outlook keeps working — no manifest change needed unless the domain changes.

If you change the Railway domain or add a custom domain:

1. Update `PUBLIC_BASE_URL` in Railway variables.
2. Redeploy.
3. Remove the old add-in in Outlook and sideload the new `manifest.xml`.

## GitHub Actions (optional)

[`.github/workflows/ci.yml`](../.github/workflows/ci.yml) runs lint, build, and manifest validation on every push/PR. It does **not** deploy — Railway handles deployment via its GitHub integration.

## Troubleshooting

| Problem | Fix |
| --- | --- |
| Manifest still shows `localhost` | Set `PUBLIC_BASE_URL` and redeploy; check build logs for "Injected manifest URLs" |
| Add-in install fails | Confirm icon/taskpane URLs load in a browser tab |
| Build fails on Railway | Check deploy logs; run `npm run build` locally |
| Domain changed | Update `PUBLIC_BASE_URL`, redeploy, re-sideload manifest |

## Local development vs Railway

| | Local | Railway |
| --- | --- | --- |
| URL | `http://localhost:43123` | `https://*.up.railway.app` |
| Outlook sideload | Needs ngrok | Works directly |
| Manifest URLs | `localhost` in git | Injected at build from `PUBLIC_BASE_URL` |
