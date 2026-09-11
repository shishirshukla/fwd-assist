# Deploy Forward Guard on Railway (no ngrok)

Railway gives you a stable **HTTPS** URL. Outlook can load the add-in from that URL permanently.

## Overview

```
GitHub (push to main) → Railway deploy → Next.js serves /manifest.xml dynamically
                                                              ↓
                                              https://your-app.up.railway.app
                                                              ↓
                                              Outlook Web sideload manifest.xml
```

The manifest is **not** baked in at build time. `/manifest.xml` is generated on each request from `public/manifest.template.xml` using your Railway environment variables. That means redeploys pick up URL changes as soon as `PUBLIC_BASE_URL` is set.

## 1. Push the repo to GitHub

```bash
git remote add github git@github.com:YOUR_USER/fwd-assist.git
git push -u github main
```

## 2. Create a Railway project from GitHub

1. Open [Railway](https://railway.com) → **New Project**.
2. **Deploy from GitHub repo** → select **`fwd-assist`** / branch **`main`**.

## 3. Enable public HTTPS networking

1. Service → **Settings** → **Networking** → **Generate Domain**.
2. Copy the URL, e.g. `https://fwd-assist-production.up.railway.app`.

## 4. Set environment variables

In Railway → **Variables**:

| Variable | Value | Required |
| --- | --- | --- |
| `PUBLIC_BASE_URL` | `https://fwd-assist-production.up.railway.app` | **Yes** — use your exact HTTPS URL |
| `LETTER_SUBMIT_URL` | Override letter API base URL | No — defaults to the pensioner `submit-letter` API |

Railway also sets `RAILWAY_PUBLIC_DOMAIN` at **runtime** as a fallback, but `PUBLIC_BASE_URL` is more reliable (especially with custom domains).

**After changing the domain:** update `PUBLIC_BASE_URL` and click **Redeploy**.

## 5. Deploy

```bash
npm ci --include=dev
npm run build
npm start
```

On startup you should see in the deploy logs:

```
Manifest will be served at https://your-app.up.railway.app/manifest.xml
```

If `PUBLIC_BASE_URL` is missing on Railway, the app **fails to start** with a clear error.

## 6. Verify the deployment

| URL | Expected |
| --- | --- |
| `https://YOUR-APP.up.railway.app/api/deploy-info` | JSON with `configuredBaseUrl` = your Railway URL |
| `https://YOUR-APP.up.railway.app/manifest.xml` | XML where every URL uses your Railway domain (not `localhost`) |
| `https://YOUR-APP.up.railway.app/taskpane.html` | Classification form |
| `https://YOUR-APP.up.railway.app/api/captures` | JSON list of stored forward captures |
| `https://YOUR-APP.up.railway.app/captures` | Human-readable capture log |

### Quick check

```bash
curl -s https://YOUR-APP.up.railway.app/api/deploy-info
curl -s https://YOUR-APP.up.railway.app/manifest.xml | grep -o 'https://[^"]*' | head
```

If `manifest.xml` still shows `localhost`, `PUBLIC_BASE_URL` is not set or the redeploy did not finish.

## 7. Sideload in Outlook on the web

1. Download `https://YOUR-APP.up.railway.app/manifest.xml`.
2. Outlook Web → **Apps** → **My add-ins** → **Add from file** → upload it.
3. Forward a message → **Send** → fill the form → **Send** again.

## Troubleshooting

| Problem | Fix |
| --- | --- |
| `manifest.xml` still shows `localhost` | Set `PUBLIC_BASE_URL` exactly (with `https://`), redeploy, hard-refresh the URL |
| App won't start on Railway | Deploy logs: missing `PUBLIC_BASE_URL` — add the variable |
| Add-in install fails | Confirm icon/taskpane URLs load in a browser |
| Domain changed | Update `PUBLIC_BASE_URL`, redeploy, remove old add-in, sideload new manifest |

## Local development vs Railway

| | Local | Railway |
| --- | --- | --- |
| URL | `http://localhost:43123` | `https://*.up.railway.app` |
| Manifest | Defaults to `localhost` | Uses `PUBLIC_BASE_URL` at runtime |
| Outlook sideload | Needs ngrok | Works directly |

## Troubleshooting

| Problem | Fix |
| --- | --- |
| Build failed / exit 240 | Railway skipped Tailwind (`devDependencies`) because `NODE_ENV=production`. Current `main` uses `npm ci --include=dev`. Redeploy. |
| `Cannot find module '@tailwindcss/postcss'` | Same as above — pull latest and redeploy |
| Send blocked but sidebar stays closed | Outlook Web never auto-opens the task pane on Send. Click **Open form** (or **Take Action**) on the Send dialog. **Don't Send** only dismisses the alert. |

