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
- Manifest: [`public/manifest.xml`](public/manifest.xml)

## Sideload in Outlook on the web

Outlook add-ins must be served over **HTTPS**. Point every `https://localhost:43123` URL in `public/manifest.xml` at your HTTPS host (trusted local certs, or a tunnel such as ngrok), then:

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
| `public/manifest.xml` | Outlook add-in manifest |
| `public/launchevent.js` | Send intercept (no DOM; Office event runtime) |
| `public/commands.html` | Command / runtime HTML host |
| `public/taskpane.html` | Classification form hosted in Outlook |
| `public/simulator.html` | Browser simulator of the Send flow |
