import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import test from "node:test";

const source = readFileSync(new URL("../outlook/launchevent.js", import.meta.url), "utf8");
const base = "https://addin.example.com";
const api = "https://lms.example.com/submit-letter";

function runtime({ forward = true, subject = "FW: [LMS] Test", failApi = false, hangSubject = false, hangApi = false, xhrOnly = false } = {}) {
  const calls = [], completions = [], headers = [], timers = new Map();
  let nextTimer = 0;
  const result = (value) => ({ status: "succeeded", value });
  const item = {
    getComposeTypeAsync: (cb) => cb(result({ composeType: forward ? "forward" : "newMail" })),
    subject: { getAsync: (cb) => { if (!hangSubject) cb(result(subject)); } },
    body: { getAsync: (_, cb) => cb(result("From: Sender <sender@example.com>\nSent: 2026-10-05\nTo: original@example.com\nBody")) },
    to: { getAsync: (cb) => cb(result([{ emailAddress: "one@example.com" }, { emailAddress: "two@example.com" }])) },
    cc: { getAsync: (cb) => cb(result([])) },
    from: { getAsync: (cb) => cb(result({ emailAddress: "user@example.com" })) },
    internetHeaders: { setAsync: (value, cb) => { headers.push(value); cb(result()); } },
  };
  function response(url, options) {
    assert.match(url, /^https:\/\//);
    calls.push({ url, ...options, payload: JSON.parse(options.body) });
    if (url === api && failApi) throw new Error("CORS rejected");
    if (url.endsWith("/api/captures")) return {
      id: "capture-1", letterSubmitShouldRun: /\[LMS\]/i.test(subject), letterSubmitUrl: api,
      letterSubmits: [{ fields: { department: "one" } }, { fields: { department: "two" } }],
    };
    if (url === api) return { letterId: `letter-${calls.filter((c) => c.url === api).length}` };
    return { ok: true };
  }
  const context = vm.createContext({
    EMAIL_TO_LMS_BASE_URL: base,
    setTimeout: (fn, ms) => { timers.set(++nextTimer, { fn, ms }); return nextTimer; },
    clearTimeout: (id) => timers.delete(id),
    Office: {
      AsyncResultStatus: { Succeeded: "succeeded" }, CoercionType: { Text: "text" },
      MailboxEnums: { ComposeType: { Forward: "forward" } },
      context: { mailbox: { item, userProfile: { emailAddress: "user@example.com" } } },
      actions: { associate: (name, fn) => { context.handler = fn; } },
    },
    fetch: xhrOnly ? undefined : async (url, options) => {
      if (url === api && hangApi) { calls.push({ url }); return new Promise(() => {}); }
      const data = response(url, options);
      return { ok: true, status: 200, text: async () => JSON.stringify(data) };
    },
    XMLHttpRequest: class {
      open(method, url) { this.method = method; this.url = url; }
      setRequestHeader() {}
      send(body) {
        try {
          this.responseText = JSON.stringify(response(this.url, { method: this.method, body }));
          this.status = 200; this.onload();
        } catch { this.onerror(); }
      }
    },
  });
  // Deliberately no window, document, location, or Office.onReady.
  vm.runInContext(source, context);
  context.handler({ completed: (value) => completions.push(value) });
  return { context, calls, completions, headers, timers };
}
async function flush() { for (let i = 0; i < 40; i++) await Promise.resolve(); }

test("classic Windows runtime captures and submits once per recipient before completing", async () => {
  const r = runtime(); await flush();
  assert.equal(r.calls[0].url, base + "/api/captures");
  assert.equal(r.calls[0].payload.senderEmailId, "sender@example.com");
  assert.equal(r.calls.filter((c) => c.url === api).length, 2);
  assert.equal(r.headers[0]["X-LETTERID-CGB"], "letter-1,letter-2");
  assert.equal(r.calls.at(-1).url, base + "/api/letter-client-result");
  assert.equal(r.calls.at(-1).payload.ok, true);
  assert.equal(r.completions.length, 1);
  assert.equal(r.completions[0].allowEvent, true);
});
test("XHR fallback works without browser globals", async () => {
  const r = runtime({ xhrOnly: true }); await flush();
  assert.equal(r.calls.filter((c) => c.url === api).length, 2);
  assert.equal(r.completions.length, 1);
});
test("ordinary messages are not captured", async () => {
  const r = runtime({ forward: false, subject: "New message" }); await flush();
  assert.equal(r.calls.length, 0); assert.equal(r.completions.length, 1);
});
test("forwards without LMS are captured without an external POST", async () => {
  const r = runtime({ subject: "FW: Test" }); await flush();
  assert.equal(r.calls.length, 1); assert.equal(r.completions.length, 1);
});
test("subject fallback detects a forward and API errors are never retried", async () => {
  const r = runtime({ forward: false, failApi: true }); await flush();
  assert.equal(r.calls.filter((c) => c.url === api).length, 2);
  assert.equal(r.calls.at(-1).payload.ok, false);
  assert.equal(r.headers.length, 0); assert.equal(r.completions.length, 1);
});
test("a stalled subject call releases send exactly once", async () => {
  const r = runtime({ forward: false, hangSubject: true }); await flush();
  for (const timer of [...r.timers.values()]) timer.fn();
  assert.equal(r.completions.length, 1); assert.equal(r.calls.length, 0);
});
test("stalled external requests time out and report failure", async () => {
  const r = runtime({ hangApi: true }); await flush();
  for (const [id, timer] of [...r.timers]) {
    if (timer.ms === 5000) { r.timers.delete(id); timer.fn(); }
  }
  await flush();
  assert.equal(r.calls.at(-1).payload.ok, false);
  assert.equal(r.completions.length, 1);
});
