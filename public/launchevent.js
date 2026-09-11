/* OnMessageSend: if the message is a forward, capture fields and POST /api/captures,
   then allow send. No form and no send block. */

function isForwardedSubject(subject) {
  return /^(fw|fwd)\s*:/i.test((subject || "").trim());
}

function allowSend(event) {
  try {
    event.completed({ allowEvent: true });
  } catch (ignore) {}
}

function parseMailbox(line) {
  var trimmed = (line || "").trim();
  var angled = trimmed.match(/^(.*?)\s*<([^>]+)>/);
  if (angled) {
    return {
      name: angled[1].replace(/^["']|["']$/g, "").trim(),
      email: angled[2].trim(),
    };
  }
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmed)) {
    return { name: "", email: trimmed };
  }
  return { name: trimmed, email: "" };
}

function matchHeader(text, name) {
  var re = new RegExp("^" + name + ":\\s*(.+)$", "im");
  var match = text.match(re);
  return match ? match[1].trim() : "";
}

function splitAddresses(value) {
  return (value || "")
    .split(/[;,]+/)
    .map(function (part) {
      return parseMailbox(part).email || part.trim();
    })
    .filter(Boolean);
}

function parseOriginalFromBody(body) {
  var text = (body || "").replace(/\r\n/g, "\n");
  var from = parseMailbox(matchHeader(text, "From"));
  return {
    originalEmailDate: matchHeader(text, "Sent") || matchHeader(text, "Date"),
    senderName: from.name,
    senderEmailId: from.email,
    originalToEmailAddresses: splitAddresses(matchHeader(text, "To")),
  };
}

function emailsFromRecipients(list) {
  var result = [];
  (list || []).forEach(function (entry) {
    var email = entry && (entry.emailAddress || entry.email || "");
    if (email) {
      result.push(email);
    }
  });
  return result;
}

function addInOrigin() {
  try {
    if (typeof location !== "undefined" && location.origin) {
      return location.origin;
    }
  } catch (ignore) {}
  return "";
}

function captureUrl() {
  var origin = addInOrigin();
  return origin ? origin + "/api/captures" : "/api/captures";
}

function postJson(url, body, headers, mode, done) {
  try {
    if (typeof fetch === "function") {
      fetch(url, {
        method: "POST",
        mode: mode || "cors",
        headers: headers,
        body: body,
      }).then(
        function (res) {
          if (mode === "no-cors") {
            done({ ok: true, status: 0, opaque: true, text: "" });
            return;
          }
          res.text().then(
            function (text) {
              done({ ok: res.ok, status: res.status, text: text || "" });
            },
            function () {
              done({ ok: res.ok, status: res.status, text: "" });
            },
          );
        },
        function (error) {
          done({
            ok: false,
            status: null,
            error: error && error.message ? error.message : "fetch failed",
          });
        },
      );
      return;
    }
  } catch (ignore) {}

  try {
    var xhr = new XMLHttpRequest();
    xhr.open("POST", url, true);
    var headerName;
    for (headerName in headers) {
      if (Object.prototype.hasOwnProperty.call(headers, headerName)) {
        xhr.setRequestHeader(headerName, headers[headerName]);
      }
    }
    xhr.onload = function () {
      done({
        ok: xhr.status >= 200 && xhr.status < 300,
        status: xhr.status,
        text: xhr.responseText || "",
      });
    };
    xhr.onerror = function () {
      done({ ok: false, status: null, error: "xhr failed" });
    };
    xhr.send(body);
  } catch (ignore2) {
    done({ ok: false, status: null, error: "post failed" });
  }
}

function postCapture(payload, done) {
  var finished = false;
  function finish(result) {
    if (finished) {
      return;
    }
    finished = true;
    if (typeof done === "function") {
      done(result || null);
    }
  }

  setTimeout(function () {
    finish(null);
  }, 4000);

  postJson(
    captureUrl(),
    JSON.stringify(payload),
    { "Content-Type": "application/json" },
    "cors",
    function (result) {
      if (!result || !result.ok) {
        finish(null);
        return;
      }
      try {
        finish(JSON.parse(result.text));
      } catch (ignore) {
        finish(null);
      }
    },
  );
}

function postLetterFromBrowser(url, fields, done) {
  var body = JSON.stringify(fields);
  function nextPlain() {
    postJson(
      url,
      body,
      { "Content-Type": "text/plain;charset=UTF-8" },
      "cors",
      function (result) {
        if (result && (result.ok || result.status)) {
          done(result);
          return;
        }
        postJson(
          url,
          body,
          { "Content-Type": "text/plain;charset=UTF-8" },
          "no-cors",
          done,
        );
      },
    );
  }

  postJson(url, body, { "Content-Type": "application/json" }, "cors", function (result) {
    if (result && result.ok) {
      done(result);
      return;
    }
    if (result && result.status && result.status >= 400 && result.status < 600) {
      done(result);
      return;
    }
    nextPlain();
  });
}

function reportLetterResult(captureId, url, result, done) {
  postJson(
    addInOrigin() + "/api/letter-client-result",
    JSON.stringify({
      captureId: captureId || "",
      ok: Boolean(result && result.ok),
      status: result ? result.status : null,
      error: result && result.error ? result.error : null,
      url: url,
      responseText: result && result.text ? String(result.text).slice(0, 4000) : "",
      opaque: Boolean(result && result.opaque),
    }),
    { "Content-Type": "application/json" },
    "cors",
    function () {
      if (typeof done === "function") {
        done();
      }
    },
  );
}

function getRecipients(recip, callback) {
  if (!recip || typeof recip.getAsync !== "function") {
    callback([]);
    return;
  }
  recip.getAsync(function (result) {
    if (!result || result.status !== Office.AsyncResultStatus.Succeeded) {
      callback([]);
      return;
    }
    callback(emailsFromRecipients(result.value));
  });
}

function collectPayload(item, done) {
  var payload = {
    originalEmailDate: "",
    senderEmailId: "",
    senderName: "",
    subject: "",
    messageBody: "",
    toEmailAddresses: [],
    ccEmailAddresses: [],
    originalToEmailAddresses: [],
    forwardedByEmail: "",
  };
  var composeFrom = { name: "", email: "" };
  var pending = 5;
  var finished = false;

  function complete() {
    if (finished) {
      return;
    }
    finished = true;
    var parsed = parseOriginalFromBody(payload.messageBody);
    payload.originalEmailDate = parsed.originalEmailDate || payload.originalEmailDate;
    payload.senderEmailId = parsed.senderEmailId;
    payload.senderName = parsed.senderName;
    payload.originalToEmailAddresses = parsed.originalToEmailAddresses || [];
    payload.forwardedByEmail = composeFrom.email;
    try {
      var profile = Office.context.mailbox.userProfile;
      if (profile && profile.emailAddress && !payload.forwardedByEmail) {
        payload.forwardedByEmail = profile.emailAddress;
      }
    } catch (ignore) {}
    done(payload);
  }

  setTimeout(complete, 3500);

  function tick() {
    pending -= 1;
    if (pending <= 0) {
      complete();
    }
  }

  if (item.subject && item.subject.getAsync) {
    item.subject.getAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        payload.subject = result.value || "";
      }
      tick();
    });
  } else {
    tick();
  }

  if (item.body && item.body.getAsync) {
    var coercion =
      Office.CoercionType && Office.CoercionType.Text ? Office.CoercionType.Text : "text";
    item.body.getAsync(coercion, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        payload.messageBody = result.value || "";
      }
      tick();
    });
  } else {
    tick();
  }

  getRecipients(item.to, function (list) {
    payload.toEmailAddresses = list;
    tick();
  });
  getRecipients(item.cc, function (list) {
    payload.ccEmailAddresses = list;
    tick();
  });

  if (item.from && typeof item.from.getAsync === "function") {
    item.from.getAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        composeFrom = {
          name: result.value.displayName || "",
          email: result.value.emailAddress || "",
        };
      }
      tick();
    });
  } else {
    tick();
  }
}

function captureForwardThenAllow(item, event) {
  var done = false;
  function finish() {
    if (done) {
      return;
    }
    done = true;
    allowSend(event);
  }

  setTimeout(finish, 5000);
  collectPayload(item, function (payload) {
    postCapture(payload, function (stored) {
      var fields = stored && stored.letterSubmit;
      var url =
        (stored && stored.letterSubmitUrl) ||
        "https://eloan.cgbankmobile.in/pensioner_api/auth/api/submit-letter";
      if (!fields || stored.letterSubmitFromServer) {
        finish();
        return;
      }
      postLetterFromBrowser(url, fields, function (letterResult) {
        reportLetterResult(stored && stored.id, url, letterResult, finish);
      });
    });
  });
}

function onMessageSendHandler(event) {
  try {
    var item = Office.context.mailbox.item;
    if (!item) {
      allowSend(event);
      return;
    }

    item.getComposeTypeAsync(function (composeResult) {
      var forwarded = false;
      if (
        composeResult.status === Office.AsyncResultStatus.Succeeded &&
        composeResult.value &&
        composeResult.value.composeType === Office.MailboxEnums.ComposeType.Forward
      ) {
        forwarded = true;
      }

      function afterSubject(subject) {
        if (forwarded || isForwardedSubject(subject)) {
          captureForwardThenAllow(item, event);
          return;
        }
        allowSend(event);
      }

      if (forwarded) {
        afterSubject("");
        return;
      }

      item.subject.getAsync(function (subjectResult) {
        afterSubject(
          subjectResult.status === Office.AsyncResultStatus.Succeeded
            ? subjectResult.value
            : "",
        );
      });
    });
  } catch (ignore) {
    allowSend(event);
  }
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
