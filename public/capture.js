/* Shared by the Outlook task pane and the browser simulator. */
(function (root) {
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

  function splitAddresses(value) {
    return (value || "")
      .split(/[;,]+/)
      .map(function (part) {
        return parseMailbox(part).email || part.trim();
      })
      .filter(Boolean);
  }

  function parseLetterId(text) {
    if (!text) {
      return "";
    }
    var trimmed = String(text).trim();
    function fromObject(data) {
      if (!data || typeof data !== "object") {
        return "";
      }
      var keys = ["letterId", "LetterId", "letterID", "letter_id"];
      var i;
      for (i = 0; i < keys.length; i += 1) {
        if (data[keys[i]] != null && String(data[keys[i]]).trim()) {
          return String(data[keys[i]]).trim();
        }
      }
      if (data.data && typeof data.data === "object") {
        return fromObject(data.data);
      }
      if (data.result && typeof data.result === "object") {
        return fromObject(data.result);
      }
      return "";
    }
    try {
      var parsed = JSON.parse(trimmed);
      if (typeof parsed === "string" || typeof parsed === "number") {
        return String(parsed).trim();
      }
      return fromObject(parsed);
    } catch (ignore) {
      var match = trimmed.match(/letterId["'\s:=]+([A-Za-z0-9._-]+)/i);
      return match ? match[1] : "";
    }
  }

  function postCapture(payload, done) {
    var finished = false;
    function finish() {
      if (finished) {
        return;
      }
      finished = true;
      if (typeof done === "function") {
        done();
      }
    }

    setTimeout(finish, 8000);

    try {
      fetch("/api/captures", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      })
        .then(function (res) {
          return res.json();
        })
        .then(function (stored) {
          var shouldRun =
            stored && stored.letterSubmitShouldRun !== undefined
              ? stored.letterSubmitShouldRun
              : /\[LMS\]/i.test((payload && payload.subject) || "");
          if (
            !stored ||
            stored.letterSubmitFromServer ||
            !shouldRun ||
            (!stored.letterSubmit && !(stored.letterSubmits && stored.letterSubmits.length))
          ) {
            finish();
            return;
          }
          var url = stored.letterSubmitUrl;
          var payloads =
            stored.letterSubmits && stored.letterSubmits.length
              ? stored.letterSubmits.map(function (row) {
                  return row.fields;
                })
              : [stored.letterSubmit];
          var results = new Array(payloads.length);
          var left = payloads.length;
          function allDone() {
            var letterIds = [];
            var texts = [];
            var ok = false;
            var status = null;
            var error = null;
            results.forEach(function (result) {
              if (!result) {
                return;
              }
              if (result.ok) {
                ok = true;
              }
              var letterId = result.ok && !result.opaque ? parseLetterId(result.text) : "";
              if (letterId) {
                letterIds.push(letterId);
              }
              if (result.text) {
                texts.push(result.text);
              }
              if (result.status != null && status == null) {
                status = result.status;
              }
              if (result.error && !error) {
                error = result.error;
              }
            });
            var letterId = letterIds.join(",");
            function sendReport(headerSet) {
              fetch("/api/letter-client-result", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                  captureId: stored.id,
                  ok: ok,
                  status: status,
                  error: error,
                  url: url,
                  letterId: letterId,
                  headerSet: Boolean(headerSet),
                  responseText: texts.join("\n").slice(0, 4000),
                  opaque: false,
                }),
              }).then(finish, finish);
            }
            var item =
              window.Office && Office.context && Office.context.mailbox
                ? Office.context.mailbox.item
                : null;
            if (letterId && item && item.internetHeaders && item.internetHeaders.setAsync) {
              item.internetHeaders.setAsync(
                { "X-LETTERID-CGB": String(letterId) },
                function (asyncResult) {
                  sendReport(
                    asyncResult && asyncResult.status === Office.AsyncResultStatus.Succeeded,
                  );
                },
              );
              return;
            }
            sendReport(false);
          }
          payloads.forEach(function (fields, index) {
            fetch(url, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(fields),
            }).then(
              function (letterRes) {
                letterRes.text().then(
                  function (text) {
                    results[index] = { ok: letterRes.ok, status: letterRes.status, text: text };
                    left -= 1;
                    if (left <= 0) {
                      allDone();
                    }
                  },
                  function () {
                    results[index] = { ok: letterRes.ok, status: letterRes.status, text: "" };
                    left -= 1;
                    if (left <= 0) {
                      allDone();
                    }
                  },
                );
              },
              function (error) {
                fetch(url, {
                  method: "POST",
                  mode: "no-cors",
                  headers: { "Content-Type": "text/plain;charset=UTF-8" },
                  body: JSON.stringify(fields),
                }).then(
                  function () {
                    results[index] = {
                      ok: true,
                      status: 0,
                      opaque: true,
                      text: "",
                      error: error && error.message ? error.message : "letter fetch failed",
                    };
                    left -= 1;
                    if (left <= 0) {
                      allDone();
                    }
                  },
                  function () {
                    results[index] = {
                      ok: false,
                      error: error && error.message ? error.message : "letter fetch failed",
                    };
                    left -= 1;
                    if (left <= 0) {
                      allDone();
                    }
                  },
                );
              },
            );
          });
        }, finish);
    } catch (ignore) {
      finish();
    }
  }

  function getRecipients(recip, callback) {
    if (!recip || typeof recip.getAsync !== "function") {
      callback([]);
      return;
    }
    recip.getAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        callback([]);
        return;
      }
      callback(emailsFromRecipients(result.value));
    });
  }

  function collectFromOutlook(item, classification, done) {
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
      classification: classification || null,
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
      payload.originalToEmailAddresses =
        parsed.originalToEmailAddresses && parsed.originalToEmailAddresses.length
          ? parsed.originalToEmailAddresses
          : payload.originalToEmailAddresses;
      payload.forwardedByEmail = composeFrom.email;
      try {
        var profile =
          window.Office && Office.context && Office.context.mailbox
            ? Office.context.mailbox.userProfile
            : null;
        if (profile && profile.emailAddress && !payload.forwardedByEmail) {
          payload.forwardedByEmail = profile.emailAddress;
        }
      } catch (ignore) {}
      done(payload);
    }

    setTimeout(complete, 4000);

    function finishIfReady() {
      pending -= 1;
      if (pending > 0) {
        return;
      }
      complete();
    }

    if (item.dateTimeCreated) {
      try {
        payload.originalEmailDate = new Date(item.dateTimeCreated).toISOString();
      } catch (ignore) {}
    }

    if (item.subject && item.subject.getAsync) {
      item.subject.getAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          payload.subject = result.value || "";
        }
        finishIfReady();
      });
    } else {
      finishIfReady();
    }

    if (item.body && item.body.getAsync) {
      var coercion =
        window.Office && Office.CoercionType && Office.CoercionType.Text
          ? Office.CoercionType.Text
          : "text";
      item.body.getAsync(coercion, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          payload.messageBody = result.value || "";
        }
        finishIfReady();
      });
    } else {
      finishIfReady();
    }

    getRecipients(item.to, function (list) {
      payload.toEmailAddresses = list;
      finishIfReady();
    });
    getRecipients(item.cc, function (list) {
      payload.ccEmailAddresses = list;
      finishIfReady();
    });

    if (item.from && typeof item.from.getAsync === "function") {
      item.from.getAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
          composeFrom = {
            name: result.value.displayName || "",
            email: result.value.emailAddress || "",
          };
        }
        finishIfReady();
      });
    } else {
      finishIfReady();
    }
  }

  root.EmailToLMSCapture = {
    parseOriginalFromBody: parseOriginalFromBody,
    splitAddresses: splitAddresses,
    post: postCapture,
    collectFromOutlook: collectFromOutlook,
  };
})(window);
