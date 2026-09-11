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

    setTimeout(finish, 5000);

    try {
      fetch("/api/captures", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      }).then(finish, finish);
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

  root.ForwardGuardCapture = {
    parseOriginalFromBody: parseOriginalFromBody,
    splitAddresses: splitAddresses,
    post: postCapture,
    collectFromOutlook: collectFromOutlook,
  };
})(window);
