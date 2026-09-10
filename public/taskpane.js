(function () {
  var outlookReady = false;

  function todayIso() {
    var now = new Date();
    var month = String(now.getMonth() + 1).padStart(2, "0");
    var day = String(now.getDate()).padStart(2, "0");
    return now.getFullYear() + "-" + month + "-" + day;
  }

  function classificationLine(priority, endDate, category) {
    return (
      "[Forward classification: Priority=" +
      priority +
      " | End Date=" +
      endDate +
      " | Category=" +
      category +
      "]"
    );
  }

  function setError(id, message) {
    document.getElementById(id + "-error").textContent = message || "";
  }

  function setStatus(message) {
    document.getElementById("status").hidden = false;
    document.getElementById("status").textContent = message;
  }

  function enableSave() {
    document.getElementById("save").disabled = false;
  }

  function tryMessageParent(payload) {
    try {
      if (Office.context && Office.context.ui && typeof Office.context.ui.messageParent === "function") {
        Office.context.ui.messageParent(JSON.stringify(payload));
        return true;
      }
    } catch (ignore) {}
    return false;
  }

  function setSessionComplete(item, callback) {
    if (item.sessionData && typeof item.sessionData.setAsync === "function") {
      item.sessionData.setAsync("forwardMetadataComplete", "true", function () {
        callback();
      });
      return;
    }
    callback();
  }

  function sendComposeItem() {
    var item = Office.context.mailbox.item;
    if (!item || typeof item.sendAsync !== "function") {
      setStatus("Classification saved. Click Send to send the message.");
      enableSave();
      return;
    }

    item.sendAsync(function (sendResult) {
      if (sendResult.status === Office.AsyncResultStatus.Succeeded) {
        setStatus("Classification saved. Message sent.");
        return;
      }
      var detail =
        sendResult.error && sendResult.error.message
          ? sendResult.error.message
          : "Click Send to send the message.";
      setStatus("Classification saved. " + detail);
      enableSave();
    });
  }

  function continueAfterSave(priority, endDate, category) {
    document.getElementById("form").hidden = true;
    var handedOff = tryMessageParent({
      action: "classificationSaved",
      priority: priority,
      endDate: endDate,
      category: category,
    });
    if (handedOff) {
      setStatus("Classification saved. Sending…");
      return;
    }
    sendComposeItem();
  }

  var endDateInput = document.getElementById("endDate");
  endDateInput.min = todayIso();

  function persistInOutlook(priority, endDate, category) {
    var item = Office.context.mailbox.item;
    if (!item) {
      continueAfterSave(priority, endDate, category);
      return;
    }

    item.loadCustomPropertiesAsync(function (propResult) {
      if (propResult.status !== Office.AsyncResultStatus.Succeeded) {
        enableSave();
        setStatus("Could not save classification on this message.");
        return;
      }

      var props = propResult.value;
      props.set("forwardMetadataComplete", "true");
      props.set("forwardPriority", priority);
      props.set("forwardEndDate", endDate);
      props.set("forwardCategory", category);
      props.saveAsync(function (saveResult) {
        if (saveResult.status !== Office.AsyncResultStatus.Succeeded) {
          enableSave();
          setStatus("Could not persist classification. Try again.");
          return;
        }

        setSessionComplete(item, function () {
          item.body.prependAsync(
            "<p>" + classificationLine(priority, endDate, category) + "</p>",
            { coercionType: Office.CoercionType.Html },
            function () {
              function finish() {
                try {
                  item.notificationMessages.removeAsync("ForwardGuardNotice");
                } catch (ignore) {}
                continueAfterSave(priority, endDate, category);
              }

              if (item.internetHeaders && item.internetHeaders.setAsync) {
                item.internetHeaders.setAsync(
                  {
                    "X-Forward-Priority": priority,
                    "X-Forward-End-Date": endDate,
                    "X-Forward-Category": category,
                  },
                  function () {
                    finish();
                  }
                );
                return;
              }
              finish();
            }
          );
        });
      });
    });
  }

  document.getElementById("form").addEventListener("submit", function (event) {
    event.preventDefault();
    var priority = document.getElementById("priority").value;
    var endDate = document.getElementById("endDate").value;
    var category = document.getElementById("category").value;
    var valid = true;

    setError("priority", "");
    setError("endDate", "");
    setError("category", "");

    if (!priority) {
      setError("priority", "Choose a priority.");
      valid = false;
    }
    if (!endDate) {
      setError("endDate", "Choose an end date.");
      valid = false;
    } else if (endDate < todayIso()) {
      setError("endDate", "End date cannot be in the past.");
      valid = false;
    }
    if (!category) {
      setError("category", "Choose a category.");
      valid = false;
    }
    if (!valid) {
      return;
    }

    document.getElementById("save").disabled = true;

    if (outlookReady && window.Office && Office.context && Office.context.mailbox) {
      persistInOutlook(priority, endDate, category);
      return;
    }

    document.getElementById("form").hidden = true;
    setStatus("Classification saved. Message sent.");
  });

  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      outlookReady = info && info.host === Office.HostType.Outlook;
      if (!outlookReady || !Office.context || !Office.context.mailbox) {
        return;
      }
      var item = Office.context.mailbox.item;
      if (item && item.getInitializationContextAsync) {
        item.getInitializationContextAsync(function () {});
      }
    });
  }
})();
