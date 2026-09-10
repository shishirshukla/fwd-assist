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

  var endDateInput = document.getElementById("endDate");
  endDateInput.min = todayIso();

  function persistInOutlook(priority, endDate, category) {
    var item = Office.context.mailbox.item;
    var saveButton = document.getElementById("save");
    item.loadCustomPropertiesAsync(function (propResult) {
      if (propResult.status !== Office.AsyncResultStatus.Succeeded) {
        saveButton.disabled = false;
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
          saveButton.disabled = false;
          setStatus("Could not persist classification. Try again.");
          return;
        }

        item.body.prependAsync(
          "<p>" + classificationLine(priority, endDate, category) + "</p>",
          { coercionType: Office.CoercionType.Html },
          function () {
            if (item.internetHeaders && item.internetHeaders.setAsync) {
              item.internetHeaders.setAsync({
                "X-Forward-Priority": priority,
                "X-Forward-End-Date": endDate,
                "X-Forward-Category": category,
              });
            }
            try {
              item.notificationMessages.removeAsync("ForwardGuardNotice");
            } catch (ignore) {}
            document.getElementById("form").hidden = true;
            setStatus("Classification saved. You can send this forwarded message now.");
          }
        );
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
    setStatus("Classification saved. You can send this forwarded message now.");
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
