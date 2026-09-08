(function () {
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

  function inOutlook() {
    return /[?&]_host_Info=/i.test(window.location.search);
  }

  function loadOfficeJs(callback) {
    if (window.Office) {
      callback();
      return;
    }
    var script = document.createElement("script");
    script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
    script.onload = callback;
    script.onerror = function () {
      document.getElementById("status").hidden = false;
      document.getElementById("status").textContent =
        "Office.js could not be loaded.";
    };
    document.head.appendChild(script);
  }

  function setError(id, message) {
    document.getElementById(id + "-error").textContent = message || "";
  }

  var endDateInput = document.getElementById("endDate");
  endDateInput.min = todayIso();

  function persistInOutlook(priority, endDate, category) {
    var item = Office.context.mailbox.item;
    var saveButton = document.getElementById("save");
    item.loadCustomPropertiesAsync(function (propResult) {
      if (propResult.status !== Office.AsyncResultStatus.Succeeded) {
        saveButton.disabled = false;
        document.getElementById("status").hidden = false;
        document.getElementById("status").textContent =
          "Could not save classification on this message.";
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
          document.getElementById("status").hidden = false;
          document.getElementById("status").textContent =
            "Could not persist classification. Try again.";
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
            document.getElementById("form").hidden = true;
            document.getElementById("status").hidden = false;
            document.getElementById("status").textContent =
              "Classification saved. You can send this forwarded message now.";
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

    if (!inOutlook() || !window.Office) {
      document.getElementById("form").hidden = true;
      document.getElementById("status").hidden = false;
      document.getElementById("status").textContent =
        "Classification saved. You can send this forwarded message now.";
      return;
    }

    persistInOutlook(priority, endDate, category);
  });

  if (inOutlook()) {
    loadOfficeJs(function () {
      if (window.Office && Office.onReady) {
        Office.onReady(function () {});
      }
    });
  }
})();
