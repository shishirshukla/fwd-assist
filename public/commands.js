/* Loaded by commands.html when Outlook runs Open form (Smart Alerts Take Action). */

function completeFunction(event) {
  if (event && typeof event.completed === "function") {
    event.completed();
  }
}

function sendComposeItem(callback) {
  var done = typeof callback === "function" ? callback : function () {};
  var item =
    Office.context && Office.context.mailbox ? Office.context.mailbox.item : null;
  if (!item || typeof item.sendAsync !== "function") {
    done();
    return;
  }
  item.sendAsync(function () {
    done();
  });
}

function persistThenSend(data, callback) {
  var item =
    Office.context && Office.context.mailbox ? Office.context.mailbox.item : null;
  if (!item) {
    sendComposeItem(callback);
    return;
  }

  item.loadCustomPropertiesAsync(function (propResult) {
    if (propResult.status !== Office.AsyncResultStatus.Succeeded) {
      sendComposeItem(callback);
      return;
    }

    var props = propResult.value;
    props.set("forwardMetadataComplete", "true");
    if (data && data.priority) {
      props.set("forwardPriority", data.priority);
    }
    if (data && data.endDate) {
      props.set("forwardEndDate", data.endDate);
    }
    if (data && data.category) {
      props.set("forwardCategory", data.category);
    }
    props.saveAsync(function () {
      if (item.sessionData && typeof item.sessionData.setAsync === "function") {
        item.sessionData.setAsync("forwardMetadataComplete", "true", function () {
          sendComposeItem(callback);
        });
        return;
      }
      sendComposeItem(callback);
    });
  });
}

function handleDialogMessage(dialog, arg) {
  var data = {};
  try {
    data = JSON.parse(arg.message || "{}");
  } catch (ignore) {
    return;
  }
  if (data.action !== "classificationSaved") {
    return;
  }
  try {
    dialog.close();
  } catch (ignore) {}
  persistThenSend(data);
}

function tryShowTaskpane() {
  try {
    if (Office.addin && typeof Office.addin.showAsTaskpane === "function") {
      return Office.addin.showAsTaskpane();
    }
  } catch (ignore) {}
  return Promise.reject(new Error("showAsTaskpane unavailable"));
}

function tryShowDialog() {
  return new Promise(function (resolve, reject) {
    try {
      if (!Office.context || !Office.context.ui || !Office.context.ui.displayDialogAsync) {
        reject(new Error("dialog unavailable"));
        return;
      }
      var origin =
        typeof window !== "undefined" && window.location && window.location.origin
          ? window.location.origin
          : "";
      if (!origin) {
        reject(new Error("no origin"));
        return;
      }
      Office.context.ui.displayDialogAsync(
        origin + "/taskpane.html",
        { height: 58, width: 36, displayInIframe: true, promptBeforeOpen: false },
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            var dialog = result.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
              handleDialogMessage(dialog, arg);
            });
            resolve();
          } else {
            reject(result.error || new Error("dialog failed"));
          }
        }
      );
    } catch (error) {
      reject(error);
    }
  });
}

function openClassificationForm(event) {
  tryShowTaskpane().then(
    function () {
      completeFunction(event);
    },
    function () {
      tryShowDialog().then(
        function () {
          completeFunction(event);
        },
        function () {
          completeFunction(event);
        }
      );
    }
  );
}

function associate() {
  if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
    Office.actions.associate("openClassificationForm", openClassificationForm);
  }
}

associate();
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(associate);
}
