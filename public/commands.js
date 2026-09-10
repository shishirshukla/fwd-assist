/* Loaded by commands.html when Outlook runs Open form (Smart Alerts Take Action). */

function completeFunction(event) {
  if (event && typeof event.completed === "function") {
    event.completed();
  }
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
