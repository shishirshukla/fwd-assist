/* Outlook OnMessageSend (Smart Alerts).
   Do not open dialogs or the task pane from this handler. Those APIs are
   unsupported here and can stop the add-in from opening later.
   Outlook opens the pane when the user clicks Take Action / Open form,
   using commandId that must match a ShowTaskpane button in the manifest. */

function isForwardedSubject(subject) {
  return /^(fw|fwd)\s*:/i.test((subject || "").trim());
}

function allowSend(event) {
  event.completed({ allowEvent: true });
}

function completeBlock(event) {
  event.completed({
    allowEvent: false,
    errorMessage:
      "This is a forwarded email. Click Open form, complete Priority, End Date, and Category, then save to send.",
    errorMessageMarkdown:
      "This is a **forwarded email**.\n\nClick **Open form**, fill Priority, End Date, and Category, then save to send.",
    cancelLabel: "Open form",
    commandId: "msgComposeOpenFormButton",
    contextData: JSON.stringify({ reason: "forward-classification" }),
  });
}

function addInsight(item, callback) {
  var done = typeof callback === "function" ? callback : function () {};
  if (!item.notificationMessages || !item.notificationMessages.replaceAsync) {
    done();
    return;
  }

  try {
    item.notificationMessages.replaceAsync(
      "ForwardGuardNotice",
      {
        type: "insightMessage",
        message: "Fill Priority, End Date, and Category, then save to send.",
        icon: "Icon16",
        actions: [
          {
            actionText: "Open form",
            actionType: "showTaskPane",
            commandId: "msgComposeOpenPaneButton",
            contextData: "{\"reason\":\"forward-classification\"}",
          },
        ],
      },
      function (result) {
        if (result && result.status === Office.AsyncResultStatus.Succeeded) {
          done();
          return;
        }
        try {
          item.notificationMessages.replaceAsync(
            "ForwardGuardNotice",
            {
              type: "errorMessage",
              message:
                "Forwarded email: click Open form on the Send dialog, or Apps → Forward Guard.",
            },
            function () {
              done();
            }
          );
        } catch (ignore) {
          done();
        }
      }
    );
  } catch (ignore) {
    done();
  }
}

function blockSend(item, event) {
  var finished = false;
  function finish() {
    if (finished) {
      return;
    }
    finished = true;
    completeBlock(event);
  }

  addInsight(item, finish);
  setTimeout(finish, 1200);
}

function metadataIsComplete(customProps) {
  return customProps.get("forwardMetadataComplete") === "true";
}

function loadCustomPropertiesThenDecide(item, event) {
  item.loadCustomPropertiesAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      if (metadataIsComplete(result.value)) {
        allowSend(event);
        return;
      }
    }
    blockSend(item, event);
  });
}

function checkCustomPropertiesThenDecide(item, event) {
  if (item.sessionData && typeof item.sessionData.getAsync === "function") {
    item.sessionData.getAsync("forwardMetadataComplete", function (sessionResult) {
      if (
        sessionResult.status === Office.AsyncResultStatus.Succeeded &&
        sessionResult.value === "true"
      ) {
        allowSend(event);
        return;
      }
      loadCustomPropertiesThenDecide(item, event);
    });
    return;
  }
  loadCustomPropertiesThenDecide(item, event);
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

      if (forwarded) {
        checkCustomPropertiesThenDecide(item, event);
        return;
      }

      item.subject.getAsync(function (subjectResult) {
        var subject =
          subjectResult.status === Office.AsyncResultStatus.Succeeded
            ? subjectResult.value
            : "";
        if (isForwardedSubject(subject)) {
          checkCustomPropertiesThenDecide(item, event);
          return;
        }
        allowSend(event);
      });
    });
  } catch (ignore) {
    try {
      completeBlock(event);
    } catch (inner) {
      allowSend(event);
    }
  }
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
