/* Office event runtime. Outlook on the web cannot auto-open a task pane from Send.
   Smart Alerts must offer Take Action (commandId + cancelLabel) to open the pane. */

function isForwardedSubject(subject) {
  return /^(fw|fwd)\s*:/i.test((subject || "").trim());
}

function allowSend(event) {
  event.completed({ allowEvent: true });
}

function blockSend(item, event) {
  try {
    if (item.notificationMessages && item.notificationMessages.replaceAsync) {
      item.notificationMessages.replaceAsync(
        "ForwardGuardNotice",
        {
          type: "errorMessage",
          message:
            "Forwarded email: in the Send dialog choose Open form, then fill Priority, End Date, and Category.",
        },
        function () {}
      );
    }
  } catch (ignore) {}

  event.completed({
    allowEvent: false,
    errorMessage:
      "This is a forwarded email. Select Open form, complete Priority, End Date, and Category, then send again.",
    errorMessageMarkdown:
      "This is a **forwarded email**.\n\nSelect **Open form**, fill Priority, End Date, and Category, then send again.",
    cancelLabel: "Open form",
    commandId: "msgComposeOpenPaneButton",
    contextData: JSON.stringify({ reason: "forward-classification" }),
  });
}

function metadataIsComplete(customProps) {
  return customProps.get("forwardMetadataComplete") === "true";
}

function checkCustomPropertiesThenDecide(item, event) {
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

function onMessageSendHandler(event) {
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
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
