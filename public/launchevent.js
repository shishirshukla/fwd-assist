/* Office event runtime — no DOM. Loaded by Outlook OnMessageSend / ItemSend. */

function isForwardedSubject(subject) {
  return /^(fw|fwd)\s*:/i.test((subject || "").trim());
}

function allowSend(event) {
  event.completed({ allowEvent: true });
}

function blockSend(event) {
  event.completed({
    allowEvent: false,
    errorMessage:
      "This is a forwarded email. Complete Priority, End Date, and Category, then send again.",
    cancelLabel: "Don't Send",
    commandId: "msgComposeOpenPaneButton",
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
    blockSend(event);
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
