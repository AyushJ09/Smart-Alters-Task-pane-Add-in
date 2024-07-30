Office.onReady();

/**
 * @param {Office.MailboxEvent} event The onMessageSendHandler event object.
 */
const onMessageSendHandler = (event) => {
  const options = {
    allowEvent: false,
    errorMessage: "Error Occurs",
    cancelLabel: "Open taskpane",
    commandId: "msgReadOpenPaneButton",
  };

  event.completed(options);
  return;
};

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
