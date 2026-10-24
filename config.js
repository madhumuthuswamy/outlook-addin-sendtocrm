window.OUTLOOK_ADDIN_CONFIG = {
    // The real Power Automate URL is never shipped to the browser.
    // commands.js posts to the same-origin "/api/sendToFlow" function,
    // which forwards to the Flow URL stored server-side (POWER_AUTOMATE_FLOW_URL app setting).
    sentFlagKey: "SentToFlow",
    sentTimestampKey: "SentToFlowTimestamp",
    sentMessageIdKey: "SentToFlowInternetMessageId",
    notificationKey: "sendToFlowStatus"
};
