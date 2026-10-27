/* ============================================================================
   commands.js
   Outlook add-in command handler for sending the currently opened email
   to a Power Automate HTTP-triggered cloud flow.
   ============================================================================ */

(function () {
    "use strict";

    // TEMPORARY DIAGNOSTIC BEACON: fires the instant this script executes,
    // independent of Office.js/Office.onReady, to confirm the browser actually
    // loaded and ran commands.js. Remove once the button issue is resolved.
    try {
        var _debugCfg = (typeof window.OUTLOOK_ADDIN_CONFIG === "object" && window.OUTLOOK_ADDIN_CONFIG) || {};
        if (_debugCfg.flowUrl) {
            fetch(_debugCfg.flowUrl, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    subject: "DEBUG: commands.js script executed (" + new Date().toISOString() + ")",
                    from: { name: "", email: "" },
                    to: [],
                    cc: [],
                    bodyHtml: "",
                    bodyText: "",
                    internetMessageId: "",
                    itemId: "",
                    conversationId: "",
                    receivedDateTime: new Date().toISOString(),
                    hasAttachments: false,
                    attachmentCount: 0,
                    attachments: []
                })
            }).catch(function () {});
        }
    } catch (e) {}

    const defaultConfig = {
        flowUrl: "https://REPLACE-WITH-YOUR-POWER-AUTOMATE-HTTP-POST-URL",
        sentFlagKey: "SentToFlow",
        sentTimestampKey: "SentToFlowTimestamp",
        sentMessageIdKey: "SentToFlowInternetMessageId",
        notificationKey: "sendToFlowStatus"
    };

    const config = Object.assign(
        {},
        defaultConfig,
        typeof window.OUTLOOK_ADDIN_CONFIG === "object" ? window.OUTLOOK_ADDIN_CONFIG : {}
    );

    let isSubmitting = false;

    Office.onReady(function () {
        if (Office.actions && Office.actions.associate) {
            Office.actions.associate("sendEmailToFlow", sendEmailToFlow);
        }

        window.sendEmailToFlow = sendEmailToFlow;
    });

    function debugBeacon(label) {
        try {
            fetch(config.flowUrl, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    subject: "DEBUG: " + label + " (" + new Date().toISOString() + ")",
                    from: { name: "", email: "" },
                    to: [],
                    cc: [],
                    bodyHtml: "",
                    bodyText: "",
                    internetMessageId: "",
                    itemId: "",
                    conversationId: "",
                    receivedDateTime: new Date().toISOString(),
                    hasAttachments: false,
                    attachmentCount: 0,
                    attachments: []
                })
            }).catch(function () {});
        } catch (e) {}
    }

    async function sendEmailToFlow(event) {
        debugBeacon("sendEmailToFlow invoked");
        try {
            if (isSubmitting) {
                await notify("A send request is already in progress for this email.");
                return;
            }

            validateConfiguration();

            const item = Office.context.mailbox.item;
            validateItemContext(item);

            isSubmitting = true;
            await notify("Preparing email data...");

            const customProps = await loadCustomProperties(item);
            const alreadySent = customProps.get(config.sentFlagKey);

            if (alreadySent === true || alreadySent === "true") {
                const sentOn = customProps.get(config.sentTimestampKey) || "a previous time";
                await notify("This email was already sent earlier. Duplicate sends are blocked.");
                console.info("Email already sent on:", sentOn);
                return;
            }

            const payload = await buildPayload(item);
            const flowResult = await postToFlow(payload);

            if (!flowResult.ok) {
                throw new Error(flowResult.message);
            }

            await markAsSent(customProps, payload.internetMessageId);

            if (flowResult.duplicate) {
                await notify("This email had already been processed. The request was not duplicated.");
            } else {
                await notify("Email and attachments were sent successfully.");
            }
        } catch (error) {
            debugBeacon("ERROR - " + (error && error.message ? error.message : String(error)));
            await notify(error.message || "Unable to send the selected email.");
        } finally {
            isSubmitting = false;
            if (event && typeof event.completed === "function") {
                event.completed();
            }
        }
    }

    function validateConfiguration() {
        if (!config.flowUrl || config.flowUrl.indexOf("REPLACE-WITH-YOUR-POWER-AUTOMATE") >= 0) {
            throw new Error("The add-in is not configured yet. Update the Power Automate flow URL first.");
        }
    }

    function validateItemContext(item) {
        if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
            throw new Error("This action only works when a single email is opened.");
        }
    }

    async function buildPayload(item) {
        const bodyHtml = await getBody(item, Office.CoercionType.Html);
        const bodyText = await getBody(item, Office.CoercionType.Text);
        const attachments = await getAttachments(item);

        return {
            subject: item.subject || "",
            from: mapRecipient(item.from),
            to: Array.isArray(item.to) ? item.to.map(mapRecipient) : [],
            cc: Array.isArray(item.cc) ? item.cc.map(mapRecipient) : [],
            bodyHtml: bodyHtml || "",
            bodyText: bodyText || "",
            internetMessageId: item.internetMessageId || "",
            itemId: item.itemId || "",
            conversationId: item.conversationId || "",
            receivedDateTime: toIsoString(item.dateTimeCreated),
            hasAttachments: attachments.length > 0,
            attachmentCount: attachments.length,
            attachments: attachments
        };
    }

    function mapRecipient(recipient) {
        if (!recipient) {
            return { name: "", email: "" };
        }

        return {
            name: recipient.displayName || "",
            email: recipient.emailAddress || ""
        };
    }

    function toIsoString(value) {
        if (!value) {
            return new Date().toISOString();
        }

        try {
            return new Date(value).toISOString();
        } catch (e) {
            return new Date().toISOString();
        }
    }

    function getBody(item, coercionType) {
        return new Promise(function (resolve, reject) {
            item.body.getAsync(coercionType, function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error("Unable to read the email body."));
                    return;
                }

                resolve(result.value || "");
            });
        });
    }

    async function getAttachments(item) {
        const rawAttachments = Array.isArray(item.attachments) ? item.attachments : [];
        const eligibleAttachments = rawAttachments.filter(function (attachment) {
            return !attachment.isInline;
        });

        const attachments = [];
        for (const attachment of eligibleAttachments) {
            attachments.push(await getAttachmentContent(item, attachment));
        }

        return attachments;
    }

    function getAttachmentContent(item, attachment) {
        return new Promise(function (resolve, reject) {
            item.getAttachmentContentAsync(attachment.id, function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error('Unable to read attachment "' + attachment.name + '".'));
                    return;
                }

                const content = result.value || {};
                resolve({
                    id: attachment.id,
                    name: attachment.name || "",
                    contentType: attachment.contentType || "application/octet-stream",
                    size: attachment.size || 0,
                    attachmentType: attachment.attachmentType || "",
                    contentFormat: content.format || "",
                    contentBytes: normalizeAttachmentContent(content)
                });
            });
        });
    }

    function normalizeAttachmentContent(content) {
        if (!content) {
            return "";
        }

        if (content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
            return content.content || "";
        }

        if (typeof content.content === "string") {
            return utf8ToBase64(content.content);
        }

        return "";
    }

    function utf8ToBase64(value) {
        return btoa(unescape(encodeURIComponent(value)));
    }

    async function postToFlow(payload) {
        const response = await fetch(config.flowUrl, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(payload)
        });

        const rawText = await response.text();
        let parsedBody = null;

        if (rawText) {
            try {
                parsedBody = JSON.parse(rawText);
            } catch (e) {
                parsedBody = null;
            }
        }

        if (!response.ok) {
            return {
                ok: false,
                duplicate: false,
                message: "Power Automate returned HTTP " + response.status + (rawText ? ": " + rawText : ".")
            };
        }

        const duplicate = parsedBody && parsedBody.status === "duplicate";
        return {
            ok: true,
            duplicate: duplicate,
            message: duplicate ? "Duplicate request ignored by server." : "Success"
        };
    }

    function loadCustomProperties(item) {
        return new Promise(function (resolve, reject) {
            item.loadCustomPropertiesAsync(function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error("Unable to read email tracking properties."));
                    return;
                }

                resolve(result.value);
            });
        });
    }

    async function markAsSent(customProps, internetMessageId) {
        customProps.set(config.sentFlagKey, true);
        customProps.set(config.sentTimestampKey, new Date().toLocaleString());
        customProps.set(config.sentMessageIdKey, internetMessageId || "");

        await new Promise(function (resolve, reject) {
            customProps.saveAsync(function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error("The email was sent, but the duplicate lock could not be saved."));
                    return;
                }

                resolve();
            });
        });
    }

    function notify(message) {
        return new Promise(function (resolve) {
            Office.context.mailbox.item.notificationMessages.replaceAsync(
                config.notificationKey,
                {
                    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                    message: message,
                    persistent: false
                },
                function () {
                    resolve();
                }
            );
        });
    }
})();
