/* ============================================================================
   commands.js
   Outlook add-in command handler for sending the currently opened email
   to a Power Automate HTTP-triggered cloud flow.
   ============================================================================ */

(function () {
    "use strict";

    // Resolving the dialog URL against window.location can throw in some
    // Outlook clients' hidden command-execution frame (location.href is not
    // always a normal resolvable page URL there). Guard it so a failure here
    // can never prevent the rest of this script (and the ribbon button
    // function) from loading.
    function resolveNpiDialogUrl() {
        try {
            return new URL("npi-prompt.html", window.location.href).href;
        } catch (e) {
            return "https://madhumuthuswamy.github.io/outlook-addin-sendtocrm/npi-prompt.html";
        }
    }

    const defaultConfig = {
        flowUrl: "https://REPLACE-WITH-YOUR-POWER-AUTOMATE-HTTP-POST-URL",
        sentFlagKey: "SentToFlow",
        sentTimestampKey: "SentToFlowTimestamp",
        notificationKey: "sendToFlowStatus",
        // Resolved relative to this page's own location by default, so the
        // dialog keeps working if the add-in is ever hosted elsewhere.
        npiDialogUrl: resolveNpiDialogUrl()
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

    async function sendEmailToFlow(event) {
        try {
            if (isSubmitting) {
                await notify("A send request is already in progress for this email.");
                return;
            }

            validateConfiguration();

            const item = Office.context.mailbox.item;
            validateItemContext(item);

            isSubmitting = true;

            const customProps = await loadCustomProperties(item);
            const alreadySent = customProps.get(config.sentFlagKey);

            if (alreadySent === true || alreadySent === "true") {
                const sentOn = customProps.get(config.sentTimestampKey) || "a previous time";
                await notify(
                    "This email was already sent to CRM on " + sentOn + ". Duplicate sends are blocked.",
                    true
                );
                return;
            }

            await notify("Looking for an NPI # in the email...");
            const bodyHtmlForScan = await getBody(item, Office.CoercionType.Html);
            const suggestedNpi = findCandidateNpi(item.subject, bodyHtmlForScan);

            await notify("Waiting for NPI # entry...");
            const npiNumber = await promptForNpi(suggestedNpi);

            await notify("Preparing email data...");
            const payload = await buildPayload(item, npiNumber);
            const flowResult = await postToFlow(payload);

            if (!flowResult.ok) {
                throw new Error(flowResult.message);
            }

            await markAsSent(customProps);

            if (flowResult.duplicate) {
                await notify("This email had already been processed. The request was not duplicated.");
            } else {
                await notify("Email and attachments were sent successfully.");
            }
        } catch (error) {
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

    async function buildPayload(item, npiNumber) {
        let bodyHtml = await getBody(item, Office.CoercionType.Html);
        const bodyText = await getBody(item, Office.CoercionType.Text);
        const attachments = await getAttachments(item);
        setInlineContentIds(bodyHtml, attachments);

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
            attachments: attachments,
            npiNumber: npiNumber || "",
            sentBy: getSentByInfo()
        };
    }

    // Identifies the Outlook user who actually clicked the ribbon button,
    // as distinct from the email's From/To/Cc. Office.js exposes the signed
    // -in mailbox owner via userProfile without requiring extra permissions.
    function getSentByInfo() {
        const profile = (Office.context.mailbox && Office.context.mailbox.userProfile) || {};
        return {
            displayName: profile.displayName || "",
            emailAddress: profile.emailAddress || "",
            timeZone: profile.timeZone || ""
        };
    }

    // Scans the subject and visible body text for an exact 10-digit number.
    // Hyperlinks and hidden HTML are removed before scanning so values from
    // URLs, tracking markup, and invisible email content are not suggested.
    function findCandidateNpi(subject, bodyHtml) {
        // Avoid lookbehind/lookahead regex assertions here: they are not
        // supported by some JS engines still used to render Outlook/OWA
        // (e.g. Internet Explorer/Trident), and an unsupported regex
        // literal causes a parse-time SyntaxError that silently kills this
        // entire script (no Office.onReady, no button handler at all).
        // Instead, find every run of digits and only accept one that is
        // exactly 10 digits long (so longer numbers are ignored).
        function searchText(text) {
            const digitRuns = (text || "").match(/[0-9]+/g) || [];
            for (const run of digitRuns) {
                if (run.length === 10) {
                    return run;
                }
            }
            return "";
        }

        return searchText(subject) || searchText(getVisibleBodyText(bodyHtml));
    }

    function getVisibleBodyText(bodyHtml) {
        if (!bodyHtml || typeof DOMParser !== "function") {
            return "";
        }

        const document = new DOMParser().parseFromString(bodyHtml, "text/html");
        if (!document || !document.body) {
            return "";
        }

        const excluded = document.querySelectorAll(
            "a, script, style, noscript, template, title, head, meta, link, svg, " +
            "[hidden], [aria-hidden='true'], " +
            "[style*='display:none'], [style*='display: none'], " +
            "[style*='visibility:hidden'], [style*='visibility: hidden']"
        );

        for (const element of excluded) {
            if (element.parentNode) {
                element.parentNode.removeChild(element);
            }
        }

        return document.body.textContent || "";
    }

    // Shows a modal dialog (npi-prompt.html) requiring the user to enter a
    // 10-digit NPI # before the send can proceed. Resolves with the NPI
    // string on success; rejects (blocking the send) if cancelled/closed.
    // An optional suggestedNpi is passed through as a query parameter so the
    // dialog can pre-fill the field (the user can still edit or clear it).
    // Opened as a real popup dialog (not displayInIframe) because this runs
    // from the invisible/background frame used for ribbon ExecuteFunction
    // commands; an in-frame dialog there would render but be invisible.
    function promptForNpi(suggestedNpi) {
        const dialogUrl = suggestedNpi
            ? config.npiDialogUrl + "?npi=" + encodeURIComponent(suggestedNpi)
            : config.npiDialogUrl;

        return new Promise(function (resolve, reject) {
            Office.context.ui.displayDialogAsync(
                dialogUrl,
                { height: 20, width: 20 },
                function (asyncResult) {
                    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                        reject(new Error("Unable to open the NPI # entry dialog."));
                        return;
                    }

                    const dialog = asyncResult.value;
                    let settled = false;

                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                        settled = true;

                        let message = null;
                        try {
                            message = JSON.parse(arg.message);
                        } catch (e) {
                            message = null;
                        }

                        dialog.close();

                        if (message && message.status === "ok" && /^\d{10}$/.test(message.npi)) {
                            resolve(message.npi);
                        } else {
                            reject(new Error("NPI # entry was cancelled. The email was not sent."));
                        }
                    });

                    dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
                        if (!settled) {
                            settled = true;
                            reject(new Error("NPI # entry was cancelled. The email was not sent."));
                        }
                    });
                }
            );
        });
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
        const attachments = [];
        for (const attachment of rawAttachments) {
            attachments.push(await getAttachmentContent(item, attachment));
        }

        return attachments;
    }

    function setInlineContentIds(bodyHtml, attachments) {
        if (!bodyHtml || !Array.isArray(attachments)) {
            return;
        }

        const contentIds = [];
        bodyHtml.replace(/cid:([^"'()\s>]+)/gi, function (match, value) {
            const contentId = decodeURIComponent(value);
            contentIds.push(contentId);
            return match;
        });

        for (const attachment of attachments) {
            if (!attachment.isInline || attachment.contentId) {
                continue;
            }

            const attachmentName = (attachment.name || "").toLowerCase();
            const matchingContentId = contentIds.find(function (contentId) {
                const cidName = contentId.split("@")[0].toLowerCase();
                return cidName === attachmentName || cidName === attachmentName.split("@")[0];
            });

            if (matchingContentId) {
                attachment.contentId = matchingContentId;
            }
        }
    }

    function getAttachmentContent(item, attachment) {
        // Some attachments (e.g. on IRM/encrypted or sensitivity-labeled
        // messages) do not expose a readable id, and content retrieval is
        // blocked by policy. Skip content for those instead of throwing,
        // so the rest of the email can still be sent.
        if (!attachment || !attachment.id) {
            return Promise.resolve({
                id: "",
                name: (attachment && attachment.name) || "",
                contentType: (attachment && attachment.contentType) || "application/octet-stream",
                size: (attachment && attachment.size) || 0,
                attachmentType: (attachment && attachment.attachmentType) || "",
                isInline: !!(attachment && attachment.isInline),
                contentId: (attachment && attachment.contentId) || "",
                contentFormat: "",
                contentBytes: "",
                contentUnavailable: true,
                contentUnavailableReason: "Attachment content could not be read (it may be protected/encrypted)."
            });
        }

        return new Promise(function (resolve) {
            item.getAttachmentContentAsync(attachment.id, function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    resolve({
                        id: attachment.id,
                        name: attachment.name || "",
                        contentType: attachment.contentType || "application/octet-stream",
                        size: attachment.size || 0,
                        attachmentType: attachment.attachmentType || "",
                        isInline: !!attachment.isInline,
                        contentId: attachment.contentId || "",
                        contentFormat: "",
                        contentBytes: "",
                        contentUnavailable: true,
                        contentUnavailableReason: (result.error && result.error.message) || "Unable to read attachment content."
                    });
                    return;
                }

                const content = result.value || {};
                resolve({
                    id: attachment.id,
                    name: attachment.name || "",
                    contentType: attachment.contentType || "application/octet-stream",
                    size: attachment.size || 0,
                    attachmentType: attachment.attachmentType || "",
                    isInline: !!attachment.isInline,
                    contentId: attachment.contentId || "",
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

    async function markAsSent(customProps) {
        customProps.set(config.sentFlagKey, true);
        customProps.set(config.sentTimestampKey, new Date().toLocaleString());

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

    function notify(message, persistent) {
        return new Promise(function (resolve) {
            Office.context.mailbox.item.notificationMessages.replaceAsync(
                config.notificationKey,
                {
                    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                    message: message,
                    icon: "Icon.16x16",
                    persistent: !!persistent
                },
                function () {
                    resolve();
                }
            );
        });
    }
})();
