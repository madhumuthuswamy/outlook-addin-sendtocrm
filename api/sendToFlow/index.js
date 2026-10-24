// Server-side proxy: keeps the Power Automate HTTP-trigger URL (with its SAS "sig")
// out of the browser. Configure the real URL as an Application Setting named
// POWER_AUTOMATE_FLOW_URL on the Static Web App (never commit it to source control).
module.exports = async function (context, req) {
    const flowUrl = process.env.POWER_AUTOMATE_FLOW_URL;

    if (!flowUrl) {
        context.res = {
            status: 500,
            headers: { "Content-Type": "application/json" },
            body: { status: "error", message: "POWER_AUTOMATE_FLOW_URL is not configured on the server." }
        };
        return;
    }

    try {
        const response = await fetch(flowUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(req.body)
        });

        const text = await response.text();

        context.res = {
            status: response.status,
            headers: { "Content-Type": "application/json" },
            body: text
        };
    } catch (err) {
        context.res = {
            status: 502,
            headers: { "Content-Type": "application/json" },
            body: { status: "error", message: "Unable to reach Power Automate: " + err.message }
        };
    }
};
