Office.onReady(() => {});

const DISCLAIMER_ID = "CorpDisclaimer_v1";

const DISCLAIMER_HTML = `
<div id="${DISCLAIMER_ID}"
     style="font-size:9pt;
            color:#666666;
            font-family:Segoe UI;
            margin-top:12px;
            border-top:1px solid #cccccc;
            padding-top:8px;">

This email and any attachments are confidential and intended solely
for the intended recipient. If you have received this communication
in error, please notify the sender and delete it immediately.

</div>
`;

function onComposeHandler(event) {

    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        function(result) {

            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                event.completed();
                return;
            }

            const body = result.value || "";

            // Prevent duplicate disclaimer
            if (body.includes(DISCLAIMER_ID)) {
                event.completed();
                return;
            }

            // Insert disclaimer at top of compose body
            Office.context.mailbox.item.body.prependAsync(
                DISCLAIMER_HTML,
                { coercionType: Office.CoercionType.Html },
                function() {
                    event.completed();
                }
            );
        }
    );
}

// Required for event-based activation
Office.actions.associate("onComposeHandler", onComposeHandler);