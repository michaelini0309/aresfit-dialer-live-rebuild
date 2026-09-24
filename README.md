# AresFit live-registry dialer bridge (local review draft)

This branch connects the existing GitHub Pages mobile dialer UI to a proposed Apps Script backend while keeping the dialer on its original GitHub Pages origin. The Apps Script `LiveDialer.html` wrapper embeds the existing page and relays only the queue-read and activity-write calls through a nonce-checked `postMessage` bridge. This keeps existing browser-local dialer storage on the same origin.

## Proposed live flow

- Apps Script derives the active Google identity and checks the active `Users` row; request parameters do not choose the rep.
- The queue is restricted to that rep's currently assigned, unsuppressed registry rows, returned in pages of 50. Due callbacks and follow-ups are prioritized. Historical callability labels are informational and are not a fresh TPS/CTPS check.
- The dialer records locally first, assigns a stable event ID, and sends the event to Apps Script. It keeps pending events for retry and blocks moving to the next lead until the event has been acknowledged by the registry.
- Apps Script checks current ownership and suppression again, appends the activity idempotently to `Activity History`, and updates the corresponding latest-event and follow-up fields.
- On initial live load, an existing same-origin session is read first. Prior per-lead history is carried over only when the stable `Global_Site_ID` matches exactly; no phone or business-name fuzzy merge is used.

## Deployment and review status

This is unpushed, undeployed source code. The Apps Script project and deployment URL have not been identified in the connected AresFit Drive, so no production endpoint is available to update or verify from this workspace. Do not use the branch as a live calling system until the business owner reviews the exact commit and deployment target, the bridge is deployed with access limited to the AresFit account/domain, and a signed-in queue-read, test-write, retry, and Sheet readback are verified.

No code was pushed, no deployment was changed, and no call event was written while preparing this branch. The native live registry was updated separately; its reconciliation report is maintained in the Codex workspace.
