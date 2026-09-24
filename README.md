# AresFit live-registry dialer bridge

The existing GitHub Pages mobile dialer is connected to an Apps Script backend while remaining on its original GitHub Pages origin. The Apps Script `LiveDialer.html` wrapper embeds the existing page and relays only queue reads and activity writes through a nonce-checked `postMessage` bridge. This keeps existing browser-local dialer storage on the same origin.

## Live flow

- Apps Script derives the active Google identity and checks the active `Users` row; request parameters do not choose the rep.
- The queue is restricted to that rep's currently assigned, unsuppressed registry rows, returned in pages of 50. Due callbacks and follow-ups are prioritized. Historical callability labels are informational and are not a fresh TPS/CTPS check.
- The dialer records locally first, assigns a stable event ID, and sends the event to Apps Script. It keeps pending events for retry and blocks moving to the next lead until the event has been acknowledged by the registry.
- Apps Script checks current ownership and suppression again, appends the activity idempotently to `Activity History`, and updates the corresponding latest-event and follow-up fields.
- On initial live load, an existing same-origin session is read first. Prior per-lead history is carried over only when the stable `Global_Site_ID` matches exactly; no phone or business-name fuzzy merge is used.

## Deployment and review status

The web app is deployed as Apps Script Version 3 at [AresFit live dialer](https://script.google.com/macros/s/AKfycbyd6EGoGOujdZOlCVpJ6ig1IFPGD3bII1cAny9bdXdi4q3GqBt8KHj2CcSvJNi7onHIkQ/exec?app=live-dialer). It targets native spreadsheet `1b1xaTfmWtlTlgx1nb7UJvBI2yx1FCKNpOWFgt9_RwKI`. The existing deployment settings remain Execute as Me (`michael@aresfit.co.uk`) and Anyone; the code requires Google's active signed-in identity to match an active AresFit `Users` row before returning leads or writing an event. An anonymous attempt using only `email=michael@aresfit.co.uk` was rejected.

The deployed URL and GitHub Pages returned HTTP 200, and the authenticated queue function completed in the Apps Script editor. The Codex in-app browser could not open the signed-in web app because it rewrote the URL to a failing `/macros/u/1/s/...` route. A complete signed-in UI read, activity write, retry, and Sheet readback remain unverified. No real lead event was written during deployment. Historical callability labels do not establish a fresh TPS/CTPS clearance.
