# QuickBooks Desktop 2024 → PLS Accounts (Web Connector)

This build adds an initial QuickBooks Desktop Web Connector bridge. It does not require Excel export for the connected P&L report. The Windows PC initiates HTTPS requests to the VPS; do not expose QuickBooks or Windows file shares to the internet.

## 1. Configure the VPS

Set these environment variables in your VPS service configuration:

- `PLS_PUBLIC_URL=https://YOUR-PLS-DOMAIN` (no trailing slash; must be a valid publicly trusted HTTPS URL)
- `QBWC_USERNAME=pls`
- `QBWC_PASSWORD=<generate a long random password and keep it private>`

Example password generation on a trusted machine: `python -c "import secrets; print(secrets.token_urlsafe(32))"`

Restart the FastAPI service after setting the variables. The app must be reachable at `https://YOUR-PLS-DOMAIN/qbwc` from the Windows PC. Configure the VPS firewall/reverse proxy for HTTPS (443); do not use a self-signed certificate.

## 2. Install on the Windows PC

1. Open QuickBooks Desktop 2024 and sign in to the intended company file as the QuickBooks Admin user.
2. Switch to Single-user Mode for the initial authorization.
3. In PLS Accounts, sign in as the PLS administrator and download `PLS_Accounts.qwc` from the QuickBooks panel.
4. Open QuickBooks Web Connector and add that `.qwc` file. Approve access to the company file. Choose the appropriate permission for your business policy; the connector is configured read-only.
5. Enter the same `QBWC_PASSWORD` configured on the VPS when Web Connector prompts for the application password. Save it in Web Connector.
6. Select the PLS Accounts row and click **Update Selected** after queueing a report in PLS.

Keep QuickBooks Desktop and Web Connector running and the company file open during the first test. For scheduled use, configure Web Connector Auto-Run only after a successful manual sync.

## 3. Run a P&L import

1. In PLS Accounts, select the destination company workspace.
2. In the QuickBooks Desktop panel, choose the report start and end dates (normally one calendar month).
3. Click **Queue P&L Sync**.
4. On the Windows PC, select PLS Accounts in Web Connector and click **Update Selected**.
5. When the connector completes, refresh PLS. The imported workbook appears in that workspace. A repeated import for the same QuickBooks period replaces the prior QuickBooks-imported sheet for that period; manually uploaded sheets are not deleted.

## 4. Important deployment notes

- This integration is for one QuickBooks Desktop company file per connected Web Connector identity. For multiple QuickBooks company files, use separate connector identities/configuration and map each to its intended PLS workspace before enabling them.
- The current bridge imports a single-period Profit & Loss summary using the QuickBooks report engine. It does not yet reproduce every possible custom column, class/location split, or custom report layout from a hand-built Excel export.
- The PLS dashboard's existing USD rate workflow remains unchanged; set the rate on the imported sheet if USD conversion is required.
- Keep `QBWC_PASSWORD` private and rotate it by updating the VPS environment and re-entering the new value in Web Connector.
- Back up `data/` before deployment. Do not commit `.env`, passwords, or real company data to Git.

## 5. QBWC error notes

- `QBWC1053`: close modal/pop-up windows in QuickBooks (including Company Settings/Employee Center), then retry. If needed, close QuickBooks and Web Connector and end remaining QuickBooks processes before reopening.
- `QBWC1039`: the exact message matters. Common causes include wrong company-file path, missing authorization, app not first added while signed in as QuickBooks Admin, or a duplicate OwnerID/FileID stamp. Do not repeatedly add the QWC file. If the exact message is “Unique OwnerID/FileID pair value required”, follow Intuit's CP3 Remove Stamp procedure for the old QWC identity, then add the new file once.
- Keep the company file and Web Connector installed on the same Windows PC/local environment. The VPS is only the SOAP web service endpoint.

Official Intuit references:
- https://quickbooks.intuit.com/learn-support/en-us/help-article/open-programs/fix-common-web-connector-errors-quickbooks-desktop/L9slolvg0_US_en_US
- https://quickbooks.intuit.com/learn-support/en-us/help-article/multi-user-mode/fix-error-qbwc1039/L2WqXx5LW_US_en_US
- https://static.developer.intuit.com/resources/QBWC_proguide.pdf
