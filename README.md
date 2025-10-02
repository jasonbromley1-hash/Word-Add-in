# Word-Add-in

Deployment checklist — Estate Clause Helper add-in

This file walks through a recommended production deployment for your Word add-in.

Recommendation (best for a law firm)
- Host the taskpane and static assets on a stable HTTPS host (Azure Static Web Apps recommended).
- Register an Azure AD app for SSO and add the client id into the add-in manifest's WebApplicationInfo.
- Deploy the manifest to Microsoft 365 via Central Deployment (Admin Center) or SharePoint App Catalog.

Quick steps (high level)
1. Build and host the front-end
   - Produce the static build artifact (taskpane.html, function-file.html, assets) under a folder (e.g., `dist/` or `host/`).
   - Deploy to Azure Static Web Apps (recommended) or any HTTPS host. Note: Central deployment requires the SourceLocation to be public HTTPS so Microsoft can fetch it.

2. Register an Azure AD application (for SSO)
   - Azure Portal -> Azure Active Directory -> App registrations -> New registration.
   - Name: Estate Clause Helper
   - Redirect URI: (optional) for web/OIDC flows; your add-in may not need it for basic SSO but note any server-side endpoints.
   - Expose an API or configure scopes if you have back-end APIs.
   - Copy Application (client) ID and paste into `word-addin/manifest.xml` WebApplicationInfo -> Id and Resource fields (use the api://{client-id} resource format if applicable).

3. Update the manifest
   - Replace `https://addins.myfirm.com` with your real host (or ngrok URL for dev) in `word-addin/manifest.xml`.
   - Bump Version if updating an existing deployment.

4. Create ZIP for upload
   - Zip the manifest (`manifest.xml`) and upload the zip to Microsoft 365 Admin Center (Settings -> Integrated apps -> Add an app -> Upload manifest).

5. Verify in Word
   - Assigned users: Insert -> My Add-ins -> My Organization or the add-in should appear on the ribbon if central deployment targeted the ribbon.

Notes and tips
- ngrok is only for short-term development; ngrok URLs rotate unless you have a paid plan with a fixed subdomain.
- If you host internally (no public endpoint), use the SharePoint App Catalog (Central deployment requires Microsoft endpoints to access your SourceLocation so it must be public).
- Ensure your static host serves correct MIME types and that the `taskpane.html` is reachable and serves over HTTPS.
- For SSO, follow Microsoft guidance for Office Add-ins SSO: https://learn.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins

Custom domain / DNS  
- Note the public endpoint of your Azure Static Web App, e.g. `myfirmaddins.z13.web.core.windows.net`.  
- At your DNS provider, create a **CNAME** record:  
  - **Name / Host**: `addins.myfirm.com`  
  - **Type**: CNAME  
  - **Value / Target**: `myfirmaddins.z13.web.core.windows.net`  
- To enable HTTPS on the custom domain, configure **Azure Front Door** or the **Static Web Apps** custom-domain feature and turn on TLS.  
- While DNS propagates (or if you skip the custom domain), end-users can still reach the add-in via the direct `*.web.core.windows.net` URL.

Automation (CI)
- The included GitHub Actions workflow (`.github/workflows/deploy_static_web_app.yml`) is a template — configure `AZURE_STATIC_WEB_APPS_API_TOKEN` as a repository secret and it will deploy the folder under `word-addin/host` to an Azure Static Web App.

If you want, I can:
- Deploy the add-in to Azure Static Web Apps for you (you will need to provide Azure credentials or run the az/gh commands locally).
- Fill in the manifest with your production hostname and Azure AD client id and repackage the manifest zip.
- Walk you through the Admin Center upload and verify the add-in as a test user.

---
Prepared to help with whichever step you want automated next.
