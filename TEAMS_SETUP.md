# Teams Setup

The app is structured as a shared Teams personal app. Each signed-in user opens the same hosted service and database; project membership and roles decide what they can see or change. Publishing an app to the organization makes it available in the Teams app catalog. A Teams admin can separately pin it for everyone with an app setup policy.

## Deploy the service

1. Build the client with `npm run build` and run the Express service with `npm start` behind a public HTTPS reverse proxy.
2. Keep one service instance and mount persistent storage for `DB_PATH`. The starter backend uses SQLite and is intended for a single instance. For multi-instance/high-volume use, move the project and membership data to a managed database such as Azure Database for PostgreSQL.
3. Configure the production environment from `.env.example`. `TEAMS_DOMAIN` must be the public hostname only, without a scheme or path. Do not use localhost for the published app.

## Configure Microsoft Entra ID

1. Create an app registration in the organization tenant and use its client ID for `ENTRA_CLIENT_ID`.
2. Expose an API scope for the app. Set the Application ID URI to `api://<TEAMS_DOMAIN>/<ENTRA_CLIENT_ID>` (or set `ENTRA_APP_ID_URI` to the exact URI you choose).
3. Enable Teams single sign-on for the registration and authorize the Microsoft Teams desktop and web client applications for the exposed delegated scope.
4. Set `ENTRA_TENANT_ID` to the organization's tenant ID. Production API requests require a valid Teams SSO bearer token from that tenant; the local role-preview header is disabled in production.
5. Add the public hostname to the app registration's redirect/origin configuration as required by the organization's Teams SSO setup.

## Build and publish the Teams package

Set `TEAMS_APP_ID` to a new GUID, `TEAMS_DOMAIN` to the public hostname, `ENTRA_CLIENT_ID` to the app registration's client ID, and optionally `ENTRA_APP_ID_URI` to the API URI. Then run:

```powershell
npm run package:teams
```

This creates `dist/fieldnotes-teams.zip`, containing a generated Teams manifest and the required icons. Upload the package in the Teams admin center, review its policies, and publish it to the organization's app catalog. To pin it for users, assign a Teams app setup policy that installs and pins Fieldnotes. Tenant-wide publishing and tenant-wide pinning are separate admin actions.

The package exposes a personal tab. It is not a Loop component and does not write into Loop pages; the supported first step is opening this shared web app from Teams. A future Loop integration would need separate Microsoft Graph/Loop API validation and tenant permissions.

## Local role preview

In development, use the **Preview role** selector in the sidebar to test Owner, Manager, Contributor, and Viewer behavior. This is a local-only preview backed by the local SQLite database. Production identity and authorization come from Teams SSO and project memberships.