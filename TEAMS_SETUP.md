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

### Reopening completed tasks

Owners, Managers, and Contributors can select **Reopen task** beside a completed
task's status or in its details dialog. This changes only its status to
**In progress** and restores the editing and commenting controls permitted by the
user's role. Its ID, comments, owner, due date, and follow-up tasks are preserved.
Viewers cannot reopen tasks. Under a **Done** filter, a reopened task leaves the
list; select **In progress** or **All tasks** to find it.

### Short descriptions and collapsible details

The short description stays visible beneath each task title. Tasks can also have
a separate detailed description, opened and closed with **Show details** /
**Hide details** in the task list and task dialog. Details start collapsed;
expanding or collapsing them never changes the saved text.

Owners and Managers can edit both descriptions in the task dialog and select
**Save changes**. Contributors can provide both when creating tasks but cannot
edit existing descriptions; Viewers can read and expand details. Completed tasks
remain read-only until reopened. Existing tasks without details keep their
original summaries. Search includes detailed text, and Word reports include both
descriptions.

For engineering review tasks, details can include **Problem**, **Impact**,
**Recommendation**, source references, and proposed closure evidence. Generated
review claims must remain labeled unverified until checked against source
requirements; example limits and standards are not automatically approved.

### Word reports

**Download report** creates an editable landscape Word document styled like the
web workspace: forest-green branding, progress metrics, topic headings, and
task/owner/due-date/status columns with status colors. Full-width sections keep
long text readable instead of squeezing descriptions into narrow table columns.

Reports include the complete active project, regardless of search, status/topic
filters, or collapsed details: task titles and IDs, topic paths, parent and
follow-up references, short and detailed descriptions, owners, due dates,
statuses, and all comments. Detailed text is always expanded, including source
references and evidence cautions. Task references in comments and follow-up
lists link to their entries in the document. Empty fields are labeled explicitly.
Long sections can flow across pages, with repeated column headings and page
numbers. Recycle-bin tasks and interactive web controls are not exported.

### Task IDs and comment links

Every task, including follow-ups, displays a **Task ID** such as `TASK-001` or
`TASK-002` beside its title, in its details dialog, and in Word reports. Numbering
is sequential within each project and continues beyond `TASK-999`. IDs are
searchable and remain unchanged when tasks are edited, moved, reopened, or
restored. Deleted task numbers are not reused.

Existing tasks and recycle-bin entries receive saved numbers automatically when
their project is first accessed. These display IDs are separate from the
unchanged internal IDs, preserving existing comment links and parent relationships.

Use **Link a task** beside a comment composer to select another task in the same
project by ID or title. This inserts `[Task title](task:task-id)` into the draft.
Add the comment as usual; in the details dialog, also select **Save changes**.
Saved references become clickable task titles that open the linked task, even
when its row is filtered out. Links resolve by ID, so renaming a task does not
break them. Removed tasks are labeled unavailable until restored. Plain text and
existing comments remain unchanged; links are not interpreted as HTML.

### Task follow-ups

Use **Add follow-up** below a task or in its details dialog to create an action
linked to that task, including a completed task. Follow-ups appear directly below
their parent (unless a search or status filter hides it) and identify the parent
by title. Each has an independent owner, due date, status, and comments; completing
a parent does not complete its follow-ups. Follow-ups can also have follow-ups.
Each parent shows a **Follow-ups** list of clickable action titles and their
statuses, both beneath the task and in its details dialog. Click a title to open
that follow-up, even if a search or status filter hides its row.
Owners, Managers, and Contributors can create them; Viewers remain read-only.

Follow-ups inherit their parent's topic. Move the parent to move its active
follow-ups together. Removing a parent does not remove its follow-ups; they remain
visible with a removed-parent label until it is restored. Restore a parent to
the same topic as its remaining follow-ups. Word reports include parent titles.
Existing tasks require no data migration. The API accepts an optional
`parentTaskId` when creating a task; this link cannot be reassigned later.

Run the focused API regression tests using `node --test server\follow-ups.test.mjs`.
They use a temporary database and do not modify local project data.

The app now lives in `LoroLite managment\Action Item managment`, with its existing
Git repository and SQLite data preserved. From that directory, run `run-app.bat`
on Windows or `npm run dev` in a terminal. The Windows launcher waits for both
the frontend and its proxied local API before opening the browser. Its HTTP
check uses basic parsing to avoid Windows PowerShell's interactive web-content
prompt and IPv4 loopback to avoid `localhost` resolution delays. If startup fails,
check the development server window and ensure ports
5173 and 3001 are not occupied by another instance.

The project-management agent is available at the LoroLite workspace level in
`..\.github\agents\project-management-expert.agent.md`; its app-operation skill is
in `..\.github\skills\action-item-management\SKILL.md`.


In development, use the **Preview role** selector in the sidebar to test Owner, Manager, Contributor, and Viewer behavior. This is a local-only preview backed by the local SQLite database. Production identity and authorization come from Teams SSO and project memberships.