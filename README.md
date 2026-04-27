# Outlook LaunchEvents Sample — Add-in + Web API

A developer sample that demonstrates every currently supported Outlook **event-based add-in** (LaunchEvent) trigger. Each event is captured by the add-in and POSTed to a companion ASP.NET Core Web API. A real-time **Event Viewer** page lets you watch the events arrive in the browser without polling.

Other tests are included in the TaskPane, such as Drag/Drop testing, sessionData persistence, and a few other miscellaneous tests.

---

## Table of Contents

- [What the sample does](#what-the-sample-does)
- [Architecture overview](#architecture-overview)
- [Project structure](#project-structure)
- [LaunchEvents handled](#launchevents-handled)
- [API endpoints](#api-endpoints)
- [Add-in task pane features](#add-in-task-pane-features)
- [Event Viewer](#event-viewer)
- [Prerequisites](#prerequisites)
- [Running locally](#running-locally)
- [Deploying to your own server](#deploying-to-your-own-server)
- [Sideloading the manifest](#sideloading-the-manifest)
- [Customising for your environment](#customising-for-your-environment)
- [Add-in settings (roaming)](#add-in-settings-roaming)
- [Developer notes](#developer-notes)

---

## What the sample does

When Outlook raises a supported event (e.g. a new message compose window opens, recipients change, the message is sent), the add-in's event runtime fires a JavaScript handler. That handler:

1. Reads the configured API URL from the add-in's roaming settings.
2. POSTs a plain-text event name (optionally prefixed with the user's display name) to `POST /TestAPI/LogEventDelayed`.
3. Calls `event.completed()` — passing `allowEvent: false` and an error message if the API call failed and **Block on API fail** is enabled.

The Web API logs each event to the console **and** broadcasts it in real time to any connected Event Viewer browser tabs via Server-Sent Events (SSE).

---

## Architecture overview

```
Outlook client (OWA / Desktop)
  │
  │  Office.js LaunchEvent runtime  (app.js — event handlers)
  │  Office.js Task Pane            (taskpane.js / taskpane.html)
  │
  │  POST /TestAPI/LogEventDelayed   (text/plain)
  ▼
ASP.NET Core 8 Web API  (TestAPIController)
  │── logs to console
  │── broadcasts via SSE to connected clients
  ▼
EventViewer.html  (browser tab, any device)
  └── GET /TestAPI/MonitorEvents  →  text/event-stream (SSE)
```

The server is self-hosted — the same process serves both the REST API and all static add-in files from `wwwroot/`.

---

## Project structure

```
WebAPISample/
├── Controllers/
│   └── TestAPIController.cs      # All API endpoints + SSE broadcast infrastructure
├── wwwroot/LaunchEventDemo/
│   ├── app.js                    # Event-based add-in runtime (LaunchEvent handlers)
│   ├── taskpane.js / .html / .css # Task pane UI and logic
│   ├── dialog.js / .html / .css  # Office dialog used from task pane
│   ├── EventViewer.html          # Real-time SSE event monitor page
│   ├── common/
│   │   ├── constants.js          # Shared constants (ES module)
│   │   └── constants.mjs         # Same constants for non-module contexts
│   ├── assets/                   # Add-in icons
│   ├── LaunchEvent Demo Manifest.xml          # Standard manifest
│   ├── LaunchEvent Demo Manifest (mobile).xml # Mobile manifest variant
│   └── daves.tips LaunchEvent Demo (all).xml  # Extended manifest (all events)
├── Program.cs                    # ASP.NET Core host configuration, CORS, static files
├── TextPlainInputFormatter.cs    # Allows controllers to receive text/plain bodies
├── appsettings.json              # Kestrel endpoint and certificate configuration
└── Properties/launchSettings.json
```

---

## LaunchEvents handled

| Event | Handler function | Notes |
|---|---|---|
| `OnNewMessageCompose` | `OnNewMessageComposeHandler` | Fires when a new compose window opens |
| `OnNewAppointmentOrganizer` | `OnNewAppointmentOrganizerHandler` | New appointment compose |
| `OnMessageCompose` | `OnMessageComposeHandler` | Message compose open/resume |
| `OnAppointmentOrganizer` | `OnAppointmentOrganizerHandler` | Appointment compose open/resume |
| `OnMessageAttachmentsChanged` | `OnMessageAttachmentsChangedHandler` | Attachment added/removed |
| `OnAppointmentAttachmentsChanged` | `OnAppointmentAttachmentsChangedHandler` | |
| `OnMessageRecipientsChanged` | `OnMessageRecipientsChangedHandler` | To/CC/BCC changed |
| `OnAppointmentAttendeesChanged` | `OnAppointmentAttendeesChangedHandler` | |
| `OnAppointmentTimeChanged` | `OnAppointmentTimeChangedHandler` | Start/end time changed |
| `OnAppointmentRecurrenceChanged` | `OnAppointmentRecurrenceChangedHandler` | Recurrence pattern changed |
| `OnInfoBarDismissClicked` | `OnInfoBarDismissClickedHandler` | User dismisses a notification bar |
| `OnMessageSend` | `onMessageSendHandler` | Send mode: `SoftBlock` |
| `OnAppointmentSend` | `OnAppointmentSendHandler` | Send mode: `SoftBlock` |
| `OnMessageFromChanged` | `OnMessageFromChangedHandler` | From address changed |
| `OnAppointmentFromChanged` | `OnAppointmentFromChangedHandler` | |
| `OnSensitivityLabelChanged` | `OnSensitivityLabelChangedHandler` | Sensitivity label changed |

> **Send events** (`OnMessageSend`, `OnAppointmentSend`) use `SoftBlock` — the send is blocked if the API call fails *and* the **Block on API fail** setting is enabled. The custom smart alert dialog feature (optional) can present a formatted Markdown error with a button that re-opens the task pane.

---

## API endpoints

All endpoints are under the route pattern `[controller]/[action]`.

| Method | Path | Description |
|---|---|---|
| `POST` | `/TestAPI/LogEvent` | Accepts `text/plain`, logs the body instantly, broadcasts via SSE. |
| `POST` | `/TestAPI/LogEventDelayed` | As above, but accepts an optional `?DelayInSeconds=N` query parameter to simulate a slow API. The event is broadcast at **receive time**, before the delay. |
| `GET` | `/TestAPI/MonitorEvents` | Opens a persistent SSE stream (`text/event-stream`). All subsequent `LogEvent`/`LogEventDelayed` calls are pushed to every connected client. |
| `GET` | `/TestAPI/GetRandomNumberAfterDelay` | Test endpoint — returns a random `Int64` after `?ReplyDelay=N` seconds. |
| `POST` | `/TestAPI/ReturnTextAfterDelay` | Test endpoint — echoes the JSON body after `?SecondsToWait=N` seconds. |

### SSE message format

Each SSE message carries a JSON `data` payload:

```json
{ "type": "event", "timestamp": "2024-06-01T10:23:45.123+00:00", "message": "LaunchEventDemo: OnMessageSend" }
```

The initial connection message uses `"type": "connected"`.

---

## Add-in task pane features

The task pane (`taskpane.html` / `taskpane.js`) provides:

- **Message information** — displays the subject of the currently selected item.
- **Tests section** — buttons to exercise individual Office.js API capabilities:
  - Set extended (MAPI) properties on the current item
  - Open an external link via Office dialog
  - Open an Office.js dialog (`dialog.html`)
  - Apply an `InsightMessage` notification
  - Retrieve full message details
  - Send a message (compose mode only)
  - Create a new appointment
  - Set a `sessionData` flag (survives event handler calls within the same compose session)
- **Drag-and-drop tests** — two drop targets demonstrating HTML5 and Office.js drag-and-drop APIs.
- **Copy body to clipboard** — copies the message body as plain text or HTML.
- **Add-in configuration** — all settings are persisted in Office roaming settings and are automatically applied to the event runtime:

| Setting | Description |
|---|---|
| API Endpoint | Full URL of the logging endpoint (default: `{origin}/TestAPI/LogEventDelayed`) |
| API Delay | Appended as `?DelayInSeconds=N` to simulate slow API responses |
| Clientside Delay | Adds a `setTimeout` before the API call fires |
| Test Recipient | Email address used by the "Send Message" / "Create Appointment" test buttons |
| Send user details | Prepends the user's display name to the event string |
| Block on API fail | Sets `allowEvent: false` if the POST returns a non-200 status |
| Log appointment ID | Forces a save to obtain an item ID before logging |
| Add events to notification bar | Appends each event name to the item's notification bar |
| Show custom smart alert dialog | Shows a formatted Markdown dialog on send events instead of logging |

- **Open Event Viewer** — a link in the page header opens `EventViewer.html` in a new tab.
- **Debug console** — task pane overrides `console.log` and renders output directly in the page for quick in-Outlook debugging.

---

## Event Viewer

`EventViewer.html` is a standalone browser page that monitors events in real time.

- **Start** — opens an `EventSource` connection to `GET /TestAPI/MonitorEvents`. The connection is kept alive by the server.
- **Stop** — closes the connection.
- **Clear** — resets the event log.

Each received event is displayed as a timestamped row. System messages (connected, stopped) are displayed in a distinct colour. The page does not require Office.js and can be opened in any browser that can reach the server.

---

## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8)
- A valid **HTTPS** certificate accessible to the server (required by both Outlook and the Office.js runtime — HTTP is not accepted)
- A Microsoft 365 account with permission to sideload add-ins

---

## Running locally

1. Clone the repository:
   ```powershell
   git clone https://github.com/David-Barrett-MS/Outlook-LaunchEvents-Sample
   cd Outlook-LaunchEvents-Sample
   ```

2. Trust the ASP.NET Core development certificate (first time only):
   ```powershell
   dotnet dev-certs https --trust
   ```

3. Run the project:
   ```powershell
   dotnet run --launch-profile https
   ```
   The API and static files will be served at `https://localhost:7140`.

4. Verify the add-in files are accessible:
   ```
   https://localhost:7140/LaunchEventDemo/index.html
   https://localhost:7140/LaunchEventDemo/taskpane.html
   ```

> **Note:** For local development the manifest must reference `https://localhost:7140`. See [Customising for your environment](#customising-for-your-environment) below.

---

## Deploying to your own server

The application is designed to run as a self-contained HTTPS executable, making it straightforward to host on any internet-facing Windows or Linux server.

### Publish a self-contained executable

```powershell
dotnet publish -c Release -r win-x64 --self-contained true -o ./publish
```

Copy the `publish/` folder to your server and run the executable. For Linux, use `-r linux-x64`.

### Configure Kestrel endpoints

Edit `appsettings.json` to point Kestrel at your server's address and certificate:

```json
"Kestrel": {
  "Endpoints": {
    "Https": {
      "Url": "https://0.0.0.0:443"
    }
  },
  "Certificates": {
    "Default": {
      "Subject": "your.domain.com",
      "Store": "webhosting",
      "Location": "LocalMachine",
      "AllowInvalid": false
    }
  }
}
```

Alternatively, supply the certificate via environment variables or a `.pfx` file — see [Kestrel HTTPS configuration](https://learn.microsoft.com/aspnet/core/fundamentals/servers/kestrel/endpoints).

### Configure CORS

`Program.cs` contains a hardcoded CORS policy. Update the `WithOrigins` call to include your actual hosting domain:

```csharp
policy.WithOrigins("https://your.domain.com", "https://outlook.office.com")
```

---

## Sideloading the manifest

### Outlook on the web (OWA)

1. Open [Outlook on the web](https://outlook.office.com).
2. Go to **Settings → Manage add-ins → Add a custom add-in → Add from file**.
3. Upload the updated `LaunchEvent Demo Manifest.xml`.

### Outlook Desktop (Windows)

1. In Outlook, go to **Home → Get Add-ins**.
2. Select **My add-ins → Add a custom add-in → Add from file**.
3. Upload the manifest XML.

### Microsoft 365 Admin Center (organisation-wide)

1. Go to [admin.microsoft.com](https://admin.microsoft.com) → **Settings → Integrated apps**.
2. Upload the manifest for all users or a specific group.

---

## Customising for your environment

All URLs in the manifest use the placeholder `https://~remoteappurl`. Replace every occurrence with your actual hosting URL before sideloading.

### Files to update

| File | What to change |
|---|---|
| `LaunchEvent Demo Manifest.xml` | Replace all `https://~remoteappurl` with your URL (9 occurrences) |
| `LaunchEvent Demo Manifest (mobile).xml` | Same replacement for the mobile manifest |
| `daves.tips LaunchEvent Demo (all).xml` | Same for the extended manifest |
| `common/constants.js` and `common/constants.mjs` | Update `externalLink` to point to your hosted `external.html` |
| `Program.cs` | Update `WithOrigins(...)` in the CORS policy |
| `appsettings.json` | Update Kestrel endpoints and certificate details |

### Quick find-and-replace (PowerShell)

```powershell
$old = "https://~remoteappurl"
$new = "https://your.domain.com"
Get-ChildItem -Path .\wwwroot\LaunchEventDemo -Filter "*.xml" | ForEach-Object {
    (Get-Content $_.FullName) -replace [regex]::Escape($old), $new | Set-Content $_.FullName
}
```

### Default API URL

When the task pane loads for the first time it sets the API URL to `{window.location.origin}/TestAPI/LogEventDelayed`. This means the default URL automatically resolves to the hosting domain — no manual configuration is needed for users as long as the add-in and API are on the same origin.

---

## Add-in settings (roaming)

Settings are stored using `Office.context.roamingSettings` and roam with the user's mailbox, so they apply consistently across Outlook clients.

| Setting key | Type | Default | Description |
|---|---|---|---|
| `apiUrl` | string | `{origin}/TestAPI/LogEventDelayed` | API endpoint for event logging |
| `apiDelay` | number | `0` | Seconds of server-side delay |
| `clientDelay` | number | `0` | Seconds of client-side delay before the XHR fires |
| `testRecipient` | string | `""` | Recipient for test send/appointment operations |
| `sendClientInfo` | bool | `false` | Prepend display name to event data |
| `blockOnAPIFail` | bool | `false` | Block send events if the API returns an error |
| `obtainAppointmentId` | bool | `false` | Force-save new items to obtain an item ID |
| `showEventsOnMessage` | bool | `false` | Show events on the item notification bar |
| `showCustomSmartAlertDialog` | bool | `false` | Show a custom smart alert on send events |

---

## Developer notes

### Two logging paths in `app.js`

- **`logEvent()`** — fire-and-forget. Always calls `event.completed({ allowEvent: true })` immediately after sending the XHR. Use for events that must not block Outlook.
- **`logEvent2()`** — waits for the XHR response. Calls `event.completed({ allowEvent: false })` if the API returns an error and **Block on API fail** is enabled. Also reads `sessionData` for send events before dispatching. Use for send-event handlers where you need to inspect the API response.

### Runtime vs. task pane

- **`app.js`** runs in the event-based (headless) runtime. It has no DOM and a strict 5-minute timeout per event. Initialisation via `Office.initialize` is **not** called for `OnMessageSend` in Outlook Desktop — settings must be read lazily inside each handler via `ReadAddinSettings()`.
- **`taskpane.js`** runs in the normal browser-based task pane runtime and has full DOM access. Settings changed in the task pane are saved to roaming settings and picked up by `app.js` on the next event.

### `text/plain` request body support

ASP.NET Core does not natively bind `text/plain` request bodies to `string` parameters. The custom `TextPlainInputFormatter` class registers this support so controllers can use `[FromBody] string`.

### SSE implementation

`TestAPIController` holds a static `ConcurrentDictionary<string, SseClient>` keyed by a per-connection GUID. Each `SseClient` wraps a `ConcurrentQueue<string>` and a `SemaphoreSlim` so the `MonitorEvents` endpoint can `await` new messages without polling. When the browser disconnects, the `CancellationToken` fires and the client entry is removed.

