# D365 SpeedUp

> Accelerate Dynamics 365 customization and development with smart tools and generators.

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Manifest](https://img.shields.io/badge/manifest-v3-green)
![Platform](https://img.shields.io/badge/platform-Chrome-yellow)

A Chrome extension for **Dynamics 365 / Dataverse** developers and administrators. Run queries, inspect metadata, generate code, check security, and troubleshoot - all from a compact popup or side panel, directly in the context of your active D365 session.

---

## Features

### 🗄️ Data & Quality
Tools for querying, validating, and managing data.

| Tool | Description |
|------|-------------|
| **FetchXML Tester** | Run any FetchXML query and view results in a sortable, searchable, filterable grid. Supports paginated results and opens records directly from the table. |
| **Get Table Data (Field Values)** | Explore all attributes for a table, including metadata (type, lengths, option sets) and optionally the display/raw values for a specific record GUID. |

### 🔁 Automation & Processes
Tools for monitoring and managing workflows, flows, and event handlers.

| Tool | Description |
|------|-------------|
| **Inspect Table Processes & JS Handlers** | Lists all Classic Workflows, Business Rules, Cloud Flows, Form JavaScript Handlers, Plugin Steps, and Service Endpoints related to a specific table - across five interactive grids. |

### 💻 Development Tools
Tools for generating C# and JavaScript code from Dataverse metadata.

| Tool | Description |
|------|-------------|
| **Generate Early-Bound Wrapper (C#)** | Generates a C# early-bound model class from a Dataverse table, grouped by type (strings, lookups, dates, options, etc.), with enums for PickList/Status/State attributes. |
| **List Table Attributes** | Lists all columns for a table with type, required level, primary flags, valid for read/create/update, and extra details (max length, targets, option set preview). |
| **Generate Relationship Diagram (Mermaid)** | Generates a Mermaid ER diagram for all relationships in a given solution. Optionally filters to a specific table. Copies the diagram to clipboard automatically. |

### 🛡️ Security & Access
Quick tools to check users, roles, teams, and privileges.

| Tool | Description |
|------|-------------|
| **User Details (Roles, Teams etc)** | Shows a user's key fields, settings, directly-assigned roles, team memberships, roles inherited via teams, and field security profiles across six grids. |
| **Privilege Checker (User + Table)** | Shows the effective privilege depth (None to Global) for all standard operations (Read, Write, Create, Delete, etc.) on a table for a given user, and which roles grant each privilege. |

### 🐞 Troubleshooting
Tools for investigating errors, trace logs, and flow health.

| Tool | Description |
|------|-------------|
| **Plugin Trace Logs** | Retrieves plugin trace logs for a specific table within a chosen time window (1h / 24h / 7d / 30d), with execution time, duration, and direct links to log records. |
| **Flows Health Check** | Summarises recent flow runs per Power Automate flow - total, succeeded, failed counts, last run status, last failure time, and error messages - in a single interactive grid. |

---

## Requirements

- **Browser:** Google Chrome (Manifest V3)
- **Environment:** Dynamics 365 CE / Dataverse model-driven app must be open in the active tab
- The extension executes scripts in the context of the D365 page and uses the `Xrm` API - it will not work on non-D365 pages

---


## Usage

1. Navigate to a Dynamics 365 / Dataverse model-driven app page in Chrome
2. Click the **D365 SpeedUp** icon in the toolbar to open the popup
3. Use the sidebar to browse categories and select a tool
4. Fill in the required inputs (many auto-populate from the current page context - table, record ID, user email)
5. Click **Run** to execute
6. Results appear inline as interactive tables or copyable code

**Tip:** Click the side panel icon in the top bar to switch to a persistent side panel view, keeping the tool open while you navigate D365.

---

## Tech Stack

- **TypeScript** - all tool logic in `src/`
- **Chrome Extension Manifest V3** - service worker, scripting API, side panel
- **Prism.js** - syntax highlighting for generated C# and Mermaid output
- **Dataverse REST API v9.2** - all data fetched via OData/FetchXML endpoints

---

## Assets & Credits

Extension icons (`assets/icon16.png`, `assets/icon48.png`, `assets/icon128.png`) were **AI-generated using ChatGPT** (DALL-E).

---

## Contributing

Issues and pull requests are welcome. Please open an issue first to discuss significant changes.

---

## License

MIT - see [LICENSE](LICENSE) for details.
