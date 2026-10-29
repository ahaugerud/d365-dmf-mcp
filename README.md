# d365-dmf-mcp

MCP server that imports CSV files into Dynamics 365 F&O via Data Management (DMF). It runs locally and authenticates as the user signed in to the Azure CLI.

## Setup on Windows

1. Install [uv](https://docs.astral.sh/uv/) and the [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli-windows).

   ```powershell
   winget install astral-sh.uv Microsoft.AzureCLI
   ```

2. Sign in with the account that has access to D365 (if there is no `az` session, the server falls back to an interactive browser login):

   ```powershell
   az login --tenant <entra-tenant-id>
   ```

3. Register the server with Claude Code (runs the published package via `uvx`; see `.mcp.json.example`):

   ```powershell
   claude mcp add d365-dmf `
     -e DMF_AZURE_TENANT_ID=<entra-tenant-id> `
     -e DMF_BASE_URL=https://<env>.sandbox.operations.eu.dynamics.com `
     -- uvx d365-dmf-mcp
   ```

   Other MCP clients: copy the block from `.mcp.json.example` into the client's MCP config. To run from a local checkout instead, use `uvx --from C:\path\to\d365-dmf-mcp d365-dmf-mcp`.

## Tools

Only the tenant and environment URL are config. Everything else is passed per call.

- `import_data(definition_group, company, files, format="CSV")`: import one or more files into a data project and wait for the result. `files` is a list of `{entity, file_path}`; `entity` is the DMF entity name as shown in D365 (e.g. `Customer groups`) and the file (`.csv` or `.xlsx`) has the entity's field names (e.g. `CUSTOMERGROUPID`) in its first row. The server generates the package manifest, looks up each entity's target and the dependency order from D365, and creates the data project if it doesn't exist. `format` is free text (`CSV`, `EXCEL`, or any CSV/Excel-based D365 source format on the environment, e.g. `CSV-Unicode`).
- `export_data(output_dir, definition_group, company)`: run an existing export data project and save its files.
- `create_data_project(name, description="")`: create an empty import data project (optional; `import_data` does this itself).
- `list_data_projects(name_filter="", operation_type="")`: browse existing data projects.
- `list_project_entities(definition_group)`: the entities in a project, with format, execution order and any non-default processing settings.
- `configure_project_entity(definition_group, entity, ...)`: show/change per-entity settings (business validation, field validation, business logic, skip staging, fail-on-error). Defaults are right for almost all imports; only relax validation for large, already-tested migration loads.
- `search_entities(query, limit=50)`: search the DMF entity catalogue by name, target entity or module.
- `get_entity_fields(entity)`: the column names (data file header), types, lengths and required flags for an entity.
- `get_execution_status(execution_id)`: check the status of an execution, with errors if it failed.

## Notes

- Import threshold record count, import task count (parallelism), set-based processing and change tracking have no API; set them in D365 (Data management > Framework parameters > Entity settings > Configure entity execution parameters) for large loads.

- Composite and XML-based entities are not supported.
- Optional env vars: `DMF_POLL_INTERVAL_SECONDS` (default 5), `DMF_POLL_TIMEOUT_SECONDS` (default 300).
