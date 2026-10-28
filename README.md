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

3. Clone this repo:

   ```powershell
   git clone git@github.com:ahaugerud/d365-dmf-mcp.git
   ```

4. Register the server with Claude Code, adjusting the path and values (see `.mcp.json.example`):

   ```powershell
   claude mcp add d365-dmf `
     -e DMF_AZURE_TENANT_ID=<entra-tenant-id> `
     -e DMF_BASE_URL=https://<env>.sandbox.operations.eu.dynamics.com `
     -- uv run --script C:\path\to\d365-dmf-mcp\server.py
   ```

   Other MCP clients: copy the `env` block from `.mcp.json.example` into the client's MCP config.

## Tools

Only the tenant and environment URL are config. Data project, company and template are passed per call.

- `export_data(output_dir, definition_group, company)`: run an export data project and save the files, plus the whole package as `<project>-package.zip`.
- `import_file(file_path, definition_group, company, template_package, format="CSV")`: import a .csv or .xlsx file into a data project and wait for the result. `template_package` is a package zip for the same entity, e.g. the one `export_data` saved. `format` is free text (`CSV`, `EXCEL`, or any CSV/Excel-based D365 source format on the environment, e.g. `CSV-Unicode`).
- `create_data_project(name, company, template_package, format="CSV")`: create a new import data project for the entity in the template package.
- `get_import_status(execution_id)`: check the status of an execution, with errors if it failed.

## Notes

- To get a template for a new entity, create an export data project for it in D365 and run `export_data` on it.
- The data project's source data format must match `format`. Use `create_data_project` with the same format to get a matching project.
- Optional env vars: `DMF_POLL_INTERVAL_SECONDS` (default 5), `DMF_POLL_TIMEOUT_SECONDS` (default 300).
