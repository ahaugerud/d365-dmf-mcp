# d365-dmf-mcp

MCP server that imports CSV files into Dynamics 365 F&O via Data Management (DMF). It runs locally and authenticates as the user signed in to the Azure CLI.

## Setup on Windows

1. Install [uv](https://docs.astral.sh/uv/) and the [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli-windows).

   ```powershell
   winget install astral-sh.uv Microsoft.AzureCLI
   ```

2. Sign in with the account that has access to D365:

   ```powershell
   az login
   ```

3. Clone this repo:

   ```powershell
   git clone git@github.com:ahaugerud/d365-dmf-mcp.git
   ```

4. Register the server with Claude Code, adjusting the path and values (see `.mcp.json.example`):

   ```powershell
   claude mcp add d365-dmf `
     -e DMF_BASE_URL=https://<env>.sandbox.operations.eu.dynamics.com `
     -e DMF_DEFINITION_GROUP=POC-CustomerGroups-Import `
     -e DMF_EXPORT_GROUP_NAME=POC-CustomerGroups-Export `
     -e DMF_COMPANY=USMF `
     -e DMF_TEMPLATE_ZIP=template/CustomerGroups-export.zip `
     -- uv run --script C:\path\to\d365-dmf-mcp\server.py
   ```

   Other MCP clients: copy the `env` block from `.mcp.json.example` into the client's MCP config.

## Tools

- `import_file`: import a .csv or .xlsx file into a data project; `format` is free text (`CSV` default, `EXCEL`, or any CSV/Excel-based D365 source format on the environment, e.g. `CSV-Unicode`) and wait for the result.
- `export_data`: run an export data project and save the files locally.
- `create_data_project`: create a new import data project from the template package.
- `get_import_status`: check the status of an execution.

## Notes

- The CSV columns must match the entity in the template package (`template/`). To use another entity, export a package from a D365 data project and point `DMF_TEMPLATE_ZIP` and `DMF_EXPORT_GROUP_NAME` at it.
- The data project's source data format must match `format`. Use `create_data_project` with the same format to get a matching project.
- Optional env vars: `DMF_POLL_INTERVAL_SECONDS` (default 5), `DMF_POLL_TIMEOUT_SECONDS` (default 300).
