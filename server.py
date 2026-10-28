# /// script
# requires-python = ">=3.10"
# dependencies = ["requests", "mcp<2"]
# ///
"""MVP MCP server: upload a CSV to a D365 F&O Data Management (DMF) import project.

Auth is the signed-in user's `az login`. Config comes from config.json
(or the file named by $DMF_MCP_CONFIG); relative paths resolve against it.
"""
import json, os, subprocess, time, uuid, zipfile
from datetime import datetime
from pathlib import Path

import requests
from mcp.server.fastmcp import FastMCP

CONFIG_PATH = Path(os.environ.get("DMF_MCP_CONFIG") or Path(__file__).with_name("config.json")).resolve()
CFG = json.loads(CONFIG_PATH.read_text())
BASE = CFG["base_url"].rstrip("/")
ACTIONS = BASE + "/data/DataManagementDefinitionGroups/Microsoft.Dynamics.DataEntities."
DONE = ("Succeeded", "PartiallySucceeded", "Failed", "Canceled")

mcp = FastMCP("d365-dmf")


def _path(p: str) -> Path:
    p = Path(p).expanduser()
    return p if p.is_absolute() else CONFIG_PATH.parent / p


def _session() -> requests.Session:
    token = subprocess.run(
        ["az", "account", "get-access-token", "--resource", BASE, "--query", "accessToken", "-o", "tsv"],
        capture_output=True, text=True, check=True).stdout.strip()
    s = requests.Session()
    s.headers["Authorization"] = f"Bearer {token}"
    return s


def _call(s, action, body):
    r = s.post(ACTIONS + action, json=body)
    if not r.ok:
        raise RuntimeError(f"{action} failed: {r.status_code} {r.text}")
    return r.json()["value"]


def _crlf(b: bytes) -> bytes:
    return b.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")


def _build_package(csv: Path, group: str, zip_path: Path):
    """Reuse manifest/header from an exported package, swap in our CSV."""
    with zipfile.ZipFile(_path(CFG["template_zip"])) as zin, zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename.lower().endswith(".csv"):
                zout.writestr(item.filename, _crlf(csv.read_bytes()))  # DMF needs CRLF
            elif item.filename == "Manifest.xml":  # UTF-16; point at our import project
                xml = zin.read(item.filename).decode("utf-16").replace(CFG["export_group_name"], group)
                zout.writestr(item, xml.encode("utf-16"))
            else:
                zout.writestr(item, zin.read(item.filename))


def _import_package(s, zip_path: Path, name: str, group: str, company: str, execute: bool) -> str:
    blob = json.loads(_call(s, "GetAzureWriteUrl", {"uniqueFileName": name}))["BlobUrl"]
    requests.put(blob, data=zip_path.read_bytes(), headers={"x-ms-blob-type": "BlockBlob"}).raise_for_status()
    zip_path.unlink(missing_ok=True)
    exec_id = f"mcp-{uuid.uuid4().hex[:12]}"
    _call(s, "ImportFromPackage", dict(packageUrl=blob, definitionGroupId=group, executionId=exec_id,
                                       execute=execute, overwrite=True, legalEntityId=company))
    return exec_id


@mcp.tool()
def import_csv(csv_path: str, definition_group: str | None = None, company: str | None = None,
               wait: bool = True, dry_run: bool = False) -> str:
    """Import a CSV file into D365 through an existing DMF data project, as the signed-in az user.

    The CSV must match the entity of the template package configured in config.json.
    Defaults for definition_group and company come from config.json.
    With wait=True, polls until the import finishes (or the configured timeout); otherwise
    returns the execution id to check with get_import_status. dry_run only builds the package.
    """
    group = definition_group or CFG["definition_group"]
    company = company or CFG["company"]
    csv = _path(csv_path)
    if not csv.is_file():
        raise FileNotFoundError(csv)
    name = f"{group}-{datetime.now():%Y%m%d%H%M%S}"
    zip_path = Path(os.environ.get("TMPDIR", "/tmp")) / f"{name}.zip"
    _build_package(csv, group, zip_path)
    if dry_run:
        return f"Built {zip_path}"

    s = _session()
    exec_id = _import_package(s, zip_path, name, group, company, execute=True)
    if not wait:
        return f"Triggered {exec_id}"
    deadline = time.time() + CFG.get("poll_timeout_seconds", 300)
    while (st := _call(s, "GetExecutionSummaryStatus", {"executionId": exec_id})) not in DONE:
        if time.time() > deadline:
            return f"Triggered {exec_id}; still {st} after timeout, use get_import_status"
        time.sleep(CFG.get("poll_interval_seconds", 5))
    return f"{exec_id}: {st}"


@mcp.tool()
def create_data_project(name: str, description: str = "", company: str | None = None) -> str:
    """Create a new DMF *import* data project for the entity of the configured template package.

    Creates the project, then loads the template package into it without executing, which
    defines the entity and field mapping from the package manifest. Fails if the name exists.
    Afterwards use import_csv(csv_path, definition_group=name).
    """
    company = company or CFG["company"]
    s = _session()
    r = s.post(BASE + "/data/DataManagementDefinitionGroups", json={
        "Name": name, "OperationType": "Import", "ProjectCategory": "Integration",
        "Description": description, "GenerateDataPackage": "No"})
    if not r.ok:
        raise RuntimeError(f"Create project failed: {r.status_code} {r.text}")
    zip_path = Path(os.environ.get("TMPDIR", "/tmp")) / f"{name}-setup-{datetime.now():%Y%m%d%H%M%S}.zip"
    with zipfile.ZipFile(_path(CFG["template_zip"])) as zin:
        _build_package(Path(zin.extract(next(n for n in zin.namelist() if n.lower().endswith(".csv")),
                                        zip_path.parent / "dmf-tpl")), name, zip_path)
    _import_package(s, zip_path, zip_path.stem, name, company, execute=False)
    return f"Created data project {name}"


@mcp.tool()
def get_import_status(execution_id: str) -> str:
    """Get the status of a DMF execution (e.g. NotRun, Executing, Succeeded, PartiallySucceeded, Failed, Canceled)."""
    return _call(_session(), "GetExecutionSummaryStatus", {"executionId": execution_id})


if __name__ == "__main__":
    mcp.run()
