"""Serve the local catalog importer and read stock from QuickBooks POS."""

import json
import math
import mimetypes
import os
import re
import subprocess
import sys
import threading
import base64
import binascii
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
HOST = "127.0.0.1"
PORT = 8765
APP_NAME = "Daychounieh Tools Local Inventory Importer"
APP_ID = "{C18A21C1-6F93-4A86-9294-7B98A1FB8FC0}"
COM_PROGID = "QBPOSXMLRPLib.RequestProcessor"
ITEM_CODE_FIELD = os.environ.get("QBPOS_ITEM_CODE_FIELD", "ItemNumber")
QBPOSXML_VERSION = "3.0"
POS_REQUEST_LOCK = threading.Lock()
PUBLISH_LOCK = threading.Lock()
LIVE_REPOSITORY_URL = "https://github.com/DaychouniehTools/Tools.git"
LIVE_REPOSITORY = Path.home() / "Documents" / "GitHub" / "DaychouniehTools-Tools"
MAX_PUBLISH_BODY_BYTES = 128 * 1024 * 1024
MAX_IMAGE_BYTES = 24 * 1024 * 1024
IMAGE_FILENAME_PATTERN = re.compile(r"[A-Za-z0-9_.-]+\.(?:jpe?g|png|webp|gif|avif)", re.I)
DESCRIPTION_FIELDS = (
    "Desc2",
    "Description",
    "ItemDescription",
    "ItemDesc",
    "SalesDesc",
    "ItemDesc2",
    "ItemDesc1",
)


def query_pos_response():
    import pythoncom
    import win32com.client

    request_xml = (
        '<?xml version="1.0" ?>'
        f'<?qbposxml version="{QBPOSXML_VERSION}"?>'
        '<QBPOSXML><QBPOSXMLMsgsRq onError="stopOnError">'
        "<ItemInventoryQueryRq></ItemInventoryQueryRq>"
        "</QBPOSXMLMsgsRq></QBPOSXML>"
    )
    pythoncom.CoInitialize()
    try:
        processor = win32com.client.Dispatch(COM_PROGID)
        processor.OpenConnection(APP_ID, APP_NAME)
        ticket = None
        try:
            ticket = processor.BeginSession("")
            response_xml = processor.ProcessRequest(ticket, request_xml)
        finally:
            try:
                if ticket:
                    processor.EndSession(ticket)
            finally:
                processor.CloseConnection()
    finally:
        pythoncom.CoUninitialize()

    return response_xml


def query_pos_inventory():
    return parse_pos_response(query_pos_response())


def parse_pos_response(response_xml):
    root = ET.fromstring(response_xml)
    response = root.find(".//ItemInventoryQueryRs")
    if response is not None:
        status_code = response.get("statusCode", "0")
        if status_code not in ("0", "1"):
            raise RuntimeError(
                f"QuickBooks POS error {status_code}: "
                f"{response.get('statusMessage', 'unknown error')}"
            )

    items = {}
    missing_codes = 0
    for item in root.iter("ItemInventoryRet"):
        fields = {}
        for element in item.iter():
            if element is not item and len(element) == 0:
                fields[element.tag] = (element.text or "").strip()

        normalized_fields = {name.casefold(): value for name, value in fields.items()}
        sku = fields.get(ITEM_CODE_FIELD, "").strip()
        if not sku:
            missing_codes += 1
            continue

        store_quantities = [
            _number(value)
            for name, value in fields.items()
            if re.fullmatch(r"OnHandStore\d+", name)
        ]
        store_quantities = [value for value in store_quantities if value is not None]
        if store_quantities:
            quantity = sum(store_quantities)
        else:
            quantity = next(
                (
                    number
                    for name in ("QuantityOnHand", "QtyOnHand", "OnHand", "TotalQtyOnHand")
                    if (number := _number(fields.get(name))) is not None
                ),
                None,
            )
        if quantity is None:
            raise RuntimeError(
                f"Could not find an on-hand quantity for POS item {sku!r}. "
                "Check that the POS item inventory response includes OnHandStore "
                "or QtyOnHand fields."
            )
        if not math.isfinite(quantity):
            raise RuntimeError(f"POS returned a non-finite quantity for item {sku!r}.")
        description = next(
            (
                normalized_fields[name.casefold()]
                for name in DESCRIPTION_FIELDS
                if normalized_fields.get(name.casefold())
            ),
            "",
        )
        web_price = _number(normalized_fields.get("webprice"))
        if web_price is not None and not math.isfinite(web_price):
            raise RuntimeError(f"POS returned a non-finite web price for item {sku!r}.")
        items[sku] = {
            "sku": sku,
            "quantity": quantity,
            "name": normalized_fields.get("desc1", "").strip(),
            "description": description,
            "price": web_price,
            "departmentCode": normalized_fields.get("departmentcode", "").strip(),
            "size": normalized_fields.get("size", "").strip(),
            "attribute": normalized_fields.get("attribute", "").strip(),
            "barcode": normalized_fields.get("upc", "").strip(),
            "keywords": normalized_fields.get("keywords", "").strip(),
        }

    if missing_codes:
        print(f"Skipped {missing_codes} item(s) missing {ITEM_CODE_FIELD}.")
    return {"items": list(items.values()), "missingItemCodes": missing_codes}


def _number(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _run_git(arguments, cwd=None, timeout=240):
    environment = os.environ.copy()
    environment["GIT_TERMINAL_PROMPT"] = "0"
    environment["GCM_INTERACTIVE"] = "auto"
    try:
        result = subprocess.run(
            ["git", *arguments],
            cwd=cwd,
            env=environment,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
        )
    except FileNotFoundError as error:
        raise RuntimeError("Git was not found. Install GitHub Desktop and restart the POS importer.") from error
    except subprocess.TimeoutExpired as error:
        raise RuntimeError(f"Git {' '.join(arguments)} timed out. Check the GitHub connection and try again.") from error
    if result.returncode:
        detail = (result.stderr or result.stdout).strip()
        if arguments and arguments[0] in {"fetch", "push", "clone"}:
            detail = (
                "GitHub sign-in or network access was not available. Confirm GitHub Desktop is signed in "
                "to an account that can write to DaychouniehTools/Tools, then retry."
            )
        raise RuntimeError(f"Git {' '.join(arguments[:2])} failed. {detail}".strip())
    return result.stdout.strip()


def _ensure_live_repository():
    if not LIVE_REPOSITORY.exists():
        LIVE_REPOSITORY.parent.mkdir(parents=True, exist_ok=True)
        _run_git(
            ["clone", "--single-branch", "--branch", "main", LIVE_REPOSITORY_URL, str(LIVE_REPOSITORY)],
            timeout=300,
        )
    if not (LIVE_REPOSITORY / ".git").is_dir():
        raise RuntimeError(
            f"The publish folder {LIVE_REPOSITORY} already exists but is not a Git clone. "
            "Rename or move that folder, then retry."
        )

    origin = _run_git(["remote", "get-url", "origin"], cwd=LIVE_REPOSITORY)
    normalized_origin = origin.strip().removesuffix(".git").rstrip("/").casefold()
    if normalized_origin not in {
        "https://github.com/daychouniehtools/tools",
        "git@github.com:daychouniehtools/tools",
        "ssh://git@github.com/daychouniehtools/tools",
    }:
        raise RuntimeError(f"Safety stop: {LIVE_REPOSITORY} is not connected to the live DaychouniehTools/Tools repository.")

    branch = _run_git(["branch", "--show-current"], cwd=LIVE_REPOSITORY)
    if branch != "main":
        raise RuntimeError(f"Safety stop: the live publish clone is on branch {branch!r}, not 'main'.")
    if _run_git(["status", "--porcelain"], cwd=LIVE_REPOSITORY):
        raise RuntimeError(
            f"Safety stop: {LIVE_REPOSITORY} has uncommitted changes. Review or commit them in GitHub Desktop before syncing."
        )

    _run_git(["fetch", "origin", "main"], cwd=LIVE_REPOSITORY, timeout=300)
    _run_git(["merge", "--ff-only", "origin/main"], cwd=LIVE_REPOSITORY)


def publish_pos_catalog(payload):
    products = payload.get("products")
    images = payload.get("images", {})
    if not isinstance(products, list) or not products or len(products) > 10000:
        raise ValueError("The reviewed catalog is empty or has an invalid number of products.")
    if not isinstance(images, dict) or len(images) > 1000:
        raise ValueError("The new product picture list is invalid.")

    seen_skus = set()
    referenced_images = set()
    for product in products:
        if not isinstance(product, dict):
            raise ValueError("The catalog contains an invalid product record.")
        sku = str(product.get("sku") or "").strip()
        if not sku or sku.casefold() in seen_skus:
            raise ValueError("The catalog contains a missing or duplicate item code.")
        seen_skus.add(sku.casefold())
        for field in ("imagePath", "additionalImagePaths"):
            paths = product.get(field, [])
            if isinstance(paths, str):
                paths = [paths]
            if not isinstance(paths, list):
                raise ValueError(f"The catalog has an invalid {field} value for item {sku}.")
            for image_path in paths:
                match = re.fullmatch(r"(?:\./)?images/([^/\\]+)", str(image_path or ""), re.I)
                if match and IMAGE_FILENAME_PATTERN.fullmatch(match.group(1)):
                    referenced_images.add(match.group(1))

    decoded_images = {}
    total_image_bytes = 0
    for filename, encoded in images.items():
        if not isinstance(filename, str) or not IMAGE_FILENAME_PATTERN.fullmatch(filename):
            raise ValueError("A new product picture has an invalid filename.")
        if filename not in referenced_images or not isinstance(encoded, str):
            raise ValueError(f"Picture {filename!r} is not referenced by the reviewed catalog.")
        try:
            image_bytes = base64.b64decode(encoded, validate=True)
        except (ValueError, binascii.Error) as error:
            raise ValueError(f"New product picture {filename!r} could not be decoded.") from error
        if not image_bytes or len(image_bytes) > MAX_IMAGE_BYTES:
            raise ValueError(f"New product picture {filename!r} is empty or larger than 24 MB.")
        total_image_bytes += len(image_bytes)
        if total_image_bytes > 96 * 1024 * 1024:
            raise ValueError("The new product pictures exceed the 96 MB publish limit.")
        decoded_images[filename] = image_bytes

    with PUBLISH_LOCK:
        _ensure_live_repository()
        catalog_path = LIVE_REPOSITORY / "catalog.js"
        catalog_text = "window.NABIL_PRODUCTS = " + json.dumps(products, ensure_ascii=False, indent=2) + ";\n"
        catalog_path.write_text(catalog_text, encoding="utf-8", newline="\n")
        staged_paths = ["catalog.js"]
        images_directory = LIVE_REPOSITORY / "images"
        images_directory.mkdir(exist_ok=True)
        for filename, image_bytes in decoded_images.items():
            (images_directory / filename).write_bytes(image_bytes)
            staged_paths.append(f"images/{filename}")

        _run_git(["add", "--", *staged_paths], cwd=LIVE_REPOSITORY)
        if not _run_git(["diff", "--cached", "--name-only"], cwd=LIVE_REPOSITORY):
            return {"changed": False, "message": "The live catalog already matches the reviewed QuickBooks update."}

        _run_git(
            ["-c", "user.name=DaychouniehTools", "-c", "user.email=75885662+DaychouniehTools@users.noreply.github.com",
             "commit", "-m", "Sync catalog from QuickBooks POS"],
            cwd=LIVE_REPOSITORY,
        )
        commit = _run_git(["rev-parse", "--short", "HEAD"], cwd=LIVE_REPOSITORY)
        _run_git(["push", "origin", "main"], cwd=LIVE_REPOSITORY, timeout=300)

    return {
        "changed": True,
        "commit": commit,
        "productCount": len(products),
        "pictureCount": len(decoded_images),
    }


def discover_pos_fields():
    root = ET.fromstring(query_pos_response())
    first_item = next(root.iter("ItemInventoryRet"), None)
    if first_item is None:
        raise RuntimeError("QuickBooks POS returned no items to inspect.")

    fields = {
        element.tag: (element.text or "").strip()
        for element in first_item.iter()
        if element is not first_item and len(element) == 0
    }
    print("XML fields on the first QuickBooks POS inventory item:")
    for name in sorted(fields, key=str.casefold):
        print(f"  {name}")

    candidates = {
        name: value
        for name, value in fields.items()
        if name.casefold() in {
            "alu", "itemnumber", "itemno", "itemcode", "sku", "upc", "listid"
        }
        or re.fullmatch(r"(?:qty.*onhand|onhand.*|qtyonhand)", name, re.I)
    }
    print("\nPossible item-code and stock fields on that item:")
    for name, value in sorted(candidates.items(), key=lambda entry: entry[0].casefold()):
        print(f"  {name} = {value}")


class ImporterHandler(BaseHTTPRequestHandler):
    def _send(self, status, content_type, body):
        encoded = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(encoded)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(encoded)

    def _request_is_local(self):
        if self.headers.get("Host") != f"{HOST}:{PORT}":
            return False
        origin = self.headers.get("Origin")
        return not origin or origin == f"http://{HOST}:{PORT}"

    def do_GET(self):
        if not self._request_is_local():
            self._send(403, "text/plain; charset=utf-8", "Local access only.")
            return

        route = urlsplit(self.path).path
        if route == "/api/pos/inventory":
            try:
                with POS_REQUEST_LOCK:
                    inventory = query_pos_inventory()
            except Exception as error:
                message = json.dumps({"error": str(error)})
                self._send(502, "application/json; charset=utf-8", message)
                return
            payload = json.dumps(
                {
                    **inventory,
                    "loadedAt": datetime.now(timezone.utc).isoformat(),
                }
            )
            self._send(200, "application/json; charset=utf-8", payload)
            return

        local_files = {
            "/": ("admin-import.html", "text/html; charset=utf-8"),
            "/admin-import.html": ("admin-import.html", "text/html; charset=utf-8"),
            "/index.html": ("index.html", "text/html; charset=utf-8"),
            "/catalog.js": ("catalog.js", "text/javascript; charset=utf-8"),
        }
        if route.startswith("/images/"):
            image_name = route[len("/images/") :]
            if not re.fullmatch(r"[A-Za-z0-9_.-]+\.(?:jpe?g|png|webp|gif|avif)", image_name, re.I):
                self._send(404, "text/plain; charset=utf-8", "Not found.")
                return
            local_files[route] = (
                f"images/{image_name}",
                mimetypes.guess_type(image_name)[0] or "application/octet-stream",
            )

        file_entry = local_files.get(route)
        if not file_entry:
            self._send(404, "text/plain; charset=utf-8", "Not found.")
            return
        relative_path, content_type = file_entry
        file_path = (ROOT / relative_path).resolve()
        if ROOT not in file_path.parents or not file_path.is_file():
            self._send(404, "text/plain; charset=utf-8", "Not found.")
            return
        self._send(200, content_type, file_path.read_bytes())

    def do_POST(self):
        if not self._request_is_local() or self.headers.get("Origin") != f"http://{HOST}:{PORT}":
            self._send(403, "application/json; charset=utf-8", json.dumps({"error": "Local access only."}))
            return
        if urlsplit(self.path).path != "/api/publish/catalog":
            self._send(404, "application/json; charset=utf-8", json.dumps({"error": "Not found."}))
            return
        if self.headers.get_content_type() != "application/json":
            self._send(415, "application/json; charset=utf-8", json.dumps({"error": "Expected a JSON publish request."}))
            return
        try:
            content_length = int(self.headers.get("Content-Length", ""))
        except ValueError:
            self._send(400, "application/json; charset=utf-8", json.dumps({"error": "Invalid request size."}))
            return
        if content_length <= 0 or content_length > MAX_PUBLISH_BODY_BYTES:
            self._send(413, "application/json; charset=utf-8", json.dumps({"error": "The publish request is empty or exceeds the 128 MB limit."}))
            return
        try:
            payload = json.loads(self.rfile.read(content_length))
            if not isinstance(payload, dict):
                raise ValueError("The publish request is invalid.")
            result = publish_pos_catalog(payload)
        except ValueError as error:
            self._send(400, "application/json; charset=utf-8", json.dumps({"error": str(error)}))
            return
        except Exception as error:
            print(f"POS catalog publish failed: {error}")
            self._send(502, "application/json; charset=utf-8", json.dumps({"error": str(error)}))
            return
        self._send(200, "application/json; charset=utf-8", json.dumps(result))

    def log_message(self, format_string, *args):
        print(f"127.0.0.1 - {format_string % args}")


def main():
    if sys.platform != "win32":
        raise SystemExit("The QuickBooks POS importer must run on Windows.")
    if not re.fullmatch(r"[A-Za-z0-9_]+", ITEM_CODE_FIELD):
        raise SystemExit("QBPOS_ITEM_CODE_FIELD must be a POS XML field name.")
    if "--discover" in sys.argv[1:]:
        discover_pos_fields()
        return
    server = ThreadingHTTPServer((HOST, PORT), ImporterHandler)
    server.daemon_threads = True
    print(f"Open http://{HOST}:{PORT}/ in Chrome or Edge.")
    print("Keep this window open while importing from QuickBooks POS.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nQuickBooks POS importer stopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
