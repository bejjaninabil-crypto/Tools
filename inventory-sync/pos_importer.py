"""Serve the local catalog importer and read stock from QuickBooks POS."""

import json
import math
import mimetypes
import os
import re
import sys
import threading
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
