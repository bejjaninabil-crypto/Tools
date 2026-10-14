# Daychounieh Tools

Daychounieh Tools is a responsive product catalog with Arabic and English product names, SKU search, category filters, product photos, product details, a shopping bag, and WhatsApp order requests. The public storefront is hosted on GitHub Pages. Catalog imports are prepared locally on your computer.

## Website

The current live shop is at `https://daychouniehtools.github.io/Tools/`, published from `DaychouniehTools/Tools` on its `main` branch. Open `index.html` to view the bundled catalog locally. The reviewed catalog has 1,692 items; the `images/` folder may contain extra files not used by the current catalog. Search matches names, item codes, categories, descriptions, and tags. Customers can click a product photo to see a larger image and its details.

Generate and verify a new QR code after the GitHub Pages address changes. The homepage displays the shop contact number prominently; its number link starts a phone call, and the WhatsApp button opens a chat.

## Import products locally

The importer does not use Supabase, Google sign-in, or a remote database. It reads the workbook in your browser and writes the updated catalog and matching photos to a local project folder you choose. Do not publish `admin-import.html`.

### QuickBooks POS v9

On the Windows computer with QuickBooks POS installed, log in to POS and double-click `Start POS Importer.bat`. Keep `inventory-sync/pos_importer.py` with the launcher. The connector requires 32-bit Python and `pywin32`; if prompted, install the module with `py -3-32 -m pip install pywin32`. Keep the connector window open. Option 1 reads item codes, stock, item name (`Desc1`), description (`Desc2`), web price, department code, size, attribute, UPC, and keywords. It updates matching catalog items and can add new positive-stock items when a matching SKU-named photo exists and the item has a name, positive web price, and department category determined from existing catalog items. New items without a photo or required catalog details are skipped. It does not read the POS picture tick or download POS pictures. Changes are local until reviewed and published through GitHub Desktop; clicking Read does not update the live site.

1. Open `admin-import.html` from the Daychounieh Tools project folder in a recent version of Chrome or Edge.
2. Click **Select Excel file** and choose your inventory workbook (`.xlsx`/`.xls`) or CSV file. This only selects the workbook; it does not import it yet.
3. Optionally click **Choose photos folder** and select the folder with product photos. Name the main photo after its item code, such as `290.jpg` or `290_Def.jpg`; the importer saves either as `290.jpg`. For an optional second photo, name it `290_2.jpg` or `290_Alt.jpg`; product details show both photos with previous/next controls.
4. Click **Step 3: Import selected Excel and photos into local project** once. When your browser asks, select your local project folder containing `index.html`, `catalog.js`, and `images`—not the separate GitHub repository folder. Products in the spreadsheet update matching catalog items by item code. If a row has Qty 1 of zero, less than zero, blank, or nonnumeric, the importer asks you to confirm removing that matching item from the catalog; its photo file stays in `images/`. Products omitted from the workbook remain in the catalog, and matched photos are copied into `images/`.
5. Check the updated local storefront. To transfer the inventory update to the GitHub Desktop copy, click **Choose local project folder** and select the local project folder, then click **Choose GitHub repository folder** and select the repository folder containing `.git`, `index.html`, and `images`. Once both are selected, click **Copy catalog and photos to GitHub Desktop**. The importer overwrites `catalog.js` and same-named image files, does not delete extra images, and does not publish online. Review the copied changes and commit and push them through GitHub Desktop to publish.

To remove every catalog item that has no usable local or web picture, click **Remove items without pictures** next to the import button. The tool checks all catalog items against existing files in the project's `images/` folder and valid web image URLs, shows a confirmation with the number to remove, and leaves all image files untouched. Back up the project before confirming; removing catalog items cannot be undone by the importer.

The workbook's `Size` column is imported and shown on product cards and in product details. You can also add optional `Dimensions` and `Sizes` (or `Variants`) columns to show physical measurements separately from choices such as S/M/L. For example, `Dimensions` can contain `20 x 15 x 8 cm`, and `Sizes` can contain `Small, Medium, Large`. Search also matches these fields.

Excel's `_x000D_` marker in text fields represents a carriage return/line break. The importer converts it to a normal line break instead of showing the marker in the product description.

For the QuickBooks inventory export, the importer uses only the website fields: `Item Number` (item code), `Item Name`, `Web Price`, `Qty 1` (availability filter), `Size`, `Department Name` (category), `Attribute`, `Item Description` (or `Brief Description` if the item description is empty), `UPC` (barcode), and `Web keywords`. `Qty 1` is required: products whose quantity is zero, negative, blank, or nonnumeric are removed from the catalog during import and are not shown on the website; their photo files are not deleted. Attributes appear in the storefront product details. For every imported item code with positive quantity, its Attribute is copied exactly from the workbook; a blank cell clears the old value. Import stops rather than silently skipping rows with missing item codes or duplicate item codes. Product images are matched from the selected photo folder by item code. Accounting, cost, reorder, and other back-office columns are ignored.

The browser fetches the Excel-reading library from jsDelivr when opening the importer for `.xls`/`.xlsx` workbooks; the spreadsheet and photos stay on your computer and are not sent to that service. CSV import uses the browser directly. Keep a backup of the project folder before importing.

The local importer is excluded from the public GitHub repository by `.gitignore`. If you move the project to another computer, copy `admin-import.html` there separately if you need the import tool.

## Catalog and orders

Prices are displayed in USD. WhatsApp order requests are sent to `+9613032791`; change `WHATSAPP_NUMBER` in `index.html` to change the shop contact. Customers review and send the prepared order themselves. The site does not process card payments or automatically reserve stock.

The public site, product names, prices, item codes, and photos are visible to visitors. The local importer can only modify files after you open it on your computer and grant access to the project folder; it is not an online admin login.

## GitHub Pages

Keep `index.html`, `catalog.js`, `.nojekyll`, and the complete `images/` folder in `DaychouniehTools/Tools` on `main`. GitHub Pages serves the storefront from that repository. Import products locally, then review and publish the updated catalog and photos using GitHub Desktop. Changes pushed only to the `bejjaninabil-crypto/Tools` fork do not update the live website. Do not add `admin-import.html` to the public repository.
