# AliExpress order parser

Local prototype for moving one manually opened AliExpress order into an Excel order block.

The project has two separate parts:

1. The Chrome extension saves the DOM of the current order page when the user clicks its button.
2. The Python script finds that saved file, shows a preview, asks for fields missing from AliExpress, checks duplicates, then creates a backup and a new Excel copy.

The project does not log in to AliExpress, read cookies, make network requests, open other orders, or change the original Excel file.

## Setup

Install the Excel library once:

```powershell
python -m pip install -r requirements.txt
```

Copy `config.example.json` to `config.json` and update the local paths. `config.json` is ignored by Git because it contains local paths.

## Chrome extension

1. Open `chrome://extensions` in Chrome.
2. Enable Developer mode.
3. Click Load unpacked.
4. Select the `extension` folder from this project.
5. Open one AliExpress order manually and click the extension icon.
6. Click Save current order HTML.

The extension checks that the page has an order number and product blocks. It saves a UTF-8 file named `AliExpress-order-<order-number>.html` in Downloads.

## Process the latest download

Double-click `run_order_parser.bat`.

The script shows the order preview and asks for manual values for columns B, C and E. If AliExpress shows an order date without a year, it uses the current year. It stops if the order ID or tracking number is already in the workbook. If it is new, the script creates a backup, inserts the order at row 51, shifts existing rows down, and saves the updated copy in the configured output folder.

## Direct parser run

```powershell
python src/parse_order.py "<saved-order.html>" `
  --workbook "<local-template.xlsx>" `
  --output "output\parsed-order.xlsx"
```
