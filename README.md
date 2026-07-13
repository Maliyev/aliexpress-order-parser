# AliExpress order parser

Small local prototype for moving one saved AliExpress order page into an Excel order block.

It does not log in to AliExpress, open a browser, or download anything. The user saves an already open order page with Ctrl+S, then runs the script on that local HTML file.

## What it reads

- order date and AliExpress order number
- tracking number
- seller name
- product title, variation, price and quantity
- order total, item subtotal and delivery price

Columns B, C and E are not available on the saved order page. The script leaves them empty unless a user provides them manually.

## Run locally

```powershell
python -m pip install -r requirements.txt

python src/parse_order.py "<saved-order.html>" `
  --workbook "<local-template.xlsx>" `
  --output "output\parsed-order.xlsx"
```

For a small workbook containing only the header and one order block, add `--new-workbook`:

```powershell
python src/parse_order.py "<saved-order.html>" `
  --workbook "<local-template.xlsx>" `
  --output "output\test-one-order.xlsx" `
  --new-workbook
```

The script refuses to overwrite the source workbook or an existing output file.
