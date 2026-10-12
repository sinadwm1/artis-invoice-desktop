# ARTIS Invoice Desktop

[راهنمای فارسی](README_FA.md)

A Windows desktop application for creating proforma invoices from customer and product records. It uses Excel workbooks for document generation and can export PDF through Microsoft Excel automation.

## Features verified in the source

- Tkinter interface with invoice, customer, product, and search sections.
- SQLite storage for customer and product records.
- Excel templates populated through `openpyxl`.
- Up to 16 invoice item rows in the prepared version; the source snapshot mapped only the first row, and this draft corrects that mapping.
- PDF export through Microsoft Excel COM on Windows.

## Requirements

- Windows, Python 3.13 recommended, and Microsoft Excel for PDF export.
- Python packages in `requirements.txt`.

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python invoice_app/main.py
```

The app creates its SQLite database and output folders during use. Generated invoices and database files are excluded from Git. Do not put customer details in sample commits.

## Template rights and design

The supplied source contained modified Vertex42 workbooks. Those workbooks and all generated invoices are excluded from this draft. The two files in `invoice_app/templates/` are newly authored, blank templates that keep the cell addresses expected by the application. Their look is intentionally simpler than the original templates. Export and visual layout should be reviewed in Excel before release.

## Limitations

The program has not been exercised end to end with a real Excel installation in this preparation environment. PDF export depends on local Excel COM availability. Released under the MIT License.
