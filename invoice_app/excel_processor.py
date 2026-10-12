from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet
from config import CELL_MAPPING, ITEM_MAPPING, ITEM_ROWS, MAX_ITEMS
import os
import logging
import pythoncom
import re
import subprocess
import sys

try:
    import win32com.client
except Exception:
    win32com = None

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

class ExcelProcessor:
    def __init__(self):
        pass

    def _get_main_cell(self, sheet: Worksheet, cell_address: str) -> Cell:
        try:
            col_letter = ''.join(filter(str.isalpha, cell_address))
            row = int(''.join(filter(str.isdigit, cell_address)))
            col = column_index_from_string(col_letter)

            for merged_range in list(sheet.merged_cells.ranges):
                if (merged_range.min_row <= row <= merged_range.max_row and
                    merged_range.min_col <= col <= merged_range.max_col):
                    return sheet.cell(merged_range.min_row, merged_range.min_col)

            return sheet[cell_address]
        except Exception as e:
            logger.error(f"خطا در یافتن سلول اصلی برای {cell_address}: {str(e)}")
            return sheet[cell_address]

    def fill_invoice(self, data, template_path, output_dir=None):
        try:
            if output_dir is None:
                if getattr(sys, 'frozen', False):
                    base_dir = os.path.dirname(sys.executable)
                    output_dir = os.path.join(base_dir, 'output', 'invoices')
                else:
                    output_dir = 'output/invoices'

            logger.debug("شروع پر کردن فاکتور")
            os.makedirs(output_dir, exist_ok=True)
            wb = load_workbook(template_path)
            sheet = wb['proforma']

            # پر کردن فیلدهای mapping عمومی
            for field, cell_address in CELL_MAPPING.items():
                value = None

                if field.startswith('buyer_'):
                    field_name = field.replace('buyer_', '')
                    buyer_data = data.get('buyer', {})
                    if field_name in buyer_data:
                        value = buyer_data[field_name]

                elif field.startswith('shipping_'):
                    field_name = field.replace('shipping_', '')
                    shipping_data = data.get('shipping', {})
                    if field_name in shipping_data:
                        value = shipping_data[field_name]

                elif field in data and data[field] is not None:
                    value = data[field]

                # نادیده گرفتن custom_hs_code و shipping_dimensions در این حلقه
                if field in ('custom_hs_code', 'shipping_dimensions'):
                    continue

                if value is not None:
                    try:
                        cell = self._get_main_cell(sheet, cell_address)
                        cell.value = value
                        logger.debug(f"پر کردن {field} در {cell_address}")
                    except Exception as e:
                        logger.error(f"خطا در پر کردن {field} در {cell_address}: {str(e)}")

            # Shipping cost (مطمئن شو که مقدار در mapping هست)
            if 'shipping_cost' in data and data['shipping_cost'] is not None:
                cell_address = CELL_MAPPING.get('shipping_cost')
                if cell_address:
                    try:
                        cell = self._get_main_cell(sheet, cell_address)
                        cell.value = data['shipping_cost']
                        logger.debug(f"پر کردن shipping_cost در {cell_address}")
                    except Exception as e:
                        logger.error(f"خطا در پر کردن shipping_cost: {str(e)}")

            # آیتم‌ها
            items = data.get('items', [])[:MAX_ITEMS]
            logger.debug(f"تعداد آیتم‌ها: {len(items)}")

            for idx, item in enumerate(items):
                row_number = ITEM_ROWS.get(idx + 1)
                if not row_number:
                    logger.warning(f"سطر مربوط به آیتم {idx+1} یافت نشد")
                    continue

                logger.debug(f"پر کردن آیتم {idx+1} در سطر {row_number}")

                for item_field, column in ITEM_MAPPING.items():
                    if item_field == 'total':
                        value = item.get('quantity', 0) * item.get('unit_price', 0)
                    else:
                        value = item.get(item_field, '')

                    cell_address = f"{column}{row_number}"

                    try:
                        cell = self._get_main_cell(sheet, cell_address)
                        cell.value = value
                        logger.debug(f"پر کردن {item_field} در {cell_address}")
                    except Exception as e:
                        logger.error(f"خطا در پر کردن آیتم {item_field} در {cell_address}: {str(e)}")

            # custom HS code (invoice-level)
            try:
                hs_val = data.get('custom_hs_code')
                cell_address = CELL_MAPPING.get('custom_hs_code')
                if hs_val is not None and cell_address:
                    cell = self._get_main_cell(sheet, cell_address)
                    cell.value = hs_val
                    logger.debug(f"پر کردن custom_hs_code در {cell_address}")
            except Exception as e:
                logger.error(f"خطا در پر کردن custom_hs_code: {str(e)}")

            # shipping dimensions (invoice-level) از دادهٔ shipping.dimensions می‌گیریم
            try:
                shipping_dims = data.get('shipping', {}).get('dimensions')
                cell_address = CELL_MAPPING.get('shipping_dimensions')
                if shipping_dims is not None and cell_address:
                    cell = self._get_main_cell(sheet, cell_address)
                    cell.value = shipping_dims
                    logger.debug(f"پر کردن shipping_dimensions در {cell_address}")
            except Exception as e:
                logger.error(f"خطا در پر کردن shipping_dimensions: {str(e)}")

            # grand total + currency (نمایش کد و نمادِ ارز کنار عدد)
            try:
                grand_total = data.get('grand_total', 0)
                cell_address = CELL_MAPPING.get('grand_total')
                if cell_address:
                    currency_code = data.get('currency_code', '')
                    currency_symbol = data.get('currency_symbol', '')
                    formatted_amount = f"{grand_total:,.2f}"
                    display_value = f"{currency_code} {currency_symbol} {formatted_amount}".strip()
                    cell = self._get_main_cell(sheet, cell_address)
                    cell.value = display_value
                    logger.debug(f"تنظیم مبلغ نهایی در {cell_address}")
            except Exception as e:
                logger.error(f"خطا در تنظیم مبلغ نهایی: {str(e)}")

            invoice_no = re.sub(r'[^\w-]+', '_', str(data.get('invoice_no', 'temp'))).strip('_') or 'temp'
            output_path = os.path.join(output_dir, f"invoice_{invoice_no}.xlsx")
            wb.save(output_path)
            logger.info("فایل اکسل ذخیره شد")
            return output_path
        except Exception as e:
            logger.error(f"خطا در پر کردن فاکتور: {str(e)}")
            raise

    def convert_to_pdf(self, excel_path, pdf_dir=None):
        try:
            if pdf_dir is None:
                if getattr(sys, 'frozen', False):
                    base_dir = os.path.dirname(sys.executable)
                    pdf_dir = os.path.join(base_dir, 'output', 'pdfs')
                else:
                    pdf_dir = 'output/pdfs'

            os.makedirs(pdf_dir, exist_ok=True)
            pdf_path = os.path.join(pdf_dir, os.path.basename(excel_path).replace('.xlsx', '.pdf'))

            logger.info("شروع تبدیل به PDF")

            try:
                pythoncom.CoInitialize()
                if win32com:
                    excel = win32com.client.Dispatch("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False

                    wb = None
                    try:
                        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
                        wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
                        logger.info("تبدیل PDF با موفقیت انجام شد (ویندوز)")
                    except Exception as e:
                        logger.error(f"خطا در تبدیل PDF (ویندوز): {str(e)}")
                        return excel_path
                    finally:
                        if wb:
                            wb.Close(False)
                        excel.Quit()
                        del wb
                        del excel
                        pythoncom.CoUninitialize()

                    return pdf_path
                else:
                    # fallback to shell unoconv if available
                    subprocess.run(['unoconv', '-f', 'pdf', '-o', pdf_path, excel_path], check=True)
                    logger.info("تبدیل PDF با موفقیت انجام شد (لینوکس/مک fallback)")
                    return pdf_path
            except Exception as e:
                logger.error(f"خطا در تبدیل PDF (ویندوز): {str(e)}")
                return excel_path
        except Exception as e:
            logger.error(f"خطا در تبدیل به PDF: {str(e)}")
            return excel_path
