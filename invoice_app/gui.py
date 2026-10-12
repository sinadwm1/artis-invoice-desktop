import tkinter as tk
from tkinter import ttk, messagebox
from database import DatabaseManager
from excel_processor import ExcelProcessor, resource_path
from config import CELL_MAPPING, ITEM_MAPPING, ITEM_ROWS, MAX_ITEMS, TEMPLATES
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os
import threading
import logging
import subprocess
import sys

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ARTIS Invoice and Proforma Invoice Maker")
        self.root.geometry("1000x700")

        # تنظیم آیکون برنامه
        try:
            icon_path = resource_path('logo.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            logger.warning(f"آیکون یافت نشد یا خطا: {e}")

        # استایل مدرن
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # فونت‌ها
        self.font_title = ('Tahoma', 14, 'bold')
        self.font_label = ('Tahoma', 10)
        self.font_button = ('Tahoma', 10, 'bold')
        self.font_table = ('Tahoma', 9)

        # رنگ‌ها
        self.primary_color = '#3498db'
        self.secondary_color = '#2c3e50'
        self.success_color = '#27ae60'
        self.background_color = '#ecf0f1'

        # استایل‌ها
        self.style.configure('.', font=self.font_label, background=self.background_color)
        self.style.configure('TButton', font=self.font_button, foreground='white')
        self.style.map('TButton', background=[('active', self.primary_color), ('!disabled', self.secondary_color)])
        self.style.configure('TFrame', background=self.background_color)
        self.style.configure('Treeview', font=self.font_table, rowheight=25, fieldbackground='white')
        self.style.configure('Treeview.Heading', font=('Tahoma', 9, 'bold'),
                             background=self.secondary_color, foreground='white')
        self.style.configure('TLabelframe', font=self.font_label, background=self.background_color)
        self.style.configure('TLabelframe.Label', font=self.font_label, background=self.background_color)
        self.style.configure('Title.TLabel', font=self.font_title,
                             background=self.primary_color, foreground='white', padding=5)
        self.style.configure('Primary.TLabel', font=('Tahoma', 10, 'bold'),
                             foreground=self.secondary_color)
        self.style.configure('Success.TButton', background=self.success_color)
        self.style.configure('custom.Horizontal.TProgressbar', background=self.success_color, thickness=20)
        self.style.configure('Status.TLabel', background=self.secondary_color,
                             foreground='white', font=('Tahoma', 9), padding=5)

        # دیتابیس و پردازشگر اکسل
        self.db = DatabaseManager()
        self.excel_processor = ExcelProcessor()

        # وضعیت
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, style='Status.TLabel')
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_var.set("آماده")

        # currency map (کمی بزرگ؛ می‌تونی موارد بیشتری اضافه کنی)
        self.CURRENCY_OPTIONS = [
            ("USD", "$", "US Dollar"),
            ("EUR", "€", "Euro"),
            ("GBP", "£", "British Pound"),
            ("IRR", "﷼", "Iranian Rial"),
            ("AED", "د.إ", "UAE Dirham"),
            ("SAR", "﷼", "Saudi Riyal"),
            ("JPY", "¥", "Japanese Yen"),
            ("RUB", "₽", "Russian Ruble"),
            ("TRY", "₺", "Turkish Lira"),
            ("INR", "₹", "Indian Rupee")
        ]
        self.currency_map = {f"{c[0]} ({c[1]})": (c[0], c[1]) for c in self.CURRENCY_OPTIONS}

        self.setup_ui()

    def setup_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(padx=10, pady=10, expand=True, fill='both')

        invoice_frame = ttk.Frame(notebook)
        notebook.add(invoice_frame, text='New Invoice')
        self.setup_invoice_tab(invoice_frame)

        customer_frame = ttk.Frame(notebook)
        notebook.add(customer_frame, text='Customers')
        self.setup_customer_tab(customer_frame)

        products_frame = ttk.Frame(notebook)
        notebook.add(products_frame, text='Products')
        self.setup_products_tab(products_frame)

        search_frame = ttk.Frame(notebook)
        notebook.add(search_frame, text='Search Invoices')
        self.setup_search_tab(search_frame)

    # --- تب New Invoice ---
    def setup_invoice_tab(self, frame):
        title = ttk.Label(frame, text="ایجاد پیش‌فاکتور جدید", style='Title.TLabel')
        title.pack(fill=tk.X, padx=5, pady=5)

        self.step1_frame = ttk.LabelFrame(frame, text="مرحله ۱: اطلاعات پایه")
        self.step1_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.step2_frame = ttk.LabelFrame(frame, text="مرحله ۲: اطلاعات تکمیلی")
        # مخفی در ابتدا
        self.step2_frame.pack_forget()

        self.setup_step1(self.step1_frame)
        self.setup_step2(self.step2_frame)

        nav_frame = ttk.Frame(frame)
        nav_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

        self.back_btn = ttk.Button(nav_frame, text="بازگشت", command=self.show_step1, state=tk.DISABLED)
        self.back_btn.pack(side=tk.LEFT, padx=5)
        self.next_btn = ttk.Button(nav_frame, text="مرحله بعد", command=self.show_step2)
        self.next_btn.pack(side=tk.LEFT, padx=5)
        self.generate_btn = ttk.Button(nav_frame, text="صدور فاکتور", command=self.generate_invoice,
                                       state=tk.DISABLED, style='Success.TButton')
        self.generate_btn.pack(side=tk.RIGHT, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(nav_frame, variable=self.progress_var,
                                            style='custom.Horizontal.TProgressbar', mode='determinate')
        self.progress_bar.pack_forget()

    def setup_step1(self, frame):
        # تاریخ
        date_frame = ttk.Frame(frame)
        date_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(date_frame, text="Issue Date (DD/MM/YY):", style='Primary.TLabel')\
            .pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(date_frame, text="Day:").pack(side=tk.LEFT)
        self.day_entry = ttk.Entry(date_frame, width=3)
        self.day_entry.pack(side=tk.LEFT, padx=(5, 15))
        ttk.Label(date_frame, text="Month:").pack(side=tk.LEFT)
        self.month_entry = ttk.Entry(date_frame, width=3)
        self.month_entry.pack(side=tk.LEFT, padx=(5, 15))
        ttk.Label(date_frame, text="Year:").pack(side=tk.LEFT)
        self.year_entry = ttk.Entry(date_frame, width=3)
        self.year_entry.pack(side=tk.LEFT)

        # مشتری
        customer_frame = ttk.Frame(frame)
        customer_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(customer_frame, text="Select Customer:", style='Primary.TLabel')\
            .pack(side=tk.LEFT, padx=(0, 5))
        self.customer_var = tk.StringVar()
        customers = [c[1] for c in self.db.get_customers()]
        self.customer_cb = ttk.Combobox(customer_frame, textvariable=self.customer_var,
                                        values=customers, width=40)
        self.customer_cb.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # تمپلیت
        template_frame = ttk.Frame(frame)
        template_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(template_frame, text="Select Template:", style='Primary.TLabel')\
            .pack(side=tk.LEFT, padx=(0, 5))
        self.template_var = tk.StringVar()
        templates = list(TEMPLATES.keys())
        self.template_cb = ttk.Combobox(template_frame, textvariable=self.template_var,
                                        values=templates, width=40)
        self.template_cb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if templates:
            self.template_cb.current(0)

        # جدول محصولات
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        columns = ('description', 'quantity', 'unit_price', 'total')
        self.items_table = ttk.Treeview(table_frame, columns=columns, show='headings', height=8)
        # تنظیم ستون‌ها
        self.items_table.heading('description', text='Description')
        self.items_table.heading('quantity', text='QTY')
        self.items_table.heading('unit_price', text='Unit Price ($)')
        self.items_table.heading('total', text='Total ($)')
        self.items_table.column('description', width=300, anchor='w')
        self.items_table.column('quantity', width=70, anchor='center')
        self.items_table.column('unit_price', width=120, anchor='e')
        self.items_table.column('total', width=120, anchor='e')
        # اسکرول
        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.items_table.yview)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.items_table.xview)
        self.items_table.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.items_table.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # دکمه‌ها
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(btn_frame, text="Add Product", command=self.add_product_to_invoice).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Remove Selected", command=self.remove_selected_product).pack(side=tk.LEFT, padx=5)

    def setup_step2(self, frame):
        form_frame = ttk.Frame(frame)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # پرداخت
        independent_frame = ttk.LabelFrame(form_frame, text="اطلاعات پرداخت")
        independent_frame.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')

        # حمل و نقل
        shipping_frame = ttk.LabelFrame(form_frame, text="اطلاعات حمل و نقل")
        shipping_frame.grid(row=1, column=0, padx=5, pady=5, sticky='nsew')

        # هزینه‌ها
        cost_frame = ttk.LabelFrame(form_frame, text="هزینه‌ها")
        cost_frame.grid(row=0, column=1, rowspan=2, padx=5, pady=5, sticky='nsew')

        self.independent_vars = {}
        self.shipping_vars = {}
        self.shipping_cost_vars = {}

        # فیلدهای پرداخت
        for i, (label, field) in enumerate([("Term of Payment:", "term_of_payment")] ):
            ttk.Label(independent_frame, text=label, style='Primary.TLabel')\
                .grid(row=i, column=0, padx=5, pady=5, sticky='e')
            var = tk.StringVar()
            ttk.Entry(independent_frame, textvariable=var, width=30)\
                .grid(row=i, column=1, padx=5, pady=5, sticky='ew')
            self.independent_vars[field] = var

        # فیلدهای حمل و نقل
        shipping_fields = [
            ("Mode of Transport:", "mode_of_transport"),
            ("Terms Of Delivery:", "terms_of_delivery"),
            ("Load Date:", "load_date"),
            ("Partial Shipment:", "partial"),
            ("Country Of Beneficiary:", "beneficiary_country"),
            ("Country Of Origin:", "origin_country"),
            ("Weight Per Unit:", "weight_per_unit"),
            ("Dimension (LxWxH):", "dimensions"),
            ("Total Weight:", "total_weight")
        ]
        for i, (label, field) in enumerate(shipping_fields):
            ttk.Label(shipping_frame, text=label, style='Primary.TLabel')\
                .grid(row=i, column=0, padx=5, pady=5, sticky='e')
            var = tk.StringVar()
            entry = ttk.Entry(shipping_frame, textvariable=var, width=30)
            entry.grid(row=i, column=1, padx=5, pady=5, sticky='ew')
            self.shipping_vars[field] = var

        # HS Code (invoice-level) اضافه کردن فیلد جدید
        ttk.Label(shipping_frame, text="Invoice HS Code:", style='Primary.TLabel')\
            .grid(row=len(shipping_fields), column=0, padx=5, pady=5, sticky='e')
        self.custom_hs_var = tk.StringVar()
        ttk.Entry(shipping_frame, textvariable=self.custom_hs_var, width=30).grid(row=len(shipping_fields), column=1, padx=5, pady=5, sticky='ew')

        # هزینه حمل و نقل
        ttk.Label(cost_frame, text="Shipping Cost:", style='Primary.TLabel')\
            .grid(row=0, column=0, padx=5, pady=5, sticky='e')
        var = tk.StringVar()
        e = ttk.Entry(cost_frame, textvariable=var, width=15)
        e.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        self.shipping_cost_vars['shipping_cost'] = var
        var.trace_add('write', lambda *_: self.update_totals_display())

        # انتخاب ارز
        ttk.Label(cost_frame, text="Currency:", style='Primary.TLabel')\
            .grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.currency_var = tk.StringVar()
        values = list(self.currency_map.keys())
        self.currency_cb = ttk.Combobox(cost_frame, textvariable=self.currency_var, values=values, width=20)
        self.currency_cb.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        if values:
            self.currency_cb.current(0)
        self.currency_cb.bind('<<ComboboxSelected>>', lambda e: self.on_currency_changed())

        # نمایش Total + نماد
        ttk.Label(cost_frame, text="Grand Total:", style='Primary.TLabel')\
            .grid(row=2, column=0, padx=5, pady=15, sticky='e')
        self.total_display_var = tk.StringVar(value="0.00")
        self.total_display_label = ttk.Label(cost_frame, textvariable=self.total_display_var, font=('Tahoma', 12, 'bold'))
        self.total_display_label.grid(row=2, column=1, padx=5, pady=15, sticky='w')

        form_frame.grid_rowconfigure(0, weight=1)
        form_frame.grid_rowconfigure(1, weight=1)
        form_frame.grid_columnconfigure(0, weight=1)
        form_frame.grid_columnconfigure(1, weight=1)

    def show_step1(self):
        self.step2_frame.pack_forget()
        self.step1_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.back_btn.config(state=tk.DISABLED)
        self.next_btn.config(state=tk.NORMAL)
        self.generate_btn.config(state=tk.DISABLED)

    def show_step2(self):
        try:
            day = self.day_entry.get().strip().zfill(2)
            month = self.month_entry.get().strip().zfill(2)
            year = self.year_entry.get().strip().zfill(2)
            if len(day) != 2 or len(month) != 2 or len(year) != 2:
                raise ValueError
            if not self.customer_var.get() or not self.template_var.get():
                raise ValueError
            if not self.items_table.get_children():
                raise ValueError
            self.step1_frame.pack_forget()
            self.step2_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
            self.back_btn.config(state=tk.NORMAL)
            self.next_btn.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.NORMAL)
            # update totals when user navigates to step2
            self.update_totals_display()
        except:
            messagebox.showerror("خطا", "لطفاً همه فیلدها را بدرستی پر کنید و حداقل یک محصول اضافه شود.")

    def generate_invoice(self):
        self.back_btn.config(state=tk.DISABLED)
        self.generate_btn.config(state=tk.DISABLED)
        self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=10)
        self.progress_var.set(0)
        self.status_var.set("در حال ایجاد فاکتور...")
        threading.Thread(target=self._generate_invoice_thread, daemon=True).start()

    def _generate_invoice_thread(self):
        try:
            # 1. Issue date
            day = self.day_entry.get().strip().zfill(2)
            month = self.month_entry.get().strip().zfill(2)
            year = self.year_entry.get().strip().zfill(2)
            issue_date_str = f"{day}/{month}/{year}"
            issue_date = datetime.strptime(issue_date_str, '%d/%m/%y')
            validation_date_str = (issue_date + relativedelta(months=6)).strftime('%d/%m/%y')

            # 2. Customer data
            selected = self.customer_var.get()
            customer_data = {}
            for c in self.db.get_customers():
                if c[1] == selected:
                    customer_data = {
                        'company_name': c[1],
                        'national_id': c[2],
                        'business_card_number': c[3],
                        'address': c[4],
                        'website': c[5],
                        'phone': c[6],
                        'fax': c[7]
                    }
                    break

            # 3. Invoice number
            biz = customer_data.get('business_card_number', '000')
            last_three = biz[-3:].zfill(3)
            date_part = issue_date_str.replace('/', '')
            counter = self.db.get_daily_counter(issue_date_str)
            counter_str = f"{counter:02d}"
            invoice_no = f"{last_three}{date_part}{counter_str}"

            # 4. Items
            items = []
            for iid in self.items_table.get_children():
                vals = self.items_table.item(iid, 'values')
                # vals: (description, qty, unit_price, total)
                items.append({
                    'description': vals[0],
                    'quantity': int(vals[1]),
                    'unit_price': float(vals[2])
                })

            # 5. Independent & shipping data
            independent = {f: v.get() for f, v in self.independent_vars.items() if v.get()}
            shipping = {f: v.get() for f, v in self.shipping_vars.items() if v.get()}
            try:
                shipping_cost = float(self.shipping_cost_vars['shipping_cost'].get())
            except:
                shipping_cost = 0.0

            # currency selection
            currency_sel = self.currency_var.get()
            currency_code, currency_symbol = self.currency_map.get(currency_sel, ('', ''))

            # 6. Totals
            total_products = sum(i['quantity'] * i['unit_price'] for i in items)
            grand_total = total_products + shipping_cost

            # 7. Build data & process
            data = {
                'invoice_no': invoice_no,
                'issue_date': issue_date_str,
                'validation_date': validation_date_str,
                'buyer': customer_data,
                'shipping': shipping,
                'items': items,
                'grand_total': grand_total,
                'shipping_cost': shipping_cost,
                'custom_hs_code': self.custom_hs_var.get() or None,
                'currency_code': currency_code,
                'currency_symbol': currency_symbol,
                **independent
            }

            self.status_var.set("در حال ایجاد فایل اکسل...")
            self.progress_var.set(30)
            tpl = TEMPLATES.get(self.template_var.get())
            template_path = resource_path(os.path.join('templates', tpl))
            excel_path = self.excel_processor.fill_invoice(data, template_path)

            self.status_var.set("در حال تبدیل به PDF...")
            self.progress_var.set(70)
            pdf_path = self.excel_processor.convert_to_pdf(excel_path)

            self.progress_var.set(100)
            self.status_var.set("عملیات با موفقیت انجام شد")
            messagebox.showinfo("موفقیت", f"پیش‌فاکتور {invoice_no} ایجاد شد:\n{pdf_path}")
            self.root.after(0, self.reset_form)

        except Exception as e:
            messagebox.showerror("خطا", str(e))
        finally:
            self.root.after(0, self.progress_bar.pack_forget)
            self.root.after(0, lambda: self.status_var.set("آماده"))
            self.root.after(0, lambda: self.back_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.generate_btn.config(state=tk.NORMAL))

    def reset_form(self):
        for e in (self.day_entry, self.month_entry, self.year_entry):
            e.delete(0, tk.END)
        self.customer_var.set('')
        self.template_var.set('')
        for iid in self.items_table.get_children():
            self.items_table.delete(iid)
        for v in self.independent_vars.values():
            v.set('')
        for v in self.shipping_vars.values():
            v.set('')
        for v in self.shipping_cost_vars.values():
            v.set('')
        self.custom_hs_var.set('')
        self.currency_var.set('')
        self.update_totals_display()
        self.show_step1()

    def add_product_to_invoice(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Add Product")
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg, text="Product:", style='Primary.TLabel').grid(row=0, column=0, padx=5, pady=5)
        prods = [p[1] for p in self.db.get_products()]
        cb = ttk.Combobox(dlg, values=prods, width=30)
        cb.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(dlg, text="Quantity:", style='Primary.TLabel').grid(row=1, column=0, padx=5, pady=5)
        qty = ttk.Entry(dlg)
        qty.insert(0, "1")
        qty.grid(row=1, column=1, padx=5, pady=5)

        # نمایش اطلاعات محصول انتخابی (اختیاری ولی مفید)
        info_lbl = ttk.Label(dlg, text="", wraplength=300)
        info_lbl.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

        def on_prod_select(event=None):
            pn = cb.get()
            p = next((p for p in self.db.get_products() if p[1] == pn), None)
            if p:
                desc = f"HS: {p[2] or '-'} | Price: {p[3]} | Dimensions: {p[5] or '-'}"
                info_lbl.config(text=desc)
            else:
                info_lbl.config(text="")

        cb.bind('<<ComboboxSelected>>', on_prod_select)

        def add_it():
            pn = cb.get()
            if not pn:
                messagebox.showerror("خطا", "لطفاً محصولی انتخاب کنید")
                return
            try:
                q = int(qty.get())
            except:
                q = 1
            prod = next((p for p in self.db.get_products() if p[1] == pn), None)
            up = prod[3] if prod else 0
            self.items_table.insert('', 'end', values=(pn, q, up, q * up))

            # مهم: ابعاد محصول را خودکار در فیلد ابعاد ارسال قرار بده
            if prod:
                dims = prod[5] or ''
                self.shipping_vars['dimensions'].set(dims)

            self.update_totals_display()
            dlg.destroy()

        btnf = ttk.Frame(dlg)
        btnf.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(btnf, text="Cancel", command=dlg.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(btnf, text="Add", command=add_it, style='Success.TButton').pack(side=tk.LEFT, padx=5)

    def remove_selected_product(self):
        for iid in self.items_table.selection():
            self.items_table.delete(iid)
        self.update_totals_display()

    def update_totals_display(self):
        try:
            items = []
            for iid in self.items_table.get_children():
                vals = self.items_table.item(iid, 'values')
                items.append({'quantity': int(vals[1]), 'unit_price': float(vals[2])})
            total_products = sum(i['quantity'] * i['unit_price'] for i in items)
            try:
                shipping_cost = float(self.shipping_cost_vars['shipping_cost'].get())
            except:
                shipping_cost = 0.0
            grand_total = total_products + shipping_cost
            currency_sel = self.currency_var.get()
            code, sym = self.currency_map.get(currency_sel, ('', ''))
            display = f"{code} {sym} {grand_total:,.2f}".strip()
            self.total_display_var.set(display)
        except Exception as e:
            logger.error(f"خطا در بروز رسانی نمایش مجموع: {e}")
            self.total_display_var.set("0.00")

    def on_currency_changed(self):
        self.update_totals_display()

    # --- تب Customers ---
    def setup_customer_tab(self, frame):
        title = ttk.Label(frame, text="مدیریت مشتریان", style='Title.TLabel')
        title.pack(fill=tk.X, padx=5, pady=5)

        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        cols = ('id', 'company', 'phone', 'address', 'card')
        self.customer_table = ttk.Treeview(table_frame, columns=cols, show='headings', height=15)
        for col, txt, width in [('id','ID',50),('company','Company',200),('phone','Phone',120),
                                 ('address','Address',250),('card','Business Card',100)]:
            self.customer_table.heading(col, text=txt)
            self.customer_table.column(col, width=width, anchor='center' if col in ('id','phone','card') else 'w')
        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.customer_table.yview)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.customer_table.xview)
        self.customer_table.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.customer_table.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(btn_frame, text="Add Customer", command=self.add_customer).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Edit", command=self.edit_customer).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Delete", command=self.delete_customer).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Refresh", command=self.refresh_customers).pack(side=tk.LEFT, padx=5)

        self.refresh_customers()

    def add_customer(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Add New Customer")
        dlg.resizable(False, False)
        dlg.grab_set()
        fields = [
            ('Company Name:', 'company'),
            ('National ID:', 'national_id'),
            ('Business Card No:', 'business_card'),
            ('Address:', 'address'),
            ('Website:', 'website'),
            ('Phone:', 'phone'),
            ('Fax:', 'fax')
        ]
        ent = {}
        for i, (lbl, fn) in enumerate(fields):
            ttk.Label(dlg, text=lbl, style='Primary.TLabel').grid(row=i, column=0, padx=5, pady=5, sticky='e')
            e = ttk.Entry(dlg, width=30)
            e.grid(row=i, column=1, padx=5, pady=5)
            ent[fn] = e

        def save_c():
            data = (
                ent['company'].get(),
                ent['national_id'].get(),
                ent['business_card'].get(),
                ent['address'].get(),
                ent['website'].get(),
                ent['phone'].get(),
                ent['fax'].get()
            )
            self.db.add_customer(data)
            self.refresh_customers()
            dlg.destroy()

        btn_frame = ttk.Frame(dlg)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save", command=save_c, style='Success.TButton').pack(side=tk.LEFT, padx=5)

    def edit_customer(self):
        sel = self.customer_table.selection()
        if not sel:
            messagebox.showwarning("اخطار", "لطفاً مشتری انتخاب کنید")
            return
        cid = self.customer_table.item(sel[0], 'values')[0]
        c = self.db.get_customer_by_id(cid)
        if not c:
            messagebox.showerror("خطا", "مشتری یافت نشد")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Edit Customer")
        dlg.resizable(False, False)
        dlg.grab_set()
        fields = [
            ('Company Name:', 'company', c[1]),
            ('National ID:', 'national_id', c[2]),
            ('Business Card No:', 'business_card', c[3] or ''),
            ('Address:', 'address', c[4]),
            ('Website:', 'website', c[5]),
            ('Phone:', 'phone', c[6]),
            ('Fax:', 'fax', c[7])
        ]
        ent = {}
        for i, (lbl, fn, val) in enumerate(fields):
            ttk.Label(dlg, text=lbl, style='Primary.TLabel').grid(row=i, column=0, padx=5, pady=5, sticky='e')
            e = ttk.Entry(dlg, width=30)
            e.insert(0, val)
            e.grid(row=i, column=1, padx=5, pady=5)
            ent[fn] = e

        def save_e():
            data = (
                ent['company'].get(),
                ent['national_id'].get(),
                ent['business_card'].get(),
                ent['address'].get(),
                ent['website'].get(),
                ent['phone'].get(),
                ent['fax'].get()
            )
            self.db.update_customer(cid, data)
            self.refresh_customers()
            dlg.destroy()

        btn_frame = ttk.Frame(dlg)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save", command=save_e, style='Success.TButton').pack(side=tk.LEFT, padx=5)

    def delete_customer(self):
        sel = self.customer_table.selection()
        if not sel:
            messagebox.showwarning("اخطار", "لطفاً انتخاب کنید")
            return
        cid = self.customer_table.item(sel[0], 'values')[0]
        if messagebox.askyesno("تایید حذف", "مطمئن هستید؟"):
            self.db.delete_customer(cid)
            self.refresh_customers()

    def refresh_customers(self):
        for i in self.customer_table.get_children():
            self.customer_table.delete(i)
        for c in self.db.get_customers():
            self.customer_table.insert('', 'end', values=(c[0], c[1], c[6], c[4], c[3]))

    # --- تب Products ---
    def setup_products_tab(self, frame):
        title = ttk.Label(frame, text="مدیریت محصولات", style='Title.TLabel')
        title.pack(fill=tk.X, padx=5, pady=5)

        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        cols = ('id', 'description', 'price', 'hs_code')
        self.product_table = ttk.Treeview(table_frame, columns=cols, show='headings', height=15)
        for col, txt, width in [('id','ID',50),('description','Description',300),
                                 ('price','Price ($)',100),('hs_code','HS Code',100)]:
            self.product_table.heading(col, text=txt)
            self.product_table.column(col, width=width,
                                      anchor='center' if col=='id' else 'e' if col=='price' else 'w')
        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.product_table.yview)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.product_table.xview)
        self.product_table.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.product_table.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(btn_frame, text="Add Product", command=self.add_product).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Edit", command=self.edit_product).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Delete", command=self.delete_product).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Refresh", command=self.refresh_products).pack(side=tk.LEFT, padx=5)

        self.refresh_products()

    def add_product(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Add New Product")
        dlg.resizable(False, False)
        dlg.grab_set()
        fields = [
            ('Description:', 'description', ''),
            ('HS Code:', 'hs_code', ''),
            ('Unit Price ($):', 'unit_price', '0.0'),
            ('Weight per Unit:', 'weight', '0.0'),
            ('Dimensions (LxWxH):', 'dimensions', '')
        ]
        ent = {}
        for i, (lbl, fn, dft) in enumerate(fields):
            ttk.Label(dlg, text=lbl, style='Primary.TLabel').grid(row=i, column=0, padx=5, pady=5, sticky='e')
            e = ttk.Entry(dlg, width=30)
            e.insert(0, dft)
            e.grid(row=i, column=1, padx=5, pady=5)
            ent[fn] = e

        def save_p():
            try:
                up = float(ent['unit_price'].get())
                wt = float(ent['weight'].get())
            except:
                messagebox.showerror("خطا", "قیمت و وزن باید عدد باشند")
                return
            data = (
                ent['description'].get(),
                ent['hs_code'].get(),
                up,
                wt,
                ent['dimensions'].get()
            )
            self.db.add_product(data)
            self.refresh_products()
            dlg.destroy()

        btn_frame = ttk.Frame(dlg)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save", command=save_p, style='Success.TButton').pack(side=tk.LEFT, padx=5)

    def edit_product(self):
        sel = self.product_table.selection()
        if not sel:
            messagebox.showwarning("اخطار", "لطفاً انتخاب کنید")
            return
        pid = self.product_table.item(sel[0], 'values')[0]
        p = self.db.get_product_by_id(pid)
        if not p:
            messagebox.showerror("خطا", "محصول یافت نشد")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Edit Product")
        dlg.resizable(False, False)
        dlg.grab_set()
        fields = [
            ('Description:', 'description', p[1]),
            ('HS Code:', 'hs_code', p[2]),
            ('Unit Price ($):', 'unit_price', str(p[3])),
            ('Weight per Unit:', 'weight', str(p[4])),
            ('Dimensions (LxWxH):', 'dimensions', p[5] or '')
        ]
        ent = {}
        for i, (lbl, fn, val) in enumerate(fields):
            ttk.Label(dlg, text=lbl, style='Primary.TLabel').grid(row=i, column=0, padx=5, pady=5, sticky='e')
            e = ttk.Entry(dlg, width=30)
            e.insert(0, val)
            e.grid(row=i, column=1, padx=5, pady=5)
            ent[fn] = e

        def save_pe():
            try:
                up = float(ent['unit_price'].get())
                wt = float(ent['weight'].get())
            except:
                messagebox.showerror("خطا", "عدد وارد کنید")
                return
            data = (
                ent['description'].get(),
                ent['hs_code'].get(),
                up,
                wt,
                ent['dimensions'].get()
            )
            self.db.update_product(pid, data)
            self.refresh_products()
            dlg.destroy()

        btn_frame = ttk.Frame(dlg)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save", command=save_pe, style='Success.TButton').pack(side=tk.LEFT, padx=5)

    def delete_product(self):
        sel = self.product_table.selection()
        if not sel:
            messagebox.showwarning("اخطار", "لطفاً انتخاب کنید")
            return
        pid = self.product_table.item(sel[0], 'values')[0]
        if messagebox.askyesno("تایید حذف", "مطمئنید؟"):
            self.db.delete_product(pid)
            self.refresh_products()

    def refresh_products(self):
        for i in self.product_table.get_children():
            self.product_table.delete(i)
        for p in self.db.get_products():
            self.product_table.insert('', 'end', values=(p[0], p[1], p[3], p[2]))

    # --- تب Search Invoices ---
    def setup_search_tab(self, frame):
        title = ttk.Label(frame, text="جستجوی پیش‌فاکتورها", style='Title.TLabel')
        title.pack(fill=tk.X, padx=5, pady=5)

        search_frame = ttk.LabelFrame(frame, text="معیارهای جستجو")
        search_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(search_frame, text="Search by:", style='Primary.TLabel')\
            .grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.search_mode = tk.StringVar(value='card_last3')
        modes_frame = ttk.Frame(search_frame)
        modes_frame.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        modes = [('Card (auto 3 digits)', 'card_last3'),
                 ('Enter 3 digits direct', 'direct_last3'),
                 ('Full date (DDMMYY)', 'full_date')]
        for text, val in modes:
            ttk.Radiobutton(modes_frame, text=text, variable=self.search_mode, value=val)\
                .pack(side=tk.LEFT, padx=5)

        ttk.Label(search_frame, text="Value:", style='Primary.TLabel')\
            .grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.search_entry = ttk.Entry(search_frame, width=20)
        self.search_entry.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        ttk.Button(search_frame, text="Search", command=self.perform_search).grid(row=1, column=2, padx=5)

        result_frame = ttk.LabelFrame(frame, text="نتایج جستجو")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.result_list = tk.Listbox(result_frame, font=self.font_table, bg='white',
                                     selectbackground=self.primary_color, height=15)
        self.result_list.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.result_list.bind('<Double-Button-1>', lambda e: self.open_selected())

        open_btn = ttk.Button(result_frame, text="Open Selected",
                              command=self.open_selected, style='Success.TButton')
        open_btn.pack(side=tk.BOTTOM, padx=5, pady=5)

    def perform_search(self):
        mode = self.search_mode.get()
        val = self.search_entry.get().strip()
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
            pdf_dir = os.path.join(base_dir, 'output', 'pdfs')
        else:
            pdf_dir = 'output/pdfs'

        if not os.path.isdir(pdf_dir):
            messagebox.showerror("خطا", "پوشه PDF یافت نشد")
            return

        files = []
        for fn in os.listdir(pdf_dir):
            if not fn.lower().endswith('.pdf'):
                continue
            inv = os.path.splitext(fn)[0].split('_')[-1]
            if mode == 'card_last3' and val[-3:] == inv[:3]:
                files.append(fn)
            elif mode == 'direct_last3' and val == inv[:3]:
                files.append(fn)
            elif mode == 'full_date' and val == inv[3:9]:
                files.append(fn)

        self.result_list.delete(0, tk.END)
        if not files:
            self.result_list.insert(tk.END, "فایلی پیدا نشد")
        else:
            for f in sorted(files):
                self.result_list.insert(tk.END, f)

    def open_selected(self):
        sel = self.result_list.curselection()
        if not sel:
            return
        fn = self.result_list.get(sel[0])
        if fn == "فایلی پیدا نشد":
            return
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
            path = os.path.join(base_dir, 'output', 'pdfs', fn)
        else:
            path = os.path.join('output/pdfs', fn)
        try:
            if sys.platform.startswith('darwin'):
                subprocess.call(('open', path))
            elif os.name == 'nt':
                os.startfile(path)
            else:
                subprocess.call(('xdg-open', path))
        except Exception as e:
            messagebox.showerror("خطا در باز کردن فایل", str(e))


if __name__ == '__main__':
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()
