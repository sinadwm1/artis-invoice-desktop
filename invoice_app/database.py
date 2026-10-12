# database.py

import sqlite3
import os
import sys

class DatabaseManager:
    def __init__(self, db_path=None):
        if db_path is None:
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
                db_path = os.path.join(base_path, 'data', 'invoice_data.db')
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
                db_path = os.path.join(base_path, 'data', 'invoice_data.db')
        
        self.db_path = db_path
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        self._create_tables()
    
    def _create_tables(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        c.execute('''CREATE TABLE IF NOT EXISTS customers (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        company_name TEXT NOT NULL,
                        national_id TEXT,
                        business_card_number TEXT,
                        address TEXT,
                        website TEXT,
                        phone TEXT,
                        fax TEXT
                    )''')
        
        c.execute('''CREATE TABLE IF NOT EXISTS products (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        description TEXT NOT NULL,
                        hs_code TEXT,
                        unit_price REAL,
                        weight_per_unit REAL,
                        dimensions TEXT
                    )''')
        
        c.execute('''CREATE TABLE IF NOT EXISTS daily_counters (
                        date TEXT PRIMARY KEY,
                        counter INTEGER DEFAULT 1
                    )''')
        
        conn.commit()
        conn.close()
    
    def add_customer(self, customer_data):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('''INSERT INTO customers 
                         (company_name, national_id, business_card_number, address, website, phone, fax)
                     VALUES (?, ?, ?, ?, ?, ?, ?)''', customer_data)
        conn.commit()
        conn.close()
    
    def update_customer(self, customer_id, customer_data):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('''UPDATE customers 
                     SET company_name=?, national_id=?, business_card_number=?, address=?, website=?, phone=?, fax=?
                     WHERE id=?''', (*customer_data, customer_id))
        conn.commit()
        conn.close()
    
    def delete_customer(self, customer_id):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('DELETE FROM customers WHERE id=?', (customer_id,))
        conn.commit()
        conn.close()
    
    def get_customers(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('SELECT * FROM customers')
        customers = c.fetchall()
        conn.close()
        return customers
    
    def get_customer_by_id(self, customer_id):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('SELECT * FROM customers WHERE id=?', (customer_id,))
        customer = c.fetchone()
        conn.close()
        return customer
    
    def add_product(self, product_data):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('''INSERT INTO products 
                         (description, hs_code, unit_price, weight_per_unit, dimensions)
                     VALUES (?, ?, ?, ?, ?)''', product_data)
        conn.commit()
        conn.close()
    
    def update_product(self, product_id, product_data):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('''UPDATE products 
                     SET description=?, hs_code=?, unit_price=?, weight_per_unit=?, dimensions=?
                     WHERE id=?''', (*product_data, product_id))
        conn.commit()
        conn.close()
    
    def delete_product(self, product_id):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('DELETE FROM products WHERE id=?', (product_id,))
        conn.commit()
        conn.close()
    
    def get_products(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('SELECT * FROM products')
        products = c.fetchall()
        conn.close()
        return products
    
    def get_product_by_id(self, product_id):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('SELECT * FROM products WHERE id=?', (product_id,))
        product = c.fetchone()
        conn.close()
        return product
    
    def get_daily_counter(self, issue_date):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        c.execute("SELECT counter FROM daily_counters WHERE date=?", (issue_date,))
        result = c.fetchone()
        
        if result:
            counter = result[0] + 1
            c.execute("UPDATE daily_counters SET counter=? WHERE date=?", (counter, issue_date))
        else:
            counter = 1
            c.execute("INSERT INTO daily_counters (date, counter) VALUES (?, ?)", (issue_date, counter))
        
        conn.commit()
        conn.close()
        
        if counter > 99:
            raise ValueError("امکان صدور بیش از 99 فاکتور در یک روز وجود ندارد")
        
        return counter

#  * Author:
#  *   Name:           sinaDWM
#  *   Telegram:       @sina_dwm
#  *
#  * 
                                                                                                  
#            88                           88888888ba,   I8,        8        ,8I  88b           d88  
#            ""                           88      `"8b  `8b       d8b       d8'  888b         d888  
#                                         88        `8b  "8,     ,8"8,     ,8"   88`8b       d8'88  
# ,adPPYba,  88  8b,dPPYba,   ,adPPYYba,  88         88   Y8     8P Y8     8P    88 `8b     d8' 88  
# I8[    ""  88  88P'   `"8a  ""     `Y8  88         88   `8b   d8' `8b   d8'    88  `8b   d8'  88  
#  `"Y8ba,   88  88       88  ,adPPPPP88  88         8P    `8a a8'   `8a a8'     88   `8b d8'   88  
# aa    ]8I  88  88       88  88,    ,88  88      .a8P      `8a8'     `8a8'      88    `888'    88  
# `"YbbdP"'  88  88       88  `"8bbdP"Y8  88888888Y"'        `8'       `8'       88     `8'     88  
                                                                                                  
                                                                                                  

#  *===========================================================*/