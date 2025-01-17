import tkinter as tk
from tkinter import messagebox, ttk
from utils.file_utils import read_excel, append_excel, write_excel, log_journal
from datetime import datetime

SALES_FILE = "database/sales.xlsx"
INVENTORY_FILE = "database/inventory_items.xlsx"

def format_rupiah(value, amount):
    try:
        amount = float(amount)
        if amount.is_integer():
            amount = int(amount)
            return "Rp. {:,}".format(amount).replace(",", ".")
        else:
            return "Rp. {:,.2f}".format(amount).replace(",", ".")
    except (ValueError, TypeError):
        return amount

class SalesView:
    def __init__(self, master):
        self.master = master
        self.master.title("Sales - Notes & Nibs Stationery")
        self.master.geometry("1000x600")
        self.master.configure(bg="#FFF5DE")
        self.master.minsize(1000, 600)

        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)

        self.frame = tk.Frame(self.master, bg="#FFF5DE", padx=20, pady=20)
        self.frame.grid(row=0, column=0, sticky="nsew")

        self.frame.grid_rowconfigure(0, weight=1)
        self.frame.grid_rowconfigure(1, weight=0)
        self.frame.grid_columnconfigure(0, weight=1)

        self.table_frame = tk.Frame(self.frame, bg="#FFF5DE", padx=10, pady=10)
        self.table_frame.grid(row=0, column=0, sticky="nsew")

        self.table_frame.grid_rowconfigure(1, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

        self.button_bar_frame = tk.Frame(self.frame, bg="#FFF5DE", pady=10)
        self.button_bar_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))

        self.create_table()
        self.create_button_bar()

    def create_button_bar(self):
        button_style = {
            'font': ("Comic Sans MS", 12, "bold"),
            'width': 8,
            'relief': "flat",
            'bd': 2,
            'pady': 5,
            'fg': "#E1A269",
            'bg': "#FFEB3B",
            'activebackground': "#FFD700",
            'activeforeground': "#E1A269",
            'cursor': "hand2"
        }

        tk.Button(self.button_bar_frame, text="Add", command=self.add_sale, **button_style).pack(side="left", padx=5)
        tk.Button(self.button_bar_frame, text="Edit", command=self.edit_sale, **button_style).pack(side="left", padx=5)

        delete_style = button_style.copy()
        delete_style.update({'bg': "#F44336", 'activebackground': "#D32F2F"})
        tk.Button(self.button_bar_frame, text="Delete", command=self.delete_sale, **delete_style).pack(side="left", padx=5)

        back_style = button_style.copy()
        back_style.update({'bg': "#FF4081", 'activebackground': "#C2185B"})
        tk.Button(self.button_bar_frame, text="Back", command=self.back_to_dashboard, **back_style).pack(side="left", padx=5)

    def create_table(self):
        title_label = tk.Label(
            self.table_frame,
            text="Sales",
            font=("Comic Sans MS", 28, "bold"),
            bg="#FFF5DE",
            fg="#E1A269"
        )
        title_label.grid(row=0, column=0, pady=10, columnspan=2, sticky="n")

        self.vertical_scrollbar = tk.Scrollbar(self.table_frame, orient="vertical")
        self.horizontal_scrollbar = tk.Scrollbar(self.table_frame, orient="horizontal")

        self.tree = ttk.Treeview(
            self.table_frame,
            columns=("Sale ID", "Date", "Product Name", "Quantity Sold", "Selling Price", "Total Income", "COGS", "Gross Profit"),
            show="headings",
            yscrollcommand=self.vertical_scrollbar.set,
            xscrollcommand=self.horizontal_scrollbar.set
        )

        self.vertical_scrollbar.config(command=self.tree.yview)
        self.horizontal_scrollbar.config(command=self.tree.xview)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=("Comic Sans MS", 14, "bold"), foreground="#FFF5DE", background="#E1A269")
        style.configure("Treeview", font=("Comic Sans MS", 14), rowheight=30, background="#FFF5DE", foreground="#333333", fieldbackground="#FFF5DE")
        style.map("Treeview", background=[("selected", "#E1A269")], foreground=[("selected", "#FFF5DE")])

        self.tree.heading("Sale ID", text="Sale ID")
        self.tree.heading("Date", text="Date")
        self.tree.heading("Product Name", text="Product Name")
        self.tree.heading("Quantity Sold", text="Quantity Sold")
        self.tree.heading("Selling Price", text="Selling Price")
        self.tree.heading("Total Income", text="Total Income")
        self.tree.heading("COGS", text="COGS")
        self.tree.heading("Gross Profit", text="Gross Profit")

        self.tree.column("Sale ID", width=100, anchor="center")
        self.tree.column("Date", width=120, anchor="center")
        self.tree.column("Product Name", width=200, anchor="w")
        self.tree.column("Quantity Sold", width=150, anchor="center")
        self.tree.column("Selling Price", width=130, anchor="center")
        self.tree.column("Total Income", width=130, anchor="center")
        self.tree.column("COGS", width=100, anchor="center")
        self.tree.column("Gross Profit", width=130, anchor="center")

        self.tree.grid(row=1, column=0, sticky="nsew")
        self.vertical_scrollbar.grid(row=1, column=1, sticky="ns")
        self.horizontal_scrollbar.grid(row=2, column=0, sticky="ew")

        self.tree.tag_configure('evenrow', background="#FFF5DE")
        self.tree.tag_configure('oddrow', background="#FFEFD5")
        self.tree.tag_configure('totalrow', font=("Comic Sans MS", 14, "bold"), foreground="#FFF5DE", background="#FF2929")

        self.table_frame.grid_rowconfigure(1, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

        self.load_sales()

    def load_sales(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        sales = read_excel(SALES_FILE)
        total_quantity_sold = 0
        total_income = 0
        total_cogs = 0
        total_gross_profit = 0

        for index, sale in enumerate(sales):
            row_tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.tree.insert("", "end", values=(
                sale["sale_id"],
                sale["date"],
                sale["product_name"],
                sale["quantity_sold"],
                format_rupiah("Selling Price", sale["selling_price"]),
                format_rupiah("Total Income", sale["total_income"]),
                format_rupiah("COGS", sale["cogs"]),
                format_rupiah("Gross Profit", sale["gross_profit"])
            ), tags=(row_tag,))

            total_quantity_sold += sale["quantity_sold"]
            total_income += sale["total_income"]
            total_cogs += sale["cogs"]
            total_gross_profit += sale["gross_profit"]

        # Menambahkan baris total di bawah tabel
        self.tree.insert("", "end", values=(
            "Total",
            "",
            "",
            total_quantity_sold,
            "",
            format_rupiah("Total Income", total_income),
            format_rupiah("COGS", total_cogs),
            format_rupiah("Gross Profit", total_gross_profit)
        ), tags=('totalrow',))


    def add_sale(self):
        AddSaleWindow(self)

    def edit_sale(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a sale to edit.")
            return

        sale_id = self.tree.item(selected_item[0])['values'][0]
        sales = read_excel(SALES_FILE)
        sale_to_edit = next((sale for sale in sales if sale["sale_id"] == sale_id), None)

        if not sale_to_edit:
            messagebox.showerror("Error", "Sale not found.")
            return

        EditSaleWindow(self, sale_to_edit, sales)

    def delete_sale(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a sale to delete.")
            return

        confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected sale?")
        if not confirm:
            return

        sale_id = self.tree.item(selected_item[0])['values'][0]
        sales = read_excel(SALES_FILE)
        updated_sales = [sale for sale in sales if sale["sale_id"] != sale_id]

        if len(updated_sales) == len(sales):
            messagebox.showerror("Error", "Sale not found.")
            return

        sale_to_delete = next((sale for sale in sales if sale["sale_id"] == sale_id), None)
        if sale_to_delete:
            inventory = read_excel(INVENTORY_FILE)
            product = next((item for item in inventory if item["product_name"].lower() == sale_to_delete["product_name"].lower()), None)
            if product:
                product["quantity"] += sale_to_delete["quantity_sold"]
                product["total"] = product["quantity"] * product["purchase_price"]
                product["total_price"] = product["quantity"] * product["selling_price"]
                write_excel(INVENTORY_FILE, inventory, ["product_id", "date", "product_name", "category", "quantity", "purchase_price", "selling_price", "total", "total_price"])

        write_excel(SALES_FILE, updated_sales, ["sale_id", "date", "product_name", "quantity_sold", "selling_price", "total_income", "cogs", "gross_profit"])
        self.load_sales()
        messagebox.showinfo("Success", "Sale deleted successfully!")

    def back_to_dashboard(self):
        self.master.destroy()
        from views.dashboard_view import DashboardView
        dashboard_window = tk.Tk()
        DashboardView(dashboard_window)
        dashboard_window.mainloop()

class AddSaleWindow:
    def __init__(self, parent):
        self.parent = parent
        self.top = tk.Toplevel(parent.master)
        self.top.title("Add New Sale")
        self.top.geometry("400x600")
        self.top.configure(bg="#FFF5DE")
        self.top.grid_rowconfigure(0, weight=0)
        self.top.grid_rowconfigure(1, weight=1)
        self.top.grid_columnconfigure(0, weight=1)
        self.top.grab_set()

        header = tk.Label(self.top, text="Add New Sale", font=("Comic Sans MS", 20, "bold"), bg="#FFF5DE", fg="#E1A269")
        header.pack(pady=10)

        fields = [
            ("Date (YYYY-MM-DD)", "date_entry"),
            ("Product Name", "product_name_entry"),
            ("Quantity Sold", "quantity_sold_entry"),
            ("Selling Price", "selling_price_entry"),
        ]

        self.entries = {}

        for label_text, var_name in fields:
            label = tk.Label(self.top, text=label_text, font=("Comic Sans MS", 14), bg="#FFF5DE", fg="#E1A269")
            label.pack(pady=(10, 5))
            entry = tk.Entry(self.top, font=("Comic Sans MS", 14), width=30)
            entry.pack(pady=5)
            self.entries[var_name] = entry

        add_button = tk.Button(
            self.top,
            text="Add Sale",
            command=self.save_new_sale,
            bg="#FFEB3B",
            fg="#E1A269",
            font=("Comic Sans MS", 12, "bold"),
            activebackground="#FFD700",
            activeforeground="#E1A269",
            cursor="hand2",
            width=10
        )
        add_button.pack(pady=20)

    def save_new_sale(self):
        try:
            date = datetime.now().strftime("%Y-%m-%d")
            product_name = self.entries["product_name_entry"].get().strip()
            quantity_sold = int(self.entries["quantity_sold_entry"].get().strip())
            selling_price = float(self.entries["selling_price_entry"].get().strip())

            if not all([date, product_name]):
                messagebox.showerror("Error", "Please fill in all required fields.")
                return

            try:
                datetime.strptime(date, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Error", "Please enter the date in YYYY-MM-DD format.")
                return

            inventory = read_excel(INVENTORY_FILE)
            product = next((item for item in inventory if item["product_name"].lower() == product_name.lower()), None)

            if not product:
                messagebox.showerror("Error", f"Product '{product_name}' not found in inventory.")
                return

            if product["quantity"] < quantity_sold:
                messagebox.showerror("Error", f"Insufficient stock for product '{product_name}'. Available quantity: {product['quantity']}.")
                return

            product["quantity"] -= quantity_sold
            product["total"] = product["quantity"] * product["purchase_price"]
            product["total_price"] = product["quantity"] * product["selling_price"]

            write_excel(INVENTORY_FILE, inventory, ["product_id", "date", "product_name", "category", "quantity", "purchase_price", "selling_price", "total", "total_price"])

            cogs = quantity_sold * product["purchase_price"]
            total_income = quantity_sold * selling_price
            gross_profit = total_income - cogs

            sales = read_excel(SALES_FILE)
            new_sale_id = max([sale["sale_id"] for sale in sales], default=0) + 1

            sale_data = {
                "sale_id": new_sale_id,
                "date": date,
                "product_name": product_name,
                "quantity_sold": quantity_sold,
                "selling_price": selling_price,
                "total_income": total_income,
                "cogs": cogs,
                "gross_profit": gross_profit
            }
            headers = ["sale_id", "date", "product_name", "quantity_sold", "selling_price", "total_income", "cogs", "gross_profit"]

            append_excel(SALES_FILE, sale_data, headers)

            journal_entry = {
                "date": date,
                "transaction_type": "Sale",
                "product_name": product_name,
                "quantity": quantity_sold,
                "unit_price": selling_price,
                "total_amount": total_income
            }
            log_journal(journal_entry)

            self.parent.load_sales()
            messagebox.showinfo("Success", "Sale added successfully!")
            self.top.destroy()
        except ValueError:
            messagebox.showerror("Error", "Please enter valid data types.")

class EditSaleWindow:
    def __init__(self, parent, sale_to_edit, sales):
        self.parent = parent
        self.sale_to_edit = sale_to_edit
        self.sales = sales
        self.top = tk.Toplevel(parent.master)
        self.top.title("Edit Sale")
        self.top.geometry("400x600")
        self.top.configure(bg="#FFF5DE")
        self.top.grid_rowconfigure(0, weight=0)
        self.top.grid_rowconfigure(1, weight=1)
        self.top.grid_columnconfigure(0, weight=1)
        self.top.grab_set()

        header = tk.Label(self.top, text="Edit Sale", font=("Comic Sans MS", 20, "bold"), bg="#FFF5DE", fg="#E1A269")
        header.pack(pady=10)

        fields = [
            ("Date (YYYY-MM-DD)", "date_entry", self.sale_to_edit["date"]),
            ("Product Name", "product_name_entry", self.sale_to_edit["product_name"]),
            ("Quantity Sold", "quantity_sold_entry", self.sale_to_edit["quantity_sold"]),
            ("Selling Price", "selling_price_entry", self.sale_to_edit["selling_price"]),
        ]

        self.edit_entries = {}

        for label_text, var_name, value in fields:
            label = tk.Label(self.top, text=label_text, font=("Comic Sans MS", 14), bg="#FFF5DE", fg="#E1A269")
            label.pack(pady=(10, 5))
            entry = tk.Entry(self.top, font=("Comic Sans MS", 14), width=30)
            entry.insert(0, value)
            entry.pack(pady=5)
            self.edit_entries[var_name] = entry

        save_button = tk.Button(
            self.top,
            text="Save Sale",
            command=self.save_edited_sale,
            bg="#FFEB3B",
            fg="#E1A269",
            font=("Comic Sans MS", 12, "bold"),
            activebackground="#FFD700",
            activeforeground="#E1A269",
            cursor="hand2",
            width=10
        )
        save_button.pack(pady=20)

    def save_edited_sale(self):
        try:
            date = self.edit_entries["date_entry"].get().strip()
            product_name = self.edit_entries["product_name_entry"].get().strip()
            quantity_sold = int(self.edit_entries["quantity_sold_entry"].get().strip())
            selling_price = float(self.edit_entries["selling_price_entry"].get().strip())

            if not all([date, product_name]):
                messagebox.showerror("Error", "Please fill in all required fields.")
                return
            try:
                datetime.strptime(date, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Error", "Please enter the date in YYYY-MM-DD format.")
                return

            inventory = read_excel(INVENTORY_FILE)
            product = next((item for item in inventory if item["product_name"].lower() == product_name.lower()), None)

            if not product:
                messagebox.showerror("Error", f"Product '{product_name}' not found in inventory.")
                return

            original_quantity = self.sale_to_edit["quantity_sold"]
            original_selling_price = self.sale_to_edit["selling_price"]

            quantity_diff = quantity_sold - original_quantity
            selling_price_diff = selling_price - original_selling_price

            if quantity_diff > 0:
                if product["quantity"] < quantity_diff:
                    messagebox.showerror("Error", f"Insufficient stock for product '{product_name}'. Available quantity: {product['quantity']}.")
                    return
                product["quantity"] -= quantity_diff
            elif quantity_diff < 0:
                product["quantity"] += abs(quantity_diff)

            product["total"] = product["quantity"] * product["purchase_price"]
            product["total_price"] = product["quantity"] * product["selling_price"]

            write_excel(INVENTORY_FILE, inventory, ["product_id", "date", "product_name", "category", "quantity", "purchase_price", "selling_price", "total", "total_price"])

            cogs = quantity_sold * product["purchase_price"]
            total_income = quantity_sold * selling_price
            gross_profit = total_income - cogs

            self.sale_to_edit["date"] = date
            self.sale_to_edit["product_name"] = product_name
            self.sale_to_edit["quantity_sold"] = quantity_sold
            self.sale_to_edit["selling_price"] = selling_price
            self.sale_to_edit["total_income"] = total_income
            self.sale_to_edit["cogs"] = cogs
            self.sale_to_edit["gross_profit"] = gross_profit

            write_excel(SALES_FILE, self.sales, ["sale_id", "date", "product_name", "quantity_sold", "selling_price", "total_income", "cogs", "gross_profit"])

            journal_entry = {
                "date": date,
                "transaction_type": "Sale Edit",
                "product_name": product_name,
                "quantity": quantity_diff,
                "unit_price": selling_price_diff,
                "total_amount": quantity_diff * selling_price_diff
            }
            log_journal(journal_entry)

            messagebox.showinfo("Success", "Sale updated successfully!")
            self.parent.load_sales()
            self.top.destroy()
        except ValueError:
            messagebox.showerror("Error", "Please enter valid data types.")

if __name__ == "__main__":
    root = tk.Tk()
    SalesView(root)
    root.mainloop()