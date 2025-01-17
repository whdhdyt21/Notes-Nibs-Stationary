import tkinter as tk
from tkinter import messagebox, ttk
from utils.file_utils import read_excel, append_excel, write_excel, log_journal
from datetime import datetime

PURCHASES_FILE = "database/purchases.xlsx"
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

class PurchaseView:
    def __init__(self, master):
        self.master = master
        self.master.title("Purchases - Notes & Nibs Stationery")
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
            'width': 10,
            'relief': "flat",
            'bd': 2,
            'pady': 5,
            'fg': "#E1A269",
            'bg': "#FFEB3B",
            'activebackground': "#FFD700",
            'activeforeground': "#E1A269",
            'cursor': "hand2"
        }

        tk.Button(self.button_bar_frame, text="Add", command=self.add_purchase, **button_style).pack(side="left", padx=5)
        tk.Button(self.button_bar_frame, text="Edit", command=self.edit_purchase, **button_style).pack(side="left", padx=5)

        delete_style = button_style.copy()
        delete_style.update({'bg': "#F44336", 'activebackground': "#D32F2F"})
        tk.Button(self.button_bar_frame, text="Delete", command=self.delete_purchase, **delete_style).pack(side="left", padx=5)

        back_style = button_style.copy()
        back_style.update({'bg': "#FF4081", 'activebackground': "#C2185B"})
        tk.Button(self.button_bar_frame, text="Back", command=self.back_to_dashboard, **back_style).pack(side="left", padx=5)

    def create_table(self):
        title_label = tk.Label(
            self.table_frame,
            text="Purchases",
            font=("Comic Sans MS", 28, "bold"),
            bg="#FFF5DE",
            fg="#E1A269"
        )
        title_label.grid(row=0, column=0, pady=10, columnspan=2, sticky="n")

        self.vertical_scrollbar = tk.Scrollbar(self.table_frame, orient="vertical")
        self.horizontal_scrollbar = tk.Scrollbar(self.table_frame, orient="horizontal")

        self.tree = ttk.Treeview(
            self.table_frame,
            columns=("Purchase ID", "Date", "Product Name", "Purchase Price", "Quantity", "Total"),
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

        self.tree.heading("Purchase ID", text="Purchase ID")
        self.tree.heading("Date", text="Date")
        self.tree.heading("Product Name", text="Product Name")
        self.tree.heading("Purchase Price", text="Purchase Price")
        self.tree.heading("Quantity", text="Quantity")
        self.tree.heading("Total", text="Total")

        self.tree.column("Purchase ID", width=100, anchor="center")
        self.tree.column("Date", width=120, anchor="center")
        self.tree.column("Product Name", width=200, anchor="w")
        self.tree.column("Purchase Price", width=150, anchor="center")
        self.tree.column("Quantity", width=100, anchor="center")
        self.tree.column("Total", width=100, anchor="center")

        self.tree.grid(row=1, column=0, sticky="nsew")
        self.vertical_scrollbar.grid(row=1, column=1, sticky="ns")
        self.horizontal_scrollbar.grid(row=2, column=0, sticky="ew")

        self.tree.tag_configure('evenrow', background="#FFF5DE")
        self.tree.tag_configure('oddrow', background="#FFEFD5")
        self.tree.tag_configure('totalrow', font=("Comic Sans MS", 14, "bold"), foreground="#FFF5DE", background="#FF2929")

        self.table_frame.grid_rowconfigure(1, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

        self.load_purchases()

    def load_purchases(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        purchases = read_excel(PURCHASES_FILE)
        total_quantity = 0
        total_purchase_price = 0
        total = 0

        for index, purchase in enumerate(purchases):
            row_tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.tree.insert("", "end", values=(
                purchase.get("purchase_id", ""),
                purchase.get("date", ""),
                purchase.get("product_name", ""),
                format_rupiah("Purchase Price", purchase.get("purchase_price", "")),
                purchase.get("quantity", ""),
                format_rupiah("Total", purchase.get("total", "")),
            ), tags=(row_tag,))

            total_quantity += purchase.get("quantity", 0)
            total_purchase_price += purchase.get("purchase_price", 0) * purchase.get("quantity", 0)
            total += purchase.get("total", 0)

        # Menambahkan baris total di bawah tabel
        self.tree.insert("", "end", values=(
            "Total",
            "",
            "",
            "",
            total_quantity,
            format_rupiah("Total", total)
        ), tags=('totalrow',))

    def add_purchase(self):
        AddPurchaseWindow(self)

    def edit_purchase(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a purchase to edit.")
            return

        purchase_id = self.tree.item(selected_item[0])['values'][0]
        purchases = read_excel(PURCHASES_FILE)
        purchase_to_edit = next((purchase for purchase in purchases if purchase.get("purchase_id") == purchase_id), None)

        if not purchase_to_edit:
            messagebox.showerror("Error", "Purchase not found.")
            return

        EditPurchaseWindow(self, purchase_to_edit, purchases)

    def delete_purchase(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a purchase to delete.")
            return

        confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected purchase?")
        if not confirm:
            return

        purchase_id = self.tree.item(selected_item[0])['values'][0]
        purchases = read_excel(PURCHASES_FILE)
        updated_purchases = [purchase for purchase in purchases if purchase.get("purchase_id") != purchase_id]

        if len(updated_purchases) == len(purchases):
            messagebox.showerror("Error", "Purchase not found.")
            return

        write_excel(PURCHASES_FILE, updated_purchases, ["purchase_id", "date", "product_name", "purchase_price", "quantity", "total", "total_price"])

        messagebox.showinfo("Success", "Purchase deleted successfully!")
        self.load_purchases()

    def back_to_dashboard(self):
        self.master.destroy()
        from views.dashboard_view import DashboardView  # Impor lokal
        dashboard_window = tk.Tk()
        DashboardView(dashboard_window)
        dashboard_window.mainloop()

class AddPurchaseWindow:
    def __init__(self, parent):
        self.parent = parent
        self.top = tk.Toplevel(parent.master)
        self.top.title("Add New Purchase")
        self.top.geometry("400x600")
        self.top.configure(bg="#FFF5DE")
        self.top.grid_rowconfigure(0, weight=0)
        self.top.grid_rowconfigure(1, weight=1)
        self.top.grid_columnconfigure(0, weight=1)
        self.top.grab_set()

        header = tk.Label(self.top, text="Add New Purchase", font=("Comic Sans MS", 20, "bold"), bg="#FFF5DE", fg="#E1A269")
        header.pack(pady=10)

        fields = [
            ("Product Name", "product_name_entry"),
            ("Purchase Price", "purchase_price_entry"),
            ("Quantity", "quantity_entry"),
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
            text="Add Purchase",
            command=self.save_new_purchase,
            bg="#FFEB3B",
            fg="#E1A269",
            font=("Comic Sans MS", 12, "bold"),
            activebackground="#FFD700",
            activeforeground="#E1A269",
            cursor="hand2",
            width=12
        )
        add_button.pack(pady=20)

    def save_new_purchase(self):
        try:
            date = datetime.now().strftime("%Y-%m-%d")
            product_name = self.entries["product_name_entry"].get().strip()
            purchase_price = float(self.entries["purchase_price_entry"].get().strip())
            quantity = int(self.entries["quantity_entry"].get().strip())

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
                response = messagebox.askyesno("Product Not Found", f"Product '{product_name}' not found in inventory. Do you want to add it?")
                if response:
                    AddItemWindow(self.parent)
                return

            product["quantity"] += quantity
            product["total"] = product["quantity"] * product["purchase_price"]
            product["total_price"] = product["quantity"] * product["selling_price"]
            write_excel(INVENTORY_FILE, inventory, ["product_id", "date", "product_name", "category", "quantity", "purchase_price", "selling_price", "total", "total_price"])

            total = quantity * purchase_price
            total_price = quantity * product["selling_price"]

            purchases = read_excel(PURCHASES_FILE)
            new_purchase_id = max([purchase.get("purchase_id", 0) for purchase in purchases], default=0) + 1

            purchase_data = {
                "purchase_id": new_purchase_id,
                "date": date,
                "product_name": product_name,
                "purchase_price": purchase_price,
                "quantity": quantity,
                "total": total,
                "total_price": total_price
            }
            headers = ["purchase_id", "date", "product_name", "purchase_price", "quantity", "total", "total_price"]

            append_excel(PURCHASES_FILE, purchase_data, headers)
            journal_entry = {
                "date": date,
                "transaction_type": "Purchase",
                "product_name": product_name,
                "quantity": quantity,
                "unit_price": purchase_price,
                "total_amount": total
            }
            log_journal(journal_entry)

            self.parent.load_purchases()
            messagebox.showinfo("Success", "Purchase added successfully!")
            self.top.destroy()
        except ValueError:
            messagebox.showerror("Error", "Please enter valid data types.")

class EditPurchaseWindow:
    def __init__(self, parent, purchase_to_edit, purchases):
        self.parent = parent
        self.purchase_to_edit = purchase_to_edit
        self.purchases = purchases
        self.top = tk.Toplevel(parent.master)
        self.top.title("Edit Purchase")
        self.top.geometry("400x600")
        self.top.configure(bg="#FFF5DE")
        self.top.grid_rowconfigure(0, weight=0)
        self.top.grid_rowconfigure(1, weight=1)
        self.top.grid_columnconfigure(0, weight=1)
        self.top.grab_set()

        header = tk.Label(self.top, text="Edit Purchase", font=("Comic Sans MS", 20, "bold"), bg="#FFF5DE", fg="#E1A269")
        header.pack(pady=10)

        fields = [
            ("Date (YYYY-MM-DD)", "date_entry", self.purchase_to_edit.get("date", "")),
            ("Product Name", "product_name_entry", self.purchase_to_edit.get("product_name", "")),
            ("Purchase Price", "purchase_price_entry", self.purchase_to_edit.get("purchase_price", "")),
            ("Quantity", "quantity_entry", self.purchase_to_edit.get("quantity", "")),
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
            text="Save Purchase",
            command=self.save_edited_purchase,
            bg="#FFEB3B",
            fg="#E1A269",
            font=("Comic Sans MS", 12, "bold"),
            activebackground="#FFD700",
            activeforeground="#E1A269",
            cursor="hand2",
            width=12
        )
        save_button.pack(pady=20)

    def save_edited_purchase(self):
        try:
            date = self.edit_entries["date_entry"].get().strip()
            product_name = self.edit_entries["product_name_entry"].get().strip()
            purchase_price = float(self.edit_entries["purchase_price_entry"].get().strip())
            quantity = int(self.edit_entries["quantity_entry"].get().strip())

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

            original_quantity = self.purchase_to_edit.get("quantity", 0)
            original_purchase_price = self.purchase_to_edit.get("purchase_price", 0.0)

            quantity_diff = quantity - original_quantity
            purchase_price_diff = purchase_price - original_purchase_price

            product["quantity"] += quantity_diff
            product["total"] = product["quantity"] * product["purchase_price"]
            product["total_price"] = product["quantity"] * product["selling_price"]
            write_excel(INVENTORY_FILE, inventory, ["product_id", "date", "product_name", "category", "quantity", "purchase_price", "selling_price", "total", "total_price"])
            total = quantity * purchase_price
            total_price = quantity * product["selling_price"]

            self.purchase_to_edit["date"] = date
            self.purchase_to_edit["product_name"] = product_name
            self.purchase_to_edit["purchase_price"] = purchase_price
            self.purchase_to_edit["quantity"] = quantity
            self.purchase_to_edit["total"] = total
            self.purchase_to_edit["total_price"] = total_price
            write_excel(PURCHASES_FILE, self.purchases, ["purchase_id", "date", "product_name", "purchase_price", "quantity", "total", "total_price"])

            journal_entry = {
                "date": date,
                "transaction_type": "Purchase Edit",
                "product_name": product_name,
                "quantity": quantity_diff,
                "unit_price": purchase_price_diff,
                "total_amount": quantity_diff * purchase_price_diff
            }
            log_journal(journal_entry)

            messagebox.showinfo("Success", "Purchase updated successfully!")
            self.parent.load_purchases()
            self.top.destroy()
        except ValueError:
            messagebox.showerror("Error", "Please enter valid data types.")

class AddItemWindow:
    def __init__(self, parent):
        self.parent = parent
        self.top = tk.Toplevel(parent.master)
        self.top.title("Add New Inventory Item")
        self.top.geometry("400x500")
        self.top.configure(bg="#FFF5DE")
        self.top.grid_rowconfigure(0, weight=0)
        self.top.grid_rowconfigure(1, weight=1)
        self.top.grid_columnconfigure(0, weight=1)
        self.top.grab_set()

        header = tk.Label(self.top, text="Add New Item", font=("Comic Sans MS", 20, "bold"), bg="#FFF5DE", fg="#E1A269")
        header.pack(pady=10)

        fields = [
            ("Product Name", "product_name_entry"),
            ("Category", "category_entry"),
            ("Purchase Price", "purchase_price_entry"),
            ("Selling Price", "selling_price_entry"),
            ("Quantity", "quantity_entry"),
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
            text="Add Item",
            command=self.save_new_item,
            bg="#FFEB3B",
            fg="#E1A269",
            font=("Comic Sans MS", 12, "bold"),
            activebackground="#FFD700",
            activeforeground="#E1A269",
            cursor="hand2",
            width=10
        )
        add_button.pack(pady=20)

    def save_new_item(self):
        try:
            product_name = self.entries["product_name_entry"].get().strip()
            category = self.entries["category_entry"].get().strip()
            purchase_price = float(self.entries["purchase_price_entry"].get().strip())
            selling_price = float(self.entries["selling_price_entry"].get().strip())
            quantity = int(self.entries["quantity_entry"].get().strip())
            total = quantity * purchase_price
            total_price = quantity * selling_price

            if not all([product_name, category]):
                messagebox.showerror("Error", "Please fill in all required fields.")
                return

            inventory = read_excel(INVENTORY_FILE)
            existing_product = next((item for item in inventory if item["product_name"].lower() == product_name.lower()), None)
            if existing_product:
                messagebox.showerror("Error", f"Product '{product_name}' already exists in inventory.")
                return

            new_product_id = max([item.get("product_id", 0) for item in inventory], default=0) + 1

            inventory_data = {
                "product_id": new_product_id,
                "date": datetime.now().strftime("%Y-%m-%d"),
                "product_name": product_name,
                "category": category,
                "quantity": quantity,
                "purchase_price": purchase_price,
                "selling_price": selling_price,
                "total": total,
                "total_price": total_price
            }
            headers = ["product_id", "date", "product_name", "category", "quantity", "purchase_price", "selling_price", "total", "total_price"]

            append_excel(INVENTORY_FILE, inventory_data, headers)
            self.parent.load_purchases()
            messagebox.showinfo("Success", "Item added successfully!")
            self.top.destroy()
        except ValueError:
            messagebox.showerror("Error", "Please enter valid data types.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PurchaseView(root)
    root.mainloop()