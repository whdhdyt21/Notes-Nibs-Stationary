import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from utils.file_utils import read_excel, append_excel, write_excel

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

class InventoryView:
    def __init__(self, master):
        self.master = master
        self.master.title("Inventory - Notes & Nibs Stationery")
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

        tk.Button(self.button_bar_frame, text="Add", command=self.add_item, **button_style).pack(side="left", padx=5)
        tk.Button(self.button_bar_frame, text="Edit", command=self.edit_item, **button_style).pack(side="left", padx=5)

        delete_style = button_style.copy()
        delete_style.update({'bg': "#F44336", 'activebackground': "#D32F2F"})
        tk.Button(self.button_bar_frame, text="Delete", command=self.delete_item, **delete_style).pack(side="left", padx=5)

        back_style = button_style.copy()
        back_style.update({'bg': "#FF4081", 'activebackground': "#C2185B"})
        tk.Button(self.button_bar_frame, text="Back", command=self.back_to_dashboard, **back_style).pack(side="left", padx=5)

    def create_table(self):
        title_label = tk.Label(
            self.table_frame,
            text="Inventory Items",
            font=("Comic Sans MS", 28, "bold"),
            bg="#FFF5DE",
            fg="#E1A269"
        )
        title_label.grid(row=0, column=0, pady=10, columnspan=2, sticky="n")

        self.vertical_scrollbar = tk.Scrollbar(self.table_frame, orient="vertical")
        self.horizontal_scrollbar = tk.Scrollbar(self.table_frame, orient="horizontal")

        self.tree = ttk.Treeview(
            self.table_frame,
            columns=("Product ID", "Product Name", "Category", "Quantity", "Purchase Price", "Selling Price", "Total", "Total Price"),
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

        self.tree.heading("Product ID", text="Product ID")
        self.tree.heading("Product Name", text="Product Name")
        self.tree.heading("Category", text="Category")
        self.tree.heading("Quantity", text="Quantity")
        self.tree.heading("Purchase Price", text="Purchase Price")
        self.tree.heading("Selling Price", text="Selling Price")
        self.tree.heading("Total", text="Total")
        self.tree.heading("Total Price", text="Total Price")

        self.tree.column("Product ID", width=100, anchor="center")
        self.tree.column("Product Name", width=200, anchor="w")
        self.tree.column("Category", width=150, anchor="w")
        self.tree.column("Quantity", width=100, anchor="center")
        self.tree.column("Purchase Price", width=130, anchor="center")
        self.tree.column("Selling Price", width=130, anchor="center")
        self.tree.column("Total", width=100, anchor="center")
        self.tree.column("Total Price", width=130, anchor="center")

        self.tree.grid(row=1, column=0, sticky="nsew")
        self.vertical_scrollbar.grid(row=1, column=1, sticky="ns")
        self.horizontal_scrollbar.grid(row=2, column=0, sticky="ew")

        self.tree.tag_configure('evenrow', background="#FFF5DE")
        self.tree.tag_configure('oddrow', background="#FFEFD5")
        self.tree.tag_configure('totalrow', font=("Comic Sans MS", 14, "bold"), foreground="#FFF5DE", background="#FF2929")

        self.table_frame.grid_rowconfigure(1, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

        self.load_inventory()

    def load_inventory(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        inventory = read_excel(INVENTORY_FILE)
        total_quantity = 0
        total_purchase_price = 0
        total_selling_price = 0
        total_total = 0
        total_total_price = 0

        for index, item in enumerate(inventory):
            row_tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.tree.insert("", "end", values=(
                item["product_id"],
                item["product_name"],
                item["category"],
                item["quantity"],
                format_rupiah("purchase_price", item["purchase_price"]),
                format_rupiah("selling_price", item["selling_price"]),
                format_rupiah("total", item["total"]),
                format_rupiah("total_price", item["total_price"])
            ), tags=(row_tag,))

            total_quantity += item["quantity"]
            total_purchase_price += item["total"]
            total_selling_price += item["total_price"]
            total_total += item["quantity"] * item["purchase_price"]
            total_total_price += item["quantity"] * item["selling_price"]

        self.tree.insert("", "end", values=(
            "Total",
            "",
            "",
            total_quantity,
            "",
            "",
            format_rupiah("total", total_total),
            format_rupiah("total_price", total_total_price)
        ), tags=('totalrow',))


    def add_item(self):
        AddItemWindow(self)

    def edit_item(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select an item to edit.")
            return

        product_id = self.tree.item(selected_item[0])['values'][0]
        inventory = read_excel(INVENTORY_FILE)
        item_to_edit = next((item for item in inventory if item["product_id"] == product_id), None)

        if not item_to_edit:
            messagebox.showerror("Error", "Item not found.")
            return

        EditItemWindow(self, item_to_edit, inventory)

    def delete_item(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select an item to delete.")
            return

        confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected item?")
        if not confirm:
            return

        product_id = self.tree.item(selected_item[0])['values'][0]
        inventory = read_excel(INVENTORY_FILE)
        updated_inventory = [item for item in inventory if item["product_id"] != product_id]

        if len(updated_inventory) == len(inventory):
            messagebox.showerror("Error", "Item not found.")
            return

        write_excel(INVENTORY_FILE, updated_inventory,
                    ["product_id", "product_name", "category", "quantity", "purchase_price", "selling_price",
                     "total", "total_price"])

        messagebox.showinfo("Success", "Item deleted successfully!")
        self.load_inventory()

    def back_to_dashboard(self):
        self.master.destroy()
        from views.dashboard_view import DashboardView  # Impor lokal
        dashboard_window = tk.Tk()
        DashboardView(dashboard_window)
        dashboard_window.mainloop()


class AddItemWindow:
    def __init__(self, parent):
        self.parent = parent
        self.top = tk.Toplevel(parent.master)
        self.top.title("Add New Inventory Item")
        self.top.geometry("400x600")
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
            ("Quantity", "quantity_entry"),
            ("Purchase Price", "purchase_price_entry"),
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
            quantity = int(self.entries["quantity_entry"].get().strip())
            purchase_price = float(self.entries["purchase_price_entry"].get().strip())
            selling_price = float(self.entries["selling_price_entry"].get().strip())
            total = quantity * purchase_price
            total_price = quantity * selling_price

            if not all([product_name, category]):
                messagebox.showerror("Error", "Please fill in all required fields.")
                return

            inventory = read_excel(INVENTORY_FILE)
            # Cek apakah produk sudah ada
            existing_product = next((item for item in inventory if item["product_name"].lower() == product_name.lower()), None)
            if existing_product:
                messagebox.showerror("Error", f"Product '{product_name}' already exists in inventory.")
                return

            new_product_id = max([item["product_id"] for item in inventory], default=0) + 1

            inventory_data = {
                "product_id": new_product_id,
                "product_name": product_name,
                "category": category,
                "quantity": quantity,
                "purchase_price": purchase_price,
                "selling_price": selling_price,
                "total": total,
                "total_price": total_price
            }
            headers = ["product_id", "product_name", "category", "quantity", "purchase_price", "selling_price", "total", "total_price"]

            append_excel(INVENTORY_FILE, inventory_data, headers)
            self.parent.load_inventory()
            messagebox.showinfo("Success", "Item added successfully!")
            self.top.destroy()
        except ValueError:
            messagebox.showerror("Error", "Please enter valid data types.")

class EditItemWindow:
    def __init__(self, parent, item_to_edit, inventory):
        self.parent = parent
        self.item_to_edit = item_to_edit
        self.inventory = inventory
        self.top = tk.Toplevel(parent.master)
        self.top.title("Edit Inventory Item")
        self.top.geometry("400x600")
        self.top.configure(bg="#FFF5DE")
        self.top.grid_rowconfigure(0, weight=0)
        self.top.grid_rowconfigure(1, weight=1)
        self.top.grid_columnconfigure(0, weight=1)
        self.top.grab_set()

        header = tk.Label(self.top, text="Edit Item", font=("Comic Sans MS", 20, "bold"), bg="#FFF5DE", fg="#E1A269")
        header.pack(pady=10)

        fields = [
            ("Product Name", "product_name_entry", self.item_to_edit["product_name"]),
            ("Category", "category_entry", self.item_to_edit["category"]),
            ("Quantity", "quantity_entry", self.item_to_edit["quantity"]),
            ("Purchase Price", "purchase_price_entry", self.item_to_edit["purchase_price"]),
            ("Selling Price", "selling_price_entry", self.item_to_edit["selling_price"]),
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
            text="Save",
            command=lambda: self.save_edited_item(),
            bg="#FFEB3B",
            fg="#E1A269",
            font=("Comic Sans MS", 12, "bold"),
            activebackground="#FFD700",
            activeforeground="#E1A269",
            cursor="hand2",
            width=10
        )
        save_button.pack(pady=20)

    def save_edited_item(self):
        try:
            product_name = self.edit_entries["product_name_entry"].get().strip()
            category = self.edit_entries["category_entry"].get().strip()
            quantity = int(self.edit_entries["quantity_entry"].get().strip())
            purchase_price = float(self.edit_entries["purchase_price_entry"].get().strip())
            selling_price = float(self.edit_entries["selling_price_entry"].get().strip())
            total = quantity * purchase_price
            total_price = quantity * selling_price

            if not all([product_name, category]):
                messagebox.showerror("Error", "Please fill in all required fields.")
                return

            existing_product = next((item for item in self.inventory if item["product_name"].lower() == product_name.lower() and item["product_id"] != self.item_to_edit["product_id"]), None)
            if existing_product:
                messagebox.showerror("Error", f"Another product with name '{product_name}' already exists.")
                return

            self.item_to_edit["product_name"] = product_name
            self.item_to_edit["category"] = category
            self.item_to_edit["quantity"] = quantity
            self.item_to_edit["purchase_price"] = purchase_price
            self.item_to_edit["selling_price"] = selling_price
            self.item_to_edit["total"] = total
            self.item_to_edit["total_price"] = total_price
            write_excel(INVENTORY_FILE, self.inventory, ["product_id", "product_name", "category", "quantity", "purchase_price", "selling_price", "total", "total_price"])
            messagebox.showinfo("Success", "Item updated successfully!")
            self.parent.load_inventory()
            self.top.destroy()
        except ValueError:
            messagebox.showerror("Error", "Please enter valid data types.")

if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryView(root)
    root.mainloop()