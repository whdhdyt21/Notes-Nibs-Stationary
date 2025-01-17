import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from utils.file_utils import read_excel

JOURNAL_FILE = "database/journal_entries.xlsx"

class JournalView:
    def __init__(self, master):
        self.master = master
        self.master.title("Journal - Notes & Nibs Stationery")
        self.master.geometry("1000x800")
        self.master.configure(bg="#FFF5DE")
        self.master.minsize(1000, 800)

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

        self.create_button_bar()

        self.create_purchase_table()
        self.create_sale_table()

    def format_rupiah(self, amount):
        try:
            amount = float(amount)
            if amount.is_integer():
                amount = int(amount)
                return "Rp. {:,}".format(amount).replace(",", ".")
            else:
                return "Rp. {:,.2f}".format(amount).replace(",", ".")
        except (ValueError, TypeError):
            return amount

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

        tk.Button(self.button_bar_frame, text="Refresh", command=self.load_journal, **button_style).pack(side="left", padx=5)
        tk.Button(self.button_bar_frame, text="Back", command=self.back_to_dashboard, **button_style).pack(side="left", padx=5)

    def create_purchase_table(self):
        self.purchase_label = tk.Label(
            self.table_frame,
            text="Purchase Journal Entries",
            font=("Comic Sans MS", 18, "bold"),
            bg="#FFF5DE",
            fg="#E1A269"
        )
        self.purchase_label.grid(row=0, column=0, pady=10, columnspan=2, sticky="n")

        self.purchase_vertical_scrollbar = tk.Scrollbar(self.table_frame, orient="vertical")
        self.purchase_horizontal_scrollbar = tk.Scrollbar(self.table_frame, orient="horizontal")
        self.purchase_tree = ttk.Treeview(
            self.table_frame,
            columns=("Entry ID", "Date", "Transaction Type", "Product Name", "Quantity", "Unit Price", "Total Amount"),
            show="headings",
            yscrollcommand=self.purchase_vertical_scrollbar.set,
            xscrollcommand=self.purchase_horizontal_scrollbar.set
        )

        self.purchase_vertical_scrollbar.config(command=self.purchase_tree.yview)
        self.purchase_horizontal_scrollbar.config(command=self.purchase_tree.xview)
        self.setup_treeview_style(self.purchase_tree)

        self.purchase_tree.grid(row=1, column=0, sticky="nsew")
        self.purchase_vertical_scrollbar.grid(row=1, column=1, sticky="ns")
        self.purchase_horizontal_scrollbar.grid(row=2, column=0, sticky="ew")

        self.purchase_tree.tag_configure('evenrow', background="#FFF5DE")
        self.purchase_tree.tag_configure('oddrow', background="#FFEFD5")
        self.purchase_tree.tag_configure('totalrow', font=("Comic Sans MS", 14, "bold"), foreground="#FFF5DE", background="#FF2929")

    def create_sale_table(self):
        self.sale_label = tk.Label(
            self.table_frame,
            text="Sale Journal Entries",
            font=("Comic Sans MS", 18, "bold"),
            bg="#FFF5DE",
            fg="#E1A269"
        )
        self.sale_label.grid(row=3, column=0, pady=10, columnspan=2, sticky="n")

        self.sale_vertical_scrollbar = tk.Scrollbar(self.table_frame, orient="vertical")
        self.sale_horizontal_scrollbar = tk.Scrollbar(self.table_frame, orient="horizontal")
        self.sale_tree = ttk.Treeview(
            self.table_frame,
            columns=("Entry ID", "Date", "Transaction Type", "Product Name", "Quantity", "Unit Price", "Total Amount"),
            show="headings",
            yscrollcommand=self.sale_vertical_scrollbar.set,
            xscrollcommand=self.sale_horizontal_scrollbar.set
        )

        self.sale_vertical_scrollbar.config(command=self.sale_tree.yview)
        self.sale_horizontal_scrollbar.config(command=self.sale_tree.xview)
        self.setup_treeview_style(self.sale_tree)

        self.sale_tree.grid(row=4, column=0, sticky="nsew")
        self.sale_vertical_scrollbar.grid(row=4, column=1, sticky="ns")
        self.sale_horizontal_scrollbar.grid(row=5, column=0, sticky="ew")

        self.sale_tree.tag_configure('evenrow', background="#FFF5DE")
        self.sale_tree.tag_configure('oddrow', background="#FFEFD5")
        self.sale_tree.tag_configure('totalrow', font=("Comic Sans MS", 14, "bold"), foreground="#FFF5DE", background="#FF2929")

    def setup_treeview_style(self, treeview):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=("Comic Sans MS", 14, "bold"), foreground="#FFF5DE", background="#E1A269")
        style.configure("Treeview", font=("Comic Sans MS", 14), rowheight=30, background="#FFF5DE", foreground="#333333", fieldbackground="#FFF5DE")
        style.map("Treeview", background=[("selected", "#E1A269")], foreground=[("selected", "#FFF5DE")])

        treeview.heading("Entry ID", text="Entry ID")
        treeview.heading("Date", text="Date")
        treeview.heading("Transaction Type", text="Transaction Type")
        treeview.heading("Product Name", text="Product Name")
        treeview.heading("Quantity", text="Quantity")
        treeview.heading("Unit Price", text="Unit Price")
        treeview.heading("Total Amount", text="Total Amount")

        treeview.column("Entry ID", width=100, anchor="center")
        treeview.column("Date", width=120, anchor="center")
        treeview.column("Transaction Type", width=150, anchor="center")
        treeview.column("Product Name", width=200, anchor="w")
        treeview.column("Quantity", width=100, anchor="center")
        treeview.column("Unit Price", width=130, anchor="center")
        treeview.column("Total Amount", width=130, anchor="center")

    def load_journal(self):
        # Menghapus data lama dari kedua tabel
        for row in self.purchase_tree.get_children():
            self.purchase_tree.delete(row)
        for row in self.sale_tree.get_children():
            self.sale_tree.delete(row)

        journal_entries = read_excel(JOURNAL_FILE)

        total_purchase_quantity = 0
        total_purchase_amount = 0
        total_sale_quantity = 0
        total_sale_amount = 0

        for index, entry in enumerate(journal_entries):
            row_tag = 'evenrow' if index % 2 == 0 else 'oddrow'

            if entry["transaction_type"] == "Purchase":
                self.purchase_tree.insert("", "end", values=(
                    entry["entry_id"],
                    entry["date"],
                    entry["transaction_type"],
                    entry["product_name"],
                    entry["quantity"],
                    self.format_rupiah(entry["unit_price"]),
                    self.format_rupiah(entry["total_amount"])
                ), tags=(row_tag,))
                total_purchase_quantity += entry["quantity"]
                total_purchase_amount += entry["total_amount"]
            elif entry["transaction_type"] == "Sale":
                self.sale_tree.insert("", "end", values=(
                    entry["entry_id"],
                    entry["date"],
                    entry["transaction_type"],
                    entry["product_name"],
                    entry["quantity"],
                    self.format_rupiah(entry["unit_price"]),
                    self.format_rupiah(entry["total_amount"])
                ), tags=(row_tag,))
                total_sale_quantity += entry["quantity"]
                total_sale_amount += entry["total_amount"]

        # Menambahkan baris total untuk Purchase
        self.purchase_tree.insert("", "end", values=(
            "Total",
            "",
            "",
            "",
            total_purchase_quantity,
            "",
            self.format_rupiah(total_purchase_amount)
        ), tags=('totalrow',))

        # Menambahkan baris total untuk Sale
        self.sale_tree.insert("", "end", values=(
            "Total",
            "",
            "",
            "",
            total_sale_quantity,
            "",
            self.format_rupiah(total_sale_amount)
        ), tags=('totalrow',))

    def back_to_dashboard(self):
        self.master.destroy()
        from views.dashboard_view import DashboardView
        import tkinter as tk
        dashboard_window = tk.Tk()
        DashboardView(dashboard_window)
        dashboard_window.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    app = JournalView(root)
    root.mainloop()
