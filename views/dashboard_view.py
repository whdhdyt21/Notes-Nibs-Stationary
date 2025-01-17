import tkinter as tk
from PIL import Image, ImageTk

class DashboardView:
    def __init__(self, master):
        self.master = master
        self.master.title("Dashboard - Notes & Nibs Stationery")
        self.master.geometry("1000x600")
        self.master.configure(bg="#FFF5DE")
        self.master.minsize(1000, 600)

        self.bg_image = Image.open("assets/login_bg.png")
        self.bg_image = self.bg_image.resize((1000, 600), Image.Resampling.LANCZOS)
        self.bg_photo = ImageTk.PhotoImage(self.bg_image)
        tk.Label(self.master, image=self.bg_photo).place(relwidth=1, relheight=1)

        self.frame = tk.Frame(self.master, bg="#FFF5DE", padx=20, pady=20)
        self.frame.place(relx=0.5, rely=0.5, anchor="center")

        logo_img = ImageTk.PhotoImage(Image.open("assets/logo.png").resize((190, 190), Image.Resampling.LANCZOS))
        self.logo_label = tk.Label(self.frame, image=logo_img, bg="#FFF5DE")
        self.logo_label.image = logo_img
        self.logo_label.grid(row=0, column=0, pady=(20, 30))

        button_frame = tk.Frame(self.frame, bg="#FFF5DE")
        button_frame.grid(row=1, column=0, pady=20, sticky="nsew")

        button_config = {
            "font": ("Comic Sans MS", 16, "bold"),
            "bg": "#FFF5DE",
            "fg": "#E1A269",
            "relief": "flat",
            "width": 20,
            "height": 2
        }

        tk.Button(button_frame, text="Inventory", command=self.show_inventory, **button_config).grid(row=0, column=0, pady=15, padx=10)
        tk.Button(button_frame, text="Purchase", command=self.show_purchase, **button_config).grid(row=1, column=0, pady=15, padx=10)
        tk.Button(button_frame, text="Sales", command=self.show_sales, **button_config).grid(row=0, column=1, pady=15, padx=10)
        tk.Button(button_frame, text="Journal", command=self.show_journal, **button_config).grid(row=1, column=1, pady=15, padx=10)

        tk.Button(
            button_frame,
            text="Logout",
            command=self.logout,
            font=("Comic Sans MS", 16, "bold"),
            bg="#FFF5DE",
            fg="#FF2929",
            relief="flat",
            width=20,
            height=2
        ).grid(row=2, column=0, columnspan=2, pady=(30, 20))

    def show_inventory(self):
        self.frame.destroy()
        from views.inventory_view import InventoryView
        InventoryView(self.master)

    def show_purchase(self):
        self.frame.destroy()
        from views.purchase_view import PurchaseView
        PurchaseView(self.master)

    def show_sales(self):
        self.frame.destroy()
        from views.sales_view import SalesView
        SalesView(self.master)

    def show_journal(self):
        self.frame.destroy()
        from views.journal_view import JournalView
        JournalView(self.master)

    def logout(self):
        self.frame.destroy()
        from views.login_view import LoginView
        LoginView(self.master)