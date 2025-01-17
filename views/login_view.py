import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from utils.file_utils import read_excel, append_excel
from views.dashboard_view import DashboardView

USER_FILE = "database/user_accounts.xlsx"

class LoginView:
    def __init__(self, master):
        self.master = master
        self.master.title("Login - Notes & Nibs Stationery")
        self.master.geometry("1000x600")
        self.master.minsize(1000, 600)

        self.bg_photo = ImageTk.PhotoImage(Image.open("assets/login_bg.png").resize((1000, 600), Image.Resampling.LANCZOS))
        tk.Label(self.master, image=self.bg_photo).place(relwidth=1, relheight=1)

        self.frame = tk.Frame(self.master, bg="#FFF5DE", padx=20, pady=20)
        self.frame.place(relx=0.5, rely=0.5, anchor="center")

        logo_img = ImageTk.PhotoImage(Image.open("assets/logo.png").resize((194, 194), Image.Resampling.LANCZOS))
        tk.Label(self.frame, image=logo_img, bg="#FFF5DE").grid(row=0, column=0)
        self.logo_label = logo_img

        form_frame = tk.Frame(self.frame, bg="#FFF5DE")
        form_frame.grid(row=2, column=0, pady=10, sticky="nsew")
        form_frame.grid_columnconfigure(1, weight=1)

        tk.Label(form_frame, text="Username", font=("Arial", 16, "bold"), bg="#FFF5DE", fg="#E1A269").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        self.username_entry = tk.Entry(form_frame, font=("Comic Sans MS", 16))
        self.username_entry.grid(row=0, column=1, pady=5, padx=5, sticky="ew")

        tk.Label(form_frame, text="Password", font=("Comic Sans MS", 16, "bold"), bg="#FFF5DE", fg="#E1A269").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        self.password_entry = tk.Entry(form_frame, show="*", font=("Comic Sans MS", 16))
        self.password_entry.grid(row=1, column=1, pady=5, padx=5, sticky="ew")

        button_frame = tk.Frame(self.frame, bg="#FFF5DE")
        button_frame.grid(row=3, column=0, pady=20)

        tk.Button(button_frame, text="Sign In", command=self.sign_in, bg="#FFF5DE", fg="#E1A269", font=("Comic Sans MS", 16, "bold"), width=15).grid(row=0, column=0, padx=10)
        tk.Button(button_frame, text="Sign Up", command=self.sign_up, bg="#FFF5DE", fg="#E1A269", font=("Comic Sans MS", 16, "bold"), width=15).grid(row=0, column=1, padx=10)

    def sign_in(self):
        username, password = self.username_entry.get(), self.password_entry.get()
        for user in read_excel(USER_FILE):
            if user["username"] == username and user["password"] == password:
                self.frame.destroy()
                DashboardView(self.master)
                return
        messagebox.showerror("Login Failed", "Invalid username or password.")

    def sign_up(self):
        username, password = self.username_entry.get(), self.password_entry.get()
        if not username or not password:
            messagebox.showwarning("Input Error", "Please fill all fields.")
            return

        if any(user["username"] == username for user in read_excel(USER_FILE)):
            messagebox.showerror("Sign Up Failed", "Username already exists.")
            return

        append_excel(USER_FILE, {"username": username, "password": password}, ["username", "password"])
        messagebox.showinfo("Sign Up Success", "Account created successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = LoginView(root)
    root.mainloop()
