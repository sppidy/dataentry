import tkinter as tk
from tkinter import ttk, messagebox
from program import main

class LoginWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Login")
        self.iconbitmap("906310.ico")
        self.geometry("325x200")
        self.configure(background="#CCCCFF")
        self.create_widgets()
        style = ttk.Style(self)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#CCCCFF")
        style.configure("Custom.TLabelframe", bordercolor="black", background="#CCCCFF")
        style.configure("Custom.TLabelframe.Label", background="#CCCCFF")
        # style.configure("Custom.TLabelFrame.Border", background="#CCCCFF")
        style.configure("Custom.TEntry", background="lightblue")
        style.configure("Custom.TLabel", background="#CCCCFF")
    def create_widgets(self):
        global login_entry
        global password_entry
        frame = ttk.Frame(self, padding=20, style="Custom.TFrame")
        frame.pack()
        
        dataframe = ttk.LabelFrame(frame, text="Login", borderwidth=2, relief="groove",style="Custom.TLabelframe")
        dataframe.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        login_label = ttk.Label(dataframe, text="Username:",style="Custom.TLabel")
        login_label.grid(row=0, column=0, padx=10, pady=10)
        login_entry = ttk.Entry(dataframe, width=20)
        login_entry.grid(row=0, column=1, padx=10, pady=10)

        password_label = ttk.Label(dataframe, text="Password:",style="Custom.TLabel")
        password_label.grid(row=1, column=0, padx=10, pady=10)
        password_entry = ttk.Entry(dataframe, width=20, show="*")
        password_entry.grid(row=1, column=1, padx=10, pady=10)

        login_button = ttk.Button(frame, text="Login", command=self.login)
        login_button.grid(row=1, column=0, padx=10, pady=10)

    def login(self):
        global login_entry
        global password_entry
        
        username = login_entry.get()
        password = password_entry.get()
        
        if not username or not password:
            messagebox.showerror("Login", "Please enter both username and password.")
            return

        if username == "thirupathi" and password == "LBM123":
            self.destroy()
            main()
        else:
            messagebox.showerror("Login", "Incorrect username or password.")

if __name__ == "__main__":
    login_window = LoginWindow()
    login_window.mainloop()
