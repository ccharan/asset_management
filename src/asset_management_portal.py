"""
Asset Management Portal

Description:
This application is designed to manage and track assets within an organization.
Features include adding, removing, updating, and downloading asset details,
user authentication and enhanced dashboard functionalities.

Modules and Features:
- User Authentication Control
- Asset Management (Add, Remove, Update, Search, Filter)
- Dashboard with Key Metrics and Visualizations
- Downloadable Reports

Author:
    Name: Charan C
    Email: charanc1996@gmail.com
    GitHub: https://github.com/ccharan

Version:
    1.0.0

Date:
    31-07-2024

Dependencies:
    - tkinter
    - psycopg2
    - bcrypt
    - customtkinter

License:
    MIT License

Instructions:
    1. Ensure all dependencies are installed.
    2. Configure the database connection parameters.
    3. Run the script to start the application.

Usage:
    python asset_management_portal.py

"""


import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import psycopg2
import bcrypt
import csv
from datetime import datetime
import customtkinter
import ctypes
import platform


def make_dpi_aware():
    if platform.system() == "Windows" and int(platform.release()) >= 8:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)


make_dpi_aware()
customtkinter.set_appearance_mode("Light")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"
img = b'iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAYAAAD0eNT6AAAACXBIWXMAAA7DAAAOwwHHb6hkAAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAIABJREFUeJzs3Xl8XHW9//HX58wkbbqXbmlpi10pLbco5SqbC4p3xe1Cy47KUjYRxH0lV7lyFQWvCAIiYIECrYJed38gXgFRoAhiWZpS6JqkbaB7mmXO5/fHJGmWmWyznJnM+/l4DMn5nnO+55NJ6Pc9ZwURERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERkbyyqAuQIlXlcfZvGEmsbBRhYwXxYHj7vNCGY15+YOEgThCO7LS+20jwePu0WRkwlMB3d96Q7cI80WG6Cfe9XZbZQeCe/D7RQhjbTXn5bppiDVRN3JPJjykiMlgpAJSaT9UOx8PxlDWPJxFMwBiHcxAxhuGMASpwKggYC1QAw8BGYz68dXoUMAKj7ECn3v0vqb/TKdu8H8v22LYTaMDYh7Oj9fsG8Ddwb8CCfVi4E2wvRgPO67jXQ1BPEG6HlnqaZtdTZWGKLYiIFCUFgKLmxhXrKwnKKrHERMzG4YwDxgPjMBsHTEx+z3jwccBQIPMBu9N0wQeANG3e83Kd2x2nnoB6YDt4PWb1ONuBbQS+HWL1JLweC+pobNlE1Yz9aXoVEYmcAkAhu3j9WMrKZhLzKeCTMZ+C22RgJuZTwA4Bkrve+zTg9jDgKQCkNuDtOBhv4NSAbcGoAd+CsQ7zGiy2hZaWtXxu1s40WxYRySkFgKhUeZydm95Ec3wOQTgbsxmYH4xzMDAd88lAh2PkXdYf0LQCQOe2nAeAvqz7OrAZ8w24bcZ8E1g1YbgWp1oBQURyRQEg1z62YQoWm4/ZTNxnHvjKfJLH1JO6DZDeZbrr/IFMKwB0biuIANBbn28A68DWYb4O/AU8XE25V/PxObvS9CYi0isFgGy4rH4UicaFBOHheDAbfA7GXGAmkDwbvt8DpAKAAkC6tva6a4BqjGrwtVjwEi32HJ865DWs6x+QiEhnCgD99bENUwhji3Dmgy0g8EU48zCCTstl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'


def center_window(window, width=0, height=0):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')


class AssetManagementApp(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        center_window(self, 1200, 800)
        self.title("Assets Management Portal")
        icon = tk.PhotoImage(data=img)
        self.iconphoto(True, icon)
        self.conn = self.connect_to_db()
        self.create_tables()
        self.current_user = None
        self.register_frame = None
        self.show_login()

    def connect_to_db(self):
        try:
            conn = psycopg2.connect(
                dbname="your_db",
                user="your_user",
                password="your_password",
                host="your_host",
                port="5432"
            )
            return conn
        except Exception as e:
            messagebox.showerror("Database Connection Error", str(e))
            self.destroy()

    def create_tables(self):
        with self.conn.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS assets (
                    id SERIAL PRIMARY KEY,
                    employee_name VARCHAR(500) NOT NULL,
                    employee_id VARCHAR(500) UNIQUE NOT NULL,
                    email_id VARCHAR(500) NOT NULL,
                    location VARCHAR(500) NOT NULL,
                    hostname VARCHAR(500) NOT NULL,
                    processor VARCHAR(500) NOT NULL,
                    ram VARCHAR(500) NOT NULL,
                    hd_size VARCHAR(500) NOT NULL,
                    mouse VARCHAR(500) NOT NULL,
                    adaptor VARCHAR(500) NOT NULL,
                    headset VARCHAR(500) NOT NULL,
                    monitor VARCHAR(500) NOT NULL,
                    it_others VARCHAR(500) NOT NULL,
                    software_licenses VARCHAR(500) NOT NULL,
                    asset_entry_date VARCHAR(500) NOT NULL,
                    entered_by VARCHAR(500),
                    updated_by VARCHAR(500),
                    update_date VARCHAR(500),
                    remove_date VARCHAR(500)
                )
            """)
            cursor.execute("""
                            CREATE TABLE IF NOT EXISTS users (
                                id SERIAL PRIMARY KEY,
                                email_id VARCHAR(255) NOT NULL UNIQUE,
                                employee_name VARCHAR(255) NOT NULL,
                                employee_id VARCHAR(255) NOT NULL UNIQUE,
                                password VARCHAR(255) NOT NULL,
                                role VARCHAR(50) NOT NULL DEFAULT 'user'
                            )
                        """)
            self.conn.commit()

    def show_login(self):
        self.current_user = None
        if self.register_frame:
            self.clear_main_content2()

        self.login_frame = customtkinter.CTkFrame(self)
        self.login_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        customtkinter.CTkLabel(self.login_frame, text="Login", font=("Arial", 16)).pack(anchor="center")
        customtkinter.CTkLabel(self.login_frame, text="Username:").pack(pady=5)
        self.username = tk.StringVar()

        customtkinter.CTkEntry(self.login_frame, textvariable=self.username, width=300).pack(pady=5)
        customtkinter.CTkLabel(self.login_frame, text="Password:").pack(pady=5)
        self.password = tk.StringVar()

        customtkinter.CTkEntry(self.login_frame, textvariable=self.password, show="*", width=300).pack(pady=5)
        customtkinter.CTkButton(self.login_frame, text="Login", command=self.login, width=300).pack(pady=10)
        customtkinter.CTkButton(self.login_frame, text="Register", command=self.show_register, width=300).pack(pady=10)

    def show_register(self):
        self.clear_main_content1()
        self.register_frame = customtkinter.CTkFrame(self)
        self.register_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        customtkinter.CTkLabel(self.register_frame, text="Register", font=("Arial", 16)).pack(anchor="center")
        register_frame = customtkinter.CTkFrame(self.register_frame)
        register_frame.pack(pady=10)

        self.new_username = tk.StringVar()
        self.new_user_id = tk.StringVar()
        self.new_user_email_id = tk.StringVar()
        self.new_password = tk.StringVar()
        self.confirm_password = tk.StringVar()

        customtkinter.CTkLabel(register_frame, text="Employee Name:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(register_frame, textvariable=self.new_username, width=300).grid(row=0, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(register_frame, text="Employee ID:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(register_frame, textvariable=self.new_user_id, width=300).grid(row=1, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(register_frame, text="Email ID:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(register_frame, textvariable=self.new_user_email_id, width=300).grid(row=2, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(register_frame, text="Password:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(register_frame, textvariable=self.new_password, width=300, show="*").grid(row=3, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(register_frame, text="Confirm Password:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(register_frame, textvariable=self.confirm_password, width=300, show="*").grid(row=4, column=1, padx=5, pady=5)

        customtkinter.CTkButton(self.register_frame, text="Register", command=self.register_user).pack(pady=10)
        customtkinter.CTkButton(self.register_frame, text="Back to Login", command=self.show_login).pack(pady=10)

    def login(self):
        username = self.username.get().strip()
        password = self.password.get().strip()

        if not username or not password:
            messagebox.showerror("Error", "All fields are required!")
            return

        try:
            with self.conn.cursor() as cursor:
                cursor.execute("SELECT id, password FROM users WHERE email_id = %s", (username,))
                user = cursor.fetchone()
                if user and bcrypt.checkpw(password.encode('utf-8'), user[1].encode('utf-8')):
                    self.current_user = username
                    self.clear_main_content1()
                    self.create_widgets()
                else:
                    messagebox.showerror("Error", "Invalid username or password")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def register_user(self):
        employee_name = self.new_username.get().strip()
        employee_id = self.new_user_id.get().strip()
        employee_email_id = self.new_user_email_id.get().strip()
        password = self.new_password.get().strip()
        confirm_password = self.confirm_password.get().strip()

        if not employee_name or not employee_id or not employee_email_id or not password or not confirm_password:
            messagebox.showerror("Error", "All fields are required!")
            return

        if password != confirm_password:
            messagebox.showerror("Error", "Passwords do not match!")
            return

        hashed_password = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

        try:
            with self.conn.cursor() as cursor:
                cursor.execute("INSERT INTO users (email_id, employee_name, employee_id, password) VALUES (%s, %s, %s, %s)",
                               (employee_email_id, employee_name, employee_id, hashed_password))
            self.conn.commit()
            messagebox.showinfo("Success", "User registered successfully!")
            self.show_login()
        except psycopg2.IntegrityError:
            messagebox.showerror("Error", "Employee ID or Email already exists")
            self.conn.rollback()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def create_widgets(self):
        # Header
        header = customtkinter.CTkFrame(self,  height=50)
        header.pack(fill=tk.X, padx=10, pady=(10,10))

        customtkinter.CTkLabel(header, text="IT Asset Management",  font=("Arial", 20)).pack(side=tk.LEFT, padx=10, pady=(10,10))
        customtkinter.CTkButton(header, text="User Profile", command=self.show_profile_details).pack(side=tk.RIGHT, padx=10, pady=(10,10))
        customtkinter.CTkButton(header, text="Logout", command=self.logout).pack(side=tk.RIGHT, pady=(10,10))

        # Sidebar
        sidebar = customtkinter.CTkFrame(self, width=200)
        sidebar.pack(side=tk.LEFT, fill=customtkinter.Y, padx=10, pady=(0,10))

        customtkinter.CTkButton(sidebar, text="Dashboard", command=self.show_dashboard).pack(fill=tk.X, padx=10, pady=(5,2))
        customtkinter.CTkButton(sidebar, text="Add Asset", command=self.show_add_asset).pack(fill=tk.X, padx=10, pady=2)
        customtkinter.CTkButton(sidebar, text="Manage Assets", command=self.show_manage_assets).pack(fill=tk.X, padx=10,
                                                                                                              pady=2)
        customtkinter.CTkButton(sidebar, text="Download Reports", command=self.show_download_reports).pack(
            fill=tk.X, padx=10, pady=2)

        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(sidebar,
                                                                       values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.pack(side=tk.BOTTOM, padx=10, pady=(10, 10))

        # Main Content
        self.main_content = customtkinter.CTkFrame(self)
        self.main_content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,10), pady=(0,10))

        self.show_dashboard()

    def show_profile_details(self):
        employee_name, email_id, employee_id = self.get_profile_details()
        details = f"{employee_name}\n{employee_id}\n{email_id}"
        tk.messagebox.showinfo("Profile", details)

    def get_profile_details(self):
        with self.conn.cursor() as cursor:
            cursor.execute("SELECT employee_name, email_id, employee_id FROM users WHERE email_id = %s", (self.current_user,))
            return cursor.fetchone()

    def show_dashboard(self):
        self.clear_main_content()

        dashboard_frame = customtkinter.CTkFrame(self.main_content)
        dashboard_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=(10,10))

        customtkinter.CTkLabel(dashboard_frame, text="Dashboard Overview", font=("Arial", 16)).pack(fill=tk.BOTH)
        overview_frame = customtkinter.CTkFrame(dashboard_frame,  height=100)
        overview_frame.pack(pady=10)

        total_assets = self.get_total_assets()
        assets_added = self.get_assets_added_this_month()
        assets_updated = self.get_assets_updated_this_month()
        assets_removed = self.get_assets_removed_this_month()

        customtkinter.CTkLabel(overview_frame, text=f"Total Active Assets: {total_assets}", font=("Arial", 12)).pack(side=tk.LEFT, padx=10)
        customtkinter.CTkLabel(overview_frame, text=f"Assets Added This Month: {assets_added}", font=("Arial", 12)).pack(side=tk.LEFT, padx=10)
        customtkinter.CTkLabel(overview_frame, text=f"Assets Updated This Month: {assets_updated}", font=("Arial", 12)).pack(side=tk.LEFT, padx=10)
        customtkinter.CTkLabel(overview_frame, text=f"Assets Removed This Month: {assets_removed}", font=("Arial", 12)).pack(side=tk.LEFT, padx=10)

    def show_add_asset(self):
        self.clear_main_content()

        add_asset_frame = customtkinter.CTkFrame(self.main_content)
        add_asset_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=(10,10))

        customtkinter.CTkLabel(add_asset_frame, text="Add Asset", font=("Arial", 16)).pack(fill=tk.BOTH)
        form_frame = customtkinter.CTkFrame(add_asset_frame)
        form_frame.pack(pady=10)

        self.employee_name = tk.StringVar()
        self.employee_id = tk.StringVar()
        self.email_id = tk.StringVar()
        self.location = tk.StringVar()
        self.hostname = tk.StringVar()
        self.processor = tk.StringVar()
        self.ram = tk.StringVar()
        self.hd_size = tk.StringVar()
        self.mouse = tk.StringVar()
        self.adaptor = tk.StringVar()
        self.headset = tk.StringVar()
        self.monitor = tk.StringVar()
        self.it_others = tk.StringVar()
        self.software_licenses = tk.StringVar()
        self.asset_entry_date = tk.StringVar()

        customtkinter.CTkLabel(form_frame, text="Employee Name:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.employee_name, width=300).grid(row=0, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="Employee ID:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.employee_id, width=300).grid(row=1, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="Email ID:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.email_id, width=300).grid(row=2, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="Location:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.location, width=300).grid(row=3, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="Hostname:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.hostname, width=300).grid(row=4, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="Processor:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.processor, width=300).grid(row=5, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="RAM:").grid(row=6, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.ram, width=300).grid(row=6, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="HD Size:").grid(row=7, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.hd_size, width=300).grid(row=7, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="Mouse:").grid(row=8, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.mouse, width=300).grid(row=8, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="Adaptor:").grid(row=9, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.adaptor, width=300).grid(row=9, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="Headset:").grid(row=10, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.headset, width=300).grid(row=10, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="Monitor:").grid(row=11, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.monitor, width=300).grid(row=11, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="IT Others:").grid(row=12, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.it_others, width=300).grid(row=12, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="Software Licenses:").grid(row=13, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.software_licenses, width=300).grid(row=13, column=1, padx=5, pady=5)

        customtkinter.CTkLabel(form_frame, text="Asset Entry Date (YYYY-MM-DD):").grid(row=14, column=0, sticky="e", padx=5, pady=5)
        customtkinter.CTkEntry(form_frame, textvariable=self.asset_entry_date, width=300).grid(row=14, column=1, padx=5, pady=5)

        customtkinter.CTkButton(add_asset_frame, text="Submit",  command=self.add_asset).pack(pady=10)

    def show_manage_assets(self):
        self.clear_main_content()

        manage_assets_frame = customtkinter.CTkFrame(self.main_content)
        manage_assets_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=(10, 10))

        customtkinter.CTkLabel(manage_assets_frame, text="Manage Assets", font=("Arial", 16)).pack(fill=tk.BOTH, padx=5,
                                                                                                   pady=5)
        filter_frame = customtkinter.CTkFrame(manage_assets_frame)
        filter_frame.pack(padx=10, pady=20)

        customtkinter.CTkLabel(filter_frame, text="Search By:").pack(side=tk.LEFT, padx=5)

        customtkinter.CTkLabel(filter_frame, text="Employee Name:").pack(side=tk.LEFT, padx=5)
        self.employee_name = tk.StringVar()
        customtkinter.CTkEntry(filter_frame, textvariable=self.employee_name).pack(side=tk.LEFT, padx=5)

        customtkinter.CTkLabel(filter_frame, text="Employee ID:").pack(side=tk.LEFT, padx=5, pady=5)
        self.employee_id = tk.StringVar()
        customtkinter.CTkEntry(filter_frame, textvariable=self.employee_id).pack(side=tk.LEFT, padx=5)

        customtkinter.CTkLabel(filter_frame, text="Location:").pack(side=tk.LEFT, padx=5, pady=5)
        self.location = tk.StringVar()
        customtkinter.CTkEntry(filter_frame, textvariable=self.location).pack(side=tk.LEFT, padx=5)

        customtkinter.CTkButton(filter_frame, text="Search", command=self.search_assets).pack(pady=(5, 5), padx=(30,5))

        columns = (
            "employee_name", "employee_id", "email_id", "location", "hostname", "processor", "ram", "hd_size", "mouse",
            "adaptor", "headset", "monitor", "it_others", "software_licenses", "asset_entry_date", "entered_by", "updated_by", "update_date")
        self.assets_table = ttk.Treeview(manage_assets_frame, columns=columns, show="headings")
        self.assets_table.heading("employee_name", text="Employee Name")
        self.assets_table.heading("employee_id", text="Employee ID")
        self.assets_table.heading("email_id", text="Email ID")
        self.assets_table.heading("location", text="Location")
        self.assets_table.heading("hostname", text="Hostname")
        self.assets_table.heading("processor", text="Processor")
        self.assets_table.heading("ram", text="RAM")
        self.assets_table.heading("hd_size", text="HD Size")
        self.assets_table.heading("mouse", text="Mouse")
        self.assets_table.heading("adaptor", text="Adaptor")
        self.assets_table.heading("headset", text="Headset")
        self.assets_table.heading("monitor", text="Monitor")
        self.assets_table.heading("it_others", text="IT Others")
        self.assets_table.heading("software_licenses", text="Software Licenses")
        self.assets_table.heading("asset_entry_date", text="Asset Entry Date")
        self.assets_table.heading("entered_by", text="Entered By")
        self.assets_table.heading("updated_by", text="Updated By")
        self.assets_table.heading("update_date", text="Updated Date")

        for col_name in columns:
            self.assets_table.column(col_name, anchor=tk.CENTER)
        self.assets_table.pack(fill=tk.BOTH, expand=True)

        self.load_assets()

        action_frame = customtkinter.CTkFrame(manage_assets_frame)
        action_frame.pack(pady=(5, 10))

        customtkinter.CTkButton(action_frame, text="Update", command=self.update_asset).pack(side=tk.LEFT, padx=(0, 5))
        customtkinter.CTkButton(action_frame, text="Remove", command=self.remove_asset).pack(side=tk.LEFT, padx=(5, 0))

    def show_download_reports(self):
        self.clear_main_content()

        download_reports_frame = customtkinter.CTkFrame(self.main_content)
        download_reports_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=(10,10))

        customtkinter.CTkLabel(download_reports_frame, text="Download Reports", font=("Arial", 16)).pack(fill=tk.BOTH, padx=5, pady=5)
        customtkinter.CTkLabel(download_reports_frame, text="Select Report:").pack(anchor="w", padx=5, pady=5)

        report_frame = customtkinter.CTkFrame(download_reports_frame)
        report_frame.pack(fill=tk.X, pady=(10,10), padx=(10,10))

        self.report_type = tk.StringVar(value="All Assets")
        customtkinter.CTkRadioButton(report_frame, text="All Assets", variable=self.report_type, value="All Assets").pack(side=tk.LEFT, padx=5, pady=(10,10))
        customtkinter.CTkRadioButton(report_frame, text="All Active Assets", variable=self.report_type, value="All Active Assets").pack(side=tk.LEFT, padx=5, pady=(10,10))
        customtkinter.CTkRadioButton(report_frame, text="Assets Added This Month", variable=self.report_type, value="Assets Added This Month").pack(side=tk.LEFT, padx=5, pady=(10,10))
        customtkinter.CTkRadioButton(report_frame, text="Assets Updated This Month", variable=self.report_type, value="Assets Updated This Month").pack(side=tk.LEFT, padx=5, pady=(10,10))
        customtkinter.CTkRadioButton(report_frame, text="Assets Removed This Month", variable=self.report_type, value="Assets Removed This Month").pack(side=tk.LEFT, padx=5, pady=(10,10))

        customtkinter.CTkButton(download_reports_frame, text="Download",  command=self.download_report).pack(pady=10)

    def add_asset(self):
        employee_name = self.employee_name.get().strip()
        employee_id = self.employee_id.get().strip()
        email_id = self.email_id.get().strip()
        location = self.location.get().strip()
        hostname = self.hostname.get().strip()
        processor = self.processor.get().strip()
        ram = self.ram.get().strip()
        hd_size = self.hd_size.get().strip()
        mouse = self.mouse.get().strip()
        adaptor = self.adaptor.get().strip()
        headset = self.headset.get().strip()
        monitor = self.monitor.get().strip()
        it_others = self.it_others.get().strip()
        software_licenses = self.software_licenses.get().strip()
        asset_entry_date = self.asset_entry_date.get().strip()

        if not all([employee_name, employee_id, email_id, location, hostname, processor, ram, hd_size, mouse, adaptor, headset, monitor, it_others, software_licenses, asset_entry_date]):
            messagebox.showerror("Error", "All fields are required!")
            return

        if not self.validate_date(asset_entry_date):
            messagebox.showerror("Error", "Dates must be in YYYY-MM-DD format.")
            return

        try:
            with self.conn.cursor() as cursor:
                cursor.execute("""
                            INSERT INTO assets (employee_name, employee_id, email_id, location, hostname, processor, ram, hd_size, mouse, adaptor, headset, monitor, it_others, software_licenses, asset_entry_date, entered_by)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """, (employee_name, employee_id, email_id, location, hostname, processor, ram, hd_size, mouse, adaptor, headset, monitor, it_others, software_licenses, asset_entry_date, self.current_user))
            self.conn.commit()
            messagebox.showinfo("Success", "Asset added successfully!")
            self.show_dashboard()
        except psycopg2.IntegrityError: # UniqueViolation
            messagebox.showerror("Error", "Asset with this Employee ID already exists.")
            self.conn.rollback()

    def validate_date(self, date_text):
        try:
            datetime.strptime(date_text, '%Y-%m-%d')
            return True
        except ValueError:
            return False

    def load_assets(self):
        for row in self.assets_table.get_children():
            self.assets_table.delete(row)

        with self.conn.cursor() as cursor:
            cursor.execute("SELECT * FROM assets WHERE remove_date IS NULL")
            for row in cursor:
                self.assets_table.insert("", tk.END, values=row[1:])

    def search_assets(self):
        employee_name = self.employee_name.get().strip()
        employee_id = self.employee_id.get().strip()
        location = self.location.get().strip()

        for row in self.assets_table.get_children():
            self.assets_table.delete(row)

        query = "SELECT * FROM assets WHERE 1=1"
        params = []

        if employee_name:
            query += " AND employee_name ILIKE %s"
            params.append(f"%{employee_name}%")

        if employee_id:
            query += " AND employee_id ILIKE %s"
            params.append(f"%{employee_id}%")

        if location:
            query += " AND location ILIKE %s"
            params.append(f"%{location}%")

        query += " AND remove_date IS NULL"

        with self.conn.cursor() as cursor:
            cursor.execute(query, params)
            for row in cursor:
                self.assets_table.insert("", tk.END, values=row[1:])

    def update_asset(self):
        selected_item = self.assets_table.selection()
        if not selected_item:
            messagebox.showerror("Error", "No asset selected!")
            return

        item = self.assets_table.item(selected_item)
        asset_details = item['values']

        update_window = customtkinter.CTkToplevel(self)
        update_window.title("Update Asset")
        update_window.geometry()

        update_form_frame = customtkinter.CTkFrame(update_window)
        update_form_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        customtkinter.CTkLabel(update_form_frame, text="Employee Name:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        employee_name_entry = customtkinter.CTkEntry(update_form_frame)
        employee_name_entry.grid(row=0, column=1, padx=5, pady=5)
        employee_name_entry.insert(0, asset_details[0])

        customtkinter.CTkLabel(update_form_frame, text="Employee ID:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        employee_id_entry = customtkinter.CTkEntry(update_form_frame)
        employee_id_entry.grid(row=1, column=1, padx=5, pady=5)
        employee_id_entry.insert(0, asset_details[1])
        employee_id_entry.configure(state="disabled")

        customtkinter.CTkLabel(update_form_frame, text="Email ID:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        email_entry = customtkinter.CTkEntry(update_form_frame)
        email_entry.grid(row=2, column=1, padx=5, pady=5)
        email_entry.insert(0, asset_details[2])

        customtkinter.CTkLabel(update_form_frame, text="Location:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        location_entry = customtkinter.CTkEntry(update_form_frame)
        location_entry.grid(row=3, column=1, padx=5, pady=5)
        location_entry.insert(0, asset_details[3])

        customtkinter.CTkLabel(update_form_frame, text="Hostname:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        hostname_entry = customtkinter.CTkEntry(update_form_frame)
        hostname_entry.grid(row=4, column=1, padx=5, pady=5)
        hostname_entry.insert(0, asset_details[4])

        customtkinter.CTkLabel(update_form_frame, text="Processor:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        processor_entry = customtkinter.CTkEntry(update_form_frame)
        processor_entry.grid(row=5, column=1, padx=5, pady=5)
        processor_entry.insert(0, asset_details[5])

        customtkinter.CTkLabel(update_form_frame, text="RAM:").grid(row=6, column=0, sticky="w", padx=5, pady=5)
        ram_entry = customtkinter.CTkEntry(update_form_frame)
        ram_entry.grid(row=6, column=1, padx=5, pady=5)
        ram_entry.insert(0, asset_details[6])

        customtkinter.CTkLabel(update_form_frame, text="HDD Size:").grid(row=7, column=0, sticky="w", padx=5, pady=5)
        hd_size_entry = customtkinter.CTkEntry(update_form_frame)
        hd_size_entry.grid(row=7, column=1, padx=5, pady=5)
        hd_size_entry.insert(0, asset_details[7])

        customtkinter.CTkLabel(update_form_frame, text="Mouse:").grid(row=8, column=0, sticky="w", padx=5, pady=5)
        mouse_entry = customtkinter.CTkEntry(update_form_frame)
        mouse_entry.grid(row=8, column=1, padx=5, pady=5)
        mouse_entry.insert(0, asset_details[8])

        customtkinter.CTkLabel(update_form_frame, text="Adaptor:").grid(row=9, column=0, sticky="w", padx=5, pady=5)
        adaptor_entry = customtkinter.CTkEntry(update_form_frame)
        adaptor_entry.grid(row=9, column=1, padx=5, pady=5)
        adaptor_entry.insert(0, asset_details[9])

        customtkinter.CTkLabel(update_form_frame, text="Headset:").grid(row=10, column=0, sticky="w", padx=5, pady=5)
        headset_entry = customtkinter.CTkEntry(update_form_frame)
        headset_entry.grid(row=10, column=1, padx=5, pady=5)
        headset_entry.insert(0, asset_details[10])

        customtkinter.CTkLabel(update_form_frame, text="Monitor:").grid(row=11, column=0, sticky="w", padx=5, pady=5)
        monitor_entry = customtkinter.CTkEntry(update_form_frame)
        monitor_entry.grid(row=11, column=1, padx=5, pady=5)
        monitor_entry.insert(0, asset_details[11])

        customtkinter.CTkLabel(update_form_frame, text="IT Others:").grid(row=12, column=0, sticky="w", padx=5, pady=5)
        it_others_entry = customtkinter.CTkEntry(update_form_frame)
        it_others_entry.grid(row=12, column=1, padx=5, pady=5)
        it_others_entry.insert(0, asset_details[12])

        customtkinter.CTkLabel(update_form_frame, text="Software Licenses:").grid(row=13, column=0, sticky="w", padx=5, pady=5)
        software_licenses_entry = customtkinter.CTkEntry(update_form_frame)
        software_licenses_entry.grid(row=13, column=1, padx=5, pady=5)
        software_licenses_entry.insert(0, asset_details[13])

        customtkinter.CTkLabel(update_form_frame, text="Asset Entry Date:").grid(row=14, column=0, sticky="w", padx=5, pady=5)
        asset_entry_date_entry = customtkinter.CTkEntry(update_form_frame)
        asset_entry_date_entry.grid(row=14, column=1, padx=5, pady=5)
        asset_entry_date_entry.insert(0, asset_details[14])

        def save_update():
            updated_employee_name = employee_name_entry.get().strip()
            updated_employee_id = employee_id_entry.get().strip()
            updated_email_id = email_entry.get().strip()
            updated_location = location_entry.get().strip()
            updated_hostname = hostname_entry.get().strip()
            updated_processor = processor_entry.get().strip()
            updated_ram= ram_entry.get().strip()
            updated_hd_size= hd_size_entry.get().strip()
            updated_mouse = mouse_entry.get().strip()
            updated_adaptor= adaptor_entry.get().strip()
            updated_headset = headset_entry.get().strip()
            updated_monitor = monitor_entry.get().strip()
            updated_it_others = it_others_entry.get().strip()
            updated_software_licenses = software_licenses_entry.get().strip()
            updated_asset_entry_date= asset_entry_date_entry.get().strip()
            updated_date = datetime.now().strftime('%Y-%m-%d')

            if not all([updated_employee_name, updated_employee_id, updated_email_id, updated_location, updated_hostname,
                        updated_processor, updated_ram, updated_hd_size, updated_mouse, updated_adaptor, updated_headset,
                        updated_monitor, updated_it_others, updated_software_licenses,  updated_asset_entry_date]):
                messagebox.showerror("Error", "All fields are required!")
                return

            if not self.validate_date(updated_asset_entry_date):
                messagebox.showerror("Error", "Dates must be in YYYY-MM-DD format.")
                return

            try:
                with self.conn.cursor() as cursor:
                    cursor.execute("""
                        UPDATE assets
                        SET employee_name = %s, employee_id = %s, email_id = %s, location = %s, hostname = %s, processor = %s, ram = %s, hd_size = %s, mouse = %s, adaptor = %s, headset = %s, monitor= %s, it_others = %s, software_licenses = %s, asset_entry_date = %s, update_date = %s, updated_by = %s
                        WHERE employee_id = %s
                    """, (updated_employee_name, updated_employee_id, updated_email_id, updated_location, updated_hostname,
                          updated_processor, updated_ram, updated_hd_size, updated_mouse, updated_adaptor, updated_headset,
                          updated_monitor, updated_it_others, updated_software_licenses, updated_asset_entry_date, updated_date,
                          self.current_user, updated_employee_id))
                    self.conn.commit()

                messagebox.showinfo("Success", "Asset updated successfully!")
                self.load_assets()
                update_window.destroy()
            except psycopg2.IntegrityError:
                messagebox.showerror("Error", "Asset with this Employee ID already exists.")
                self.conn.rollback()

        customtkinter.CTkButton(update_window, text="Save", command=save_update).pack(pady=10)

    def remove_asset(self):
        selected_item = self.assets_table.selection()
        if not selected_item:
            messagebox.showerror("Error", "No asset selected!")
            return

        item = self.assets_table.item(selected_item)
        asset_details = item['values']
        employee_id = asset_details[1]

        confirm = messagebox.askyesno("Confirm", "Are you sure you want to remove this asset?")
        if confirm:
            with self.conn.cursor() as cursor:
                current_date = datetime.now().strftime('%Y-%m-%d')
                cursor.execute("UPDATE assets SET remove_date = %s WHERE employee_id = %s", (current_date, str(employee_id)))
                self.conn.commit()

            messagebox.showinfo("Success", "Asset removed successfully!")
            self.load_assets()

    def download_report(self):
        report_type = self.report_type.get()

        query = "SELECT * FROM assets"
        if report_type == "All Active Assets":
            query += " WHERE remove_date IS NULL"
        elif report_type == "Assets Added This Month":
            query += """ WHERE TO_DATE(asset_entry_date, 'YYYY-MM-DD') >= DATE_TRUNC('month', CURRENT_DATE)
                        AND TO_DATE(asset_entry_date, 'YYYY-MM-DD') < DATE_TRUNC('month', CURRENT_DATE) + INTERVAL '1 month' """
        elif report_type == "Assets Updated This Month":
            query += """ WHERE TO_DATE(update_date, 'YYYY-MM-DD') >= DATE_TRUNC('month', CURRENT_DATE)
                        AND TO_DATE(update_date, 'YYYY-MM-DD') < DATE_TRUNC('month', CURRENT_DATE) + INTERVAL '1 month' """
        elif report_type == "Assets Removed This Month":
            query += """ WHERE TO_DATE(remove_date, 'YYYY-MM-DD') >= DATE_TRUNC('month', CURRENT_DATE)
                        AND TO_DATE(remove_date, 'YYYY-MM-DD') < DATE_TRUNC('month', CURRENT_DATE) + INTERVAL '1 month' """

        assets = []
        with self.conn.cursor() as cursor:
            cursor.execute(query)
            assets = cursor.fetchall()
        print(assets)

        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not file_path:
            return

        with open(file_path, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["ID", "Employee Name", "Employee ID", "Email ID", "Location", "Hostname", "Processor",
                             "RAM", "HDD Size", "Mouse", "Adaptor", "Headset", "Monitor", "IT Others", "Software Licenses",
                             "Asset Entry Date", "Asset Entered By", "Asset Updated By", "Asset Updated Date", "Asset Remove Date"])
            for asset in assets:
                writer.writerow(asset)

        messagebox.showinfo("Success", f"Report saved to {file_path}")

    def get_total_assets(self):
        with self.conn.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) FROM assets WHERE remove_date IS NULL")
            return cursor.fetchone()[0]

    def get_assets_added_this_month(self):
        query = """SELECT COUNT(*)
                  FROM assets
                  WHERE TO_DATE(asset_entry_date, 'YYYY-MM-DD') >= DATE_TRUNC('month', CURRENT_DATE)
                  AND TO_DATE(asset_entry_date, 'YYYY-MM-DD') < DATE_TRUNC('month', CURRENT_DATE) + INTERVAL '1 month'
        """
        with self.conn.cursor() as cursor:
            cursor.execute(query)

            return cursor.fetchone()[0]

    def get_assets_updated_this_month(self):
        query = """SELECT COUNT(*)
                  FROM assets
                  WHERE TO_DATE(update_date, 'YYYY-MM-DD') >= DATE_TRUNC('month', CURRENT_DATE)
                  AND TO_DATE(update_date, 'YYYY-MM-DD') < DATE_TRUNC('month', CURRENT_DATE) + INTERVAL '1 month'
                """
        with self.conn.cursor() as cursor:
            cursor.execute(query)
            return cursor.fetchone()[0]

    def get_assets_removed_this_month(self):
        query = """SELECT COUNT(*)
                  FROM assets
                  WHERE TO_DATE(remove_date, 'YYYY-MM-DD') >= DATE_TRUNC('month', CURRENT_DATE)
                  AND TO_DATE(remove_date, 'YYYY-MM-DD') < DATE_TRUNC('month', CURRENT_DATE) + INTERVAL '1 month'
                """
        with self.conn.cursor() as cursor:
            cursor.execute(query)
            return cursor.fetchone()[0]

    def change_appearance_mode_event(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def clear_main_content(self):
        for widget in self.main_content.winfo_children():
            widget.destroy()

    def clear_main_content1(self):
        self.login_frame.destroy()

    def clear_main_content2(self):
        self.register_frame.destroy()

    def logout(self):
        self.destroy()


if __name__ == "__main__":
    app = AssetManagementApp()
    app.mainloop()
