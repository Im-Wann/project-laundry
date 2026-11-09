import sqlite3
import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os


class LaundryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🧺 Laundry Management System")
        self.root.geometry("1000x650")
        self.root.minsize(900, 600)

        # Database connection
        self.conn = sqlite3.connect('laundry.db', check_same_thread=False)
        self.init_database()

        # Style configuration
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_styles()

        # Main container
        self.main_frame = ttk.Frame(root, padding=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.build_ui()
        self.refresh_all_data()

    def configure_styles(self):
        self.style.configure('Title.TLabel', font=('Segoe UI', 24, 'bold'), foreground='#2c3e50')
        self.style.configure('Subtitle.TLabel', font=('Segoe UI', 12), foreground='#7f8c8d')
        self.style.configure('Card.TFrame', background='#ffffff', relief='flat')
        self.style.configure('TButton', font=('Segoe UI', 10), padding=10)
        self.style.configure('Primary.TButton', font=('Segoe UI', 10, 'bold'), padding=10)
        self.style.configure('Danger.TButton', font=('Segoe UI', 10), padding=10, foreground='#c0392b')
        self.style.configure('Success.TButton', font=('Segoe UI', 10), padding=10, foreground='#27ae60')
        self.style.configure('Treeview', font=('Segoe UI', 9), rowheight=30)
        self.style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))
        self.style.map('Primary.TButton', background=[('active', '#2980b9')])
        self.style.map('Success.TButton', background=[('active', '#229954')])
        self.style.map('Danger.TButton', background=[('active', '#c0392b')])

    def init_database(self):
        cursor = self.conn.cursor()
        cursor.execute('PRAGMA foreign_keys = ON')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pelanggan (
                id_pelanggan INTEGER PRIMARY KEY AUTOINCREMENT,
                nama TEXT NOT NULL
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS transaksi (
                id_transaksi INTEGER PRIMARY KEY AUTOINCREMENT,
                id_pelanggan INTEGER NOT NULL,
                tanggal TEXT NOT NULL,
                status TEXT NOT NULL,
                FOREIGN KEY (id_pelanggan) REFERENCES pelanggan (id_pelanggan)
            )
        ''')
        self.conn.commit()

    def build_ui(self):
        # Header
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))

        ttk.Label(header_frame, text="Laundry Management", style='Title.TLabel').pack(anchor=tk.W)
        ttk.Label(header_frame, text="Kelola pelanggan, transaksi, dan laporan laundry dengan mudah", style='Subtitle.TLabel').pack(anchor=tk.W)

        # Notebook for tabs
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: Pelanggan
        self.tab_pelanggan = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.tab_pelanggan, text="👥  Pelanggan")
        self.build_pelanggan_tab()

        # Tab 2: Transaksi
        self.tab_transaksi = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.tab_transaksi, text="📋  Transaksi")
        self.build_transaksi_tab()

        # Tab 3: Status Update
        self.tab_status = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.tab_status, text="🔄  Update Status")
        self.build_status_tab()

        # Tab 4: Laporan
        self.tab_laporan = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.tab_laporan, text="📊  Laporan & Export")
        self.build_laporan_tab()

    def build_pelanggan_tab(self):
        # Input section
        input_frame = ttk.LabelFrame(self.tab_pelanggan, text="Tambah Pelanggan Baru", padding=15)
        input_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(input_frame, text="Nama Pelanggan:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.entry_nama = ttk.Entry(input_frame, width=40, font=('Segoe UI', 10))
        self.entry_nama.grid(row=0, column=1, sticky=tk.EW, pady=5)
        self.entry_nama.bind('<Return>', lambda e: self.tambah_pelanggan())

        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=0, column=2, padx=(15, 0))
        ttk.Button(btn_frame, text="➕ Tambah", style='Primary.TButton', command=self.tambah_pelanggan).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 Refresh", command=self.refresh_pelanggan).pack(side=tk.LEFT, padx=5)

        input_frame.columnconfigure(1, weight=1)

        # Table section
        table_frame = ttk.LabelFrame(self.tab_pelanggan, text="Daftar Pelanggan", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True)

        columns = ('ID', 'Nama')
        self.tree_pelanggan = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='browse')
        self.tree_pelanggan.heading('ID', text='ID Pelanggan')
        self.tree_pelanggan.heading('Nama', text='Nama Pelanggan')
        self.tree_pelanggan.column('ID', width=120, anchor=tk.CENTER)
        self.tree_pelanggan.column('Nama', width=400, anchor=tk.W)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree_pelanggan.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree_pelanggan.xview)
        self.tree_pelanggan.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree_pelanggan.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar_y.grid(row=0, column=1, sticky=tk.NS)
        scrollbar_x.grid(row=1, column=0, sticky=tk.EW)

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Delete button
        ttk.Button(self.tab_pelanggan, text="🗑️ Hapus Pelanggan Terpilih", style='Danger.TButton',
                   command=self.hapus_pelanggan).pack(pady=(10, 0))

    def build_transaksi_tab(self):
        # Input section
        input_frame = ttk.LabelFrame(self.tab_transaksi, text="Transaksi Baru", padding=15)
        input_frame.pack(fill=tk.X, pady=(0, 15))

        # Pelanggan dropdown
        ttk.Label(input_frame, text="Pelanggan:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.combo_pelanggan = ttk.Combobox(input_frame, width=37, font=('Segoe UI', 10), state='readonly')
        self.combo_pelanggan.grid(row=0, column=1, sticky=tk.EW, pady=5)

        # Tanggal
        ttk.Label(input_frame, text="Tanggal (YYYY-MM-DD):").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.entry_tanggal = ttk.Entry(input_frame, width=40, font=('Segoe UI', 10))
        self.entry_tanggal.grid(row=1, column=1, sticky=tk.EW, pady=5)
        self.entry_tanggal.insert(0, datetime.now().strftime('%Y-%m-%d'))

        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=0, column=2, rowspan=2, padx=(15, 0))
        ttk.Button(btn_frame, text="➕ Buat Transaksi", style='Primary.TButton', command=self.tambah_transaksi).pack(side=tk.LEFT, padx=5, pady=2)
        ttk.Button(btn_frame, text="🔄 Refresh", command=self.refresh_transaksi).pack(side=tk.LEFT, padx=5, pady=2)

        input_frame.columnconfigure(1, weight=1)

        # Table section
        table_frame = ttk.LabelFrame(self.tab_transaksi, text="Daftar Transaksi", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True)

        columns = ('ID Transaksi', 'ID Pelanggan', 'Nama Pelanggan', 'Tanggal', 'Status')
        self.tree_transaksi = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='browse')
        for col in columns:
            self.tree_transaksi.heading(col, text=col)
        self.tree_transaksi.column('ID Transaksi', width=100, anchor=tk.CENTER)
        self.tree_transaksi.column('ID Pelanggan', width=100, anchor=tk.CENTER)
        self.tree_transaksi.column('Nama Pelanggan', width=200, anchor=tk.W)
        self.tree_transaksi.column('Tanggal', width=120, anchor=tk.CENTER)
        self.tree_transaksi.column('Status', width=100, anchor=tk.CENTER)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree_transaksi.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree_transaksi.xview)
        self.tree_transaksi.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree_transaksi.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar_y.grid(row=0, column=1, sticky=tk.NS)
        scrollbar_x.grid(row=1, column=0, sticky=tk.EW)

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Tag for status colors
        self.tree_transaksi.tag_configure('PROSES', background='#fff3cd')
        self.tree_transaksi.tag_configure('SELESAI', background='#d4edda')
        self.tree_transaksi.tag_configure('BATAL', background='#f8d7da')

    def build_status_tab(self):
        # Input section
        input_frame = ttk.LabelFrame(self.tab_status, text="Update Status Transaksi", padding=15)
        input_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(input_frame, text="Pilih Transaksi:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.combo_transaksi = ttk.Combobox(input_frame, width=50, font=('Segoe UI', 10), state='readonly')
        self.combo_transaksi.grid(row=0, column=1, sticky=tk.EW, pady=5)

        ttk.Label(input_frame, text="Status Baru:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.combo_status = ttk.Combobox(input_frame, width=50, font=('Segoe UI', 10), state='readonly',
                                          values=['SELESAI', 'BATAL'])
        self.combo_status.grid(row=1, column=1, sticky=tk.EW, pady=5)
        self.combo_status.current(0)

        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=0, column=2, rowspan=2, padx=(15, 0))
        ttk.Button(btn_frame, text="✅ Update Status", style='Success.TButton', command=self.update_status).pack(side=tk.LEFT, padx=5, pady=2)
        ttk.Button(btn_frame, text="🔄 Refresh", command=self.refresh_status_tab).pack(side=tk.LEFT, padx=5, pady=2)

        input_frame.columnconfigure(1, weight=1)

        # Table section
        table_frame = ttk.LabelFrame(self.tab_status, text="Semua Transaksi", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True)

        columns = ('ID Transaksi', 'ID Pelanggan', 'Nama Pelanggan', 'Tanggal', 'Status')
        self.tree_status = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='browse')
        for col in columns:
            self.tree_status.heading(col, text=col)
        self.tree_status.column('ID Transaksi', width=100, anchor=tk.CENTER)
        self.tree_status.column('ID Pelanggan', width=100, anchor=tk.CENTER)
        self.tree_status.column('Nama Pelanggan', width=200, anchor=tk.W)
        self.tree_status.column('Tanggal', width=120, anchor=tk.CENTER)
        self.tree_status.column('Status', width=100, anchor=tk.CENTER)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree_status.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree_status.xview)
        self.tree_status.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree_status.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar_y.grid(row=0, column=1, sticky=tk.NS)
        scrollbar_x.grid(row=1, column=0, sticky=tk.EW)

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree_status.tag_configure('PROSES', background='#fff3cd')
        self.tree_status.tag_configure('SELESAI', background='#d4edda')
        self.tree_status.tag_configure('BATAL', background='#f8d7da')

    def build_laporan_tab(self):
        # Stats cards
        stats_frame = ttk.Frame(self.tab_laporan)
        stats_frame.pack(fill=tk.X, pady=(0, 15))

        self.stat_total_pelanggan = self.create_stat_card(stats_frame, "Total Pelanggan", "0", "#3498db")
        self.stat_total_transaksi = self.create_stat_card(stats_frame, "Total Transaksi", "0", "#27ae60")
        self.stat_proses = self.create_stat_card(stats_frame, "Sedang Proses", "0", "#f39c12")
        self.stat_selesai = self.create_stat_card(stats_frame, "Selesai", "0", "#27ae60")
        self.stat_batal = self.create_stat_card(stats_frame, "Dibatalkan", "0", "#e74c3c")

        # Buttons
        btn_frame = ttk.Frame(self.tab_laporan)
        btn_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Button(btn_frame, text="📥 Export ke Excel", style='Primary.TButton', command=self.export_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 Refresh Data", command=self.refresh_laporan).pack(side=tk.LEFT, padx=5)

        # Data preview
        table_frame = ttk.LabelFrame(self.tab_laporan, text="Preview Data", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True)

        # Notebook for pelanggan/transaksi preview
        preview_notebook = ttk.Notebook(table_frame)
        preview_notebook.pack(fill=tk.BOTH, expand=True)

        # Pelanggan preview
        pelanggan_frame = ttk.Frame(preview_notebook, padding=10)
        preview_notebook.add(pelanggan_frame, text="Pelanggan")

        columns_p = ('ID', 'Nama')
        self.tree_laporan_pelanggan = ttk.Treeview(pelanggan_frame, columns=columns_p, show='headings')
        for col in columns_p:
            self.tree_laporan_pelanggan.heading(col, text=col)
        self.tree_laporan_pelanggan.column('ID', width=100, anchor=tk.CENTER)
        self.tree_laporan_pelanggan.column('Nama', width=300, anchor=tk.W)

        scroll_p = ttk.Scrollbar(pelanggan_frame, orient=tk.VERTICAL, command=self.tree_laporan_pelanggan.yview)
        self.tree_laporan_pelanggan.configure(yscrollcommand=scroll_p.set)
        self.tree_laporan_pelanggan.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_p.pack(side=tk.RIGHT, fill=tk.Y)

        # Transaksi preview
        transaksi_frame = ttk.Frame(preview_notebook, padding=10)
        preview_notebook.add(transaksi_frame, text="Transaksi")

        columns_t = ('ID Transaksi', 'ID Pelanggan', 'Nama Pelanggan', 'Tanggal', 'Status')
        self.tree_laporan_transaksi = ttk.Treeview(transaksi_frame, columns=columns_t, show='headings')
        for col in columns_t:
            self.tree_laporan_transaksi.heading(col, text=col)
        self.tree_laporan_transaksi.column('ID Transaksi', width=100, anchor=tk.CENTER)
        self.tree_laporan_transaksi.column('ID Pelanggan', width=100, anchor=tk.CENTER)
        self.tree_laporan_transaksi.column('Nama Pelanggan', width=180, anchor=tk.W)
        self.tree_laporan_transaksi.column('Tanggal', width=120, anchor=tk.CENTER)
        self.tree_laporan_transaksi.column('Status', width=100, anchor=tk.CENTER)

        scroll_t = ttk.Scrollbar(transaksi_frame, orient=tk.VERTICAL, command=self.tree_laporan_transaksi.yview)
        self.tree_laporan_transaksi.configure(yscrollcommand=scroll_t.set)
        self.tree_laporan_transaksi.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_t.pack(side=tk.RIGHT, fill=tk.Y)

    def create_stat_card(self, parent, title, value, color):
        card = ttk.Frame(parent, style='Card.TFrame', padding=20)
        card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # Color indicator
        indicator = tk.Frame(card, bg=color, height=4)
        indicator.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(card, text=title, font=('Segoe UI', 10), foreground='#7f8c8d').pack(anchor=tk.W)
        value_label = ttk.Label(card, text=value, font=('Segoe UI', 28, 'bold'), foreground=color)
        value_label.pack(anchor=tk.W)

        return value_label

    # --- Database Operations ---
    def execute_query(self, query, params=()):
        cursor = self.conn.cursor()
        cursor.execute(query, params)
        self.conn.commit()
        return cursor

    def fetch_all(self, query, params=()):
        cursor = self.conn.cursor()
        cursor.execute(query, params)
        return cursor.fetchall()

    def fetch_one(self, query, params=()):
        cursor = self.conn.cursor()
        cursor.execute(query, params)
        return cursor.fetchone()

    # --- Pelanggan CRUD ---
    def tambah_pelanggan(self):
        nama = self.entry_nama.get().strip()
        if not nama:
            messagebox.showwarning("Peringatan", "Nama pelanggan tidak boleh kosong!")
            return

        try:
            self.execute_query("INSERT INTO pelanggan (nama) VALUES (?)", (nama,))
            self.entry_nama.delete(0, tk.END)
            self.refresh_pelanggan()
            messagebox.showinfo("Sukses", f"Pelanggan '{nama}' berhasil ditambahkan!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menambah pelanggan: {e}")

    def hapus_pelanggan(self):
        selection = self.tree_pelanggan.selection()
        if not selection:
            messagebox.showwarning("Peringatan", "Pilih pelanggan yang ingin dihapus!")
            return

        item = self.tree_pelanggan.item(selection[0])
        id_pelanggan = item['values'][0]
        nama = item['values'][1]

        if messagebox.askyesno("Konfirmasi", f"Hapus pelanggan '{nama}' (ID: {id_pelanggan})?\nSemua transaksi terkait juga akan terhapus."):
            try:
                self.execute_query("DELETE FROM pelanggan WHERE id_pelanggan = ?", (id_pelanggan,))
                self.refresh_pelanggan()
                messagebox.showinfo("Sukses", "Pelanggan berhasil dihapus!")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal menghapus pelanggan: {e}")

    def refresh_pelanggan(self):
        # Clear tree
        for item in self.tree_pelanggan.get_children():
            self.tree_pelanggan.delete(item)

        # Fetch and populate
        data = self.fetch_all("SELECT id_pelanggan, nama FROM pelanggan ORDER BY id_pelanggan")
        for row in data:
            self.tree_pelanggan.insert('', tk.END, values=row)

        # Update comboboxes
        self.update_pelanggan_combobox()

    def update_pelanggan_combobox(self):
        data = self.fetch_all("SELECT id_pelanggan, nama FROM pelanggan ORDER BY id_pelanggan")
        self.combo_pelanggan['values'] = [f"{row[0]} - {row[1]}" for row in data]
        if data:
            self.combo_pelanggan.current(0)

    # --- Transaksi CRUD ---
    def tambah_transaksi(self):
        pelanggan_text = self.combo_pelanggan.get()
        if not pelanggan_text:
            messagebox.showwarning("Peringatan", "Pilih pelanggan terlebih dahulu!")
            return

        try:
            id_pelanggan = int(pelanggan_text.split(' - ')[0])
        except (ValueError, IndexError):
            messagebox.showerror("Error", "Format pelanggan tidak valid!")
            return

        tanggal = self.entry_tanggal.get().strip()
        if not tanggal:
            messagebox.showwarning("Peringatan", "Tanggal tidak boleh kosong!")
            return

        # Validate date format
        try:
            datetime.strptime(tanggal, '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Error", "Format tanggal salah! Gunakan YYYY-MM-DD")
            return

        try:
            self.execute_query(
                "INSERT INTO transaksi (id_pelanggan, tanggal, status) VALUES (?, ?, ?)",
                (id_pelanggan, tanggal, "PROSES")
            )
            self.entry_tanggal.delete(0, tk.END)
            self.entry_tanggal.insert(0, datetime.now().strftime('%Y-%m-%d'))
            self.refresh_transaksi()
            messagebox.showinfo("Sukses", "Transaksi baru berhasil dibuat dengan status PROSES!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuat transaksi: {e}")

    def refresh_transaksi(self):
        for item in self.tree_transaksi.get_children():
            self.tree_transaksi.delete(item)

        data = self.fetch_all('''
            SELECT t.id_transaksi, t.id_pelanggan, p.nama, t.tanggal, t.status
            FROM transaksi t
            JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
            ORDER BY t.id_transaksi DESC
        ''')
        for row in data:
            self.tree_transaksi.insert('', tk.END, values=row, tags=(row[4],))

        self.update_transaksi_combobox()

    def update_transaksi_combobox(self):
        data = self.fetch_all('''
            SELECT t.id_transaksi, p.nama, t.tanggal, t.status
            FROM transaksi t
            JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
            ORDER BY t.id_transaksi DESC
        ''')
        self.combo_transaksi['values'] = [f"{row[0]} - {row[1]} ({row[2]}) [{row[3]}]" for row in data]
        if data:
            self.combo_transaksi.current(0)

    # --- Status Update ---
    def update_status(self):
        transaksi_text = self.combo_transaksi.get()
        if not transaksi_text:
            messagebox.showwarning("Peringatan", "Pilih transaksi terlebih dahulu!")
            return

        try:
            id_transaksi = int(transaksi_text.split(' - ')[0])
        except (ValueError, IndexError):
            messagebox.showerror("Error", "Format transaksi tidak valid!")
            return

        new_status = self.combo_status.get()
        if not new_status:
            messagebox.showwarning("Peringatan", "Pilih status baru!")
            return

        try:
            self.execute_query("UPDATE transaksi SET status = ? WHERE id_transaksi = ?", (new_status, id_transaksi))
            self.refresh_status_tab()
            messagebox.showinfo("Sukses", f"Status transaksi ID {id_transaksi} berhasil diupdate ke {new_status}!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal update status: {e}")

    def refresh_status_tab(self):
        for item in self.tree_status.get_children():
            self.tree_status.delete(item)

        data = self.fetch_all('''
            SELECT t.id_transaksi, t.id_pelanggan, p.nama, t.tanggal, t.status
            FROM transaksi t
            JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
            ORDER BY t.id_transaksi DESC
        ''')
        for row in data:
            self.tree_status.insert('', tk.END, values=row, tags=(row[4],))

        self.update_transaksi_combobox()

    # --- Laporan & Export ---
    def refresh_laporan(self):
        # Stats
        total_pelanggan = self.fetch_one("SELECT COUNT(*) FROM pelanggan")[0]
        total_transaksi = self.fetch_one("SELECT COUNT(*) FROM transaksi")[0]
        proses = self.fetch_one("SELECT COUNT(*) FROM transaksi WHERE status = 'PROSES'")[0]
        selesai = self.fetch_one("SELECT COUNT(*) FROM transaksi WHERE status = 'SELESAI'")[0]
        batal = self.fetch_one("SELECT COUNT(*) FROM transaksi WHERE status = 'BATAL'")[0]

        self.stat_total_pelanggan.config(text=str(total_pelanggan))
        self.stat_total_transaksi.config(text=str(total_transaksi))
        self.stat_proses.config(text=str(proses))
        self.stat_selesai.config(text=str(selesai))
        self.stat_batal.config(text=str(batal))

        # Pelanggan table
        for item in self.tree_laporan_pelanggan.get_children():
            self.tree_laporan_pelanggan.delete(item)
        pelanggan_data = self.fetch_all("SELECT id_pelanggan, nama FROM pelanggan ORDER BY id_pelanggan")
        for row in pelanggan_data:
            self.tree_laporan_pelanggan.insert('', tk.END, values=row)

        # Transaksi table
        for item in self.tree_laporan_transaksi.get_children():
            self.tree_laporan_transaksi.delete(item)
        transaksi_data = self.fetch_all('''
            SELECT t.id_transaksi, t.id_pelanggan, p.nama, t.tanggal, t.status
            FROM transaksi t
            JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
            ORDER BY t.id_transaksi DESC
        ''')
        for row in transaksi_data:
            self.tree_laporan_transaksi.insert('', tk.END, values=row, tags=(row[4],))

    def export_excel(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"Laporan_Laundry_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not filename:
            return

        try:
            wb = openpyxl.Workbook()

            # Sheet Pelanggan
            ws_pelanggan = wb.active
            if ws_pelanggan is None:
                ws_pelanggan = wb.create_sheet("Pelanggan")
            else:
                ws_pelanggan.title = "Pelanggan"
            ws_pelanggan.append(["ID Pelanggan", "Nama"])
            pelanggan_data = self.fetch_all("SELECT id_pelanggan, nama FROM pelanggan ORDER BY id_pelanggan")
            for p in pelanggan_data:
                ws_pelanggan.append([p[0], p[1]])

            # Sheet Transaksi
            ws_transaksi = wb.create_sheet("Transaksi")
            ws_transaksi.append(["ID Transaksi", "ID Pelanggan", "Nama Pelanggan", "Tanggal", "Status"])
            transaksi_data = self.fetch_all('''
                SELECT t.id_transaksi, t.id_pelanggan, p.nama, t.tanggal, t.status
                FROM transaksi t
                JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
                ORDER BY t.id_transaksi DESC
            ''')
            for t in transaksi_data:
                ws_transaksi.append([t[0], t[1], t[2], t[3], t[4]])

            wb.save(filename)
            messagebox.showinfo("Sukses", f"Laporan berhasil diexport ke:\n{filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal export Excel: {e}")

    def refresh_all_data(self):
        self.refresh_pelanggan()
        self.refresh_transaksi()
        self.refresh_status_tab()
        self.refresh_laporan()

    def on_closing(self):
        self.conn.close()
        self.root.destroy()


def main():
    root = tk.Tk()
    app = LaundryApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()