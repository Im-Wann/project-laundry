import sqlite3
import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime
import os
import re
import webbrowser
from urllib.parse import quote


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
                nama TEXT NOT NULL,
                nomor_wa TEXT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS transaksi (
                id_transaksi INTEGER PRIMARY KEY AUTOINCREMENT,
                id_pelanggan INTEGER NOT NULL,
                tanggal TEXT NOT NULL,
                status TEXT NOT NULL,
                berat_baju REAL DEFAULT 0,
                harga REAL DEFAULT 0,
                FOREIGN KEY (id_pelanggan) REFERENCES pelanggan (id_pelanggan)
            )
        ''')
        columns = [row[1] for row in cursor.execute("PRAGMA table_info(pelanggan)")]
        if 'nomor_wa' not in columns:
            cursor.execute("ALTER TABLE pelanggan ADD COLUMN nomor_wa TEXT")
        transaksi_columns = [row[1] for row in cursor.execute("PRAGMA table_info(transaksi)")]
        if 'berat_baju' not in transaksi_columns:
            cursor.execute("ALTER TABLE transaksi ADD COLUMN berat_baju REAL DEFAULT 0")
        if 'harga' not in transaksi_columns:
            cursor.execute("ALTER TABLE transaksi ADD COLUMN harga REAL DEFAULT 0")
        self.conn.commit()

    @staticmethod
    def normalize_phone_number(value):
        number = re.sub(r'\D', '', value)
        if number.startswith('08'):
            number = '62' + number[1:]
        elif number.startswith('8'):
            number = '62' + number
        if not re.fullmatch(r'628\d{8,11}', number):
            raise ValueError("Nomor WhatsApp harus diawali 08 dan berisi 10-13 digit")
        return number

    def hitung_harga(self):
        """Hitung harga berdasarkan berat baju (5000 per kg)"""
        try:
            berat = float(self.entry_berat_baju.get() or 0)
            harga = berat * 5000
            self.entry_harga.delete(0, tk.END)
            self.entry_harga.insert(0, str(int(harga)))
        except ValueError:
            pass

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

        ttk.Label(input_frame, text="Nomor WhatsApp:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.entry_nomor_wa = ttk.Entry(input_frame, width=40, font=('Segoe UI', 10))
        self.entry_nomor_wa.grid(row=1, column=1, sticky=tk.EW, pady=5)
        self.entry_nomor_wa.insert(0, "08")
        self.entry_nomor_wa.bind('<Return>', lambda e: self.tambah_pelanggan())

        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=0, column=2, rowspan=2, padx=(15, 0))
        ttk.Button(btn_frame, text="➕ Tambah", style='Primary.TButton', command=self.tambah_pelanggan).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 Refresh", command=self.refresh_pelanggan).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            btn_frame,
            text="Edit Pelanggan",
            command=self.edit_pelanggan
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            btn_frame,
            text="Hapus Pelanggan",
            style='Danger.TButton',
            command=self.hapus_pelanggan
        ).pack(side=tk.LEFT, padx=5)

        input_frame.columnconfigure(1, weight=1)

        # Table section
        table_frame = ttk.LabelFrame(self.tab_pelanggan, text="Daftar Pelanggan", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True)

        columns = ('ID', 'Nama', 'WhatsApp')
        self.tree_pelanggan = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='browse')
        self.tree_pelanggan.heading('ID', text='ID Pelanggan')
        self.tree_pelanggan.heading('Nama', text='Nama Pelanggan')
        self.tree_pelanggan.heading('WhatsApp', text='Nomor WhatsApp')
        self.tree_pelanggan.column('ID', width=120, anchor=tk.CENTER)
        self.tree_pelanggan.column('Nama', width=300, anchor=tk.W)
        self.tree_pelanggan.column('WhatsApp', width=180, anchor=tk.CENTER)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree_pelanggan.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree_pelanggan.xview)
        self.tree_pelanggan.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree_pelanggan.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar_y.grid(row=0, column=1, sticky=tk.NS)
        scrollbar_x.grid(row=1, column=0, sticky=tk.EW)

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)


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

        # Berat Baju
        ttk.Label(input_frame, text="Berat Baju (kg):").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.entry_berat_baju = ttk.Entry(input_frame, width=40, font=('Segoe UI', 10))
        self.entry_berat_baju.grid(row=2, column=1, sticky=tk.EW, pady=5)
        self.entry_berat_baju.bind('<KeyRelease>', lambda e: self.hitung_harga())

        # Harga
        ttk.Label(input_frame, text="Harga (Rp):").grid(row=3, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        self.entry_harga = ttk.Entry(input_frame, width=40, font=('Segoe UI', 10))
        self.entry_harga.grid(row=3, column=1, sticky=tk.EW, pady=5)

        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=0, column=2, rowspan=4, padx=(15, 0))
        ttk.Button(btn_frame, text="➕ Buat Transaksi", style='Primary.TButton', command=self.tambah_transaksi).pack(side=tk.LEFT, padx=5, pady=2)
        ttk.Button(btn_frame, text="🔄 Refresh", command=self.refresh_transaksi).pack(side=tk.LEFT, padx=5, pady=2)
        ttk.Button(btn_frame, text="Edit Transaksi", command=self.edit_transaksi).pack(side=tk.LEFT, padx=5, pady=2)
        ttk.Button(self.tab_transaksi, text="🗑️ Hapus Transaksi Terpilih", style='Danger.TButton',
                   command=self.hapus_transaksi).pack(pady=(10, 0))

        input_frame.columnconfigure(1, weight=1)

        # Table section
        table_frame = ttk.LabelFrame(self.tab_transaksi, text="Daftar Transaksi", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True)

        columns = ('No.', 'ID Transaksi', 'ID Pelanggan', 'Nama Pelanggan', 'Tanggal', 'Berat (kg)', 'Harga (Rp)', 'Status')
        self.tree_transaksi = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='browse')
        for col in columns:
            self.tree_transaksi.heading(col, text=col)
        self.tree_transaksi.column('No.', width=50, anchor=tk.CENTER)
        self.tree_transaksi.column('ID Transaksi', width=100, anchor=tk.CENTER)
        self.tree_transaksi.column('ID Pelanggan', width=100, anchor=tk.CENTER)
        self.tree_transaksi.column('Nama Pelanggan', width=150, anchor=tk.W)
        self.tree_transaksi.column('Tanggal', width=120, anchor=tk.CENTER)
        self.tree_transaksi.column('Berat (kg)', width=100, anchor=tk.CENTER)
        self.tree_transaksi.column('Harga (Rp)', width=120, anchor=tk.CENTER)
        self.tree_transaksi.column('Status', width=80, anchor=tk.CENTER)

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

        columns = ('ID Transaksi', 'ID Pelanggan', 'Nama Pelanggan', 'Tanggal', 'Berat (kg)', 'Harga (Rp)', 'Status')
        self.tree_status = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='browse')
        for col in columns:
            self.tree_status.heading(col, text=col)
        self.tree_status.column('ID Transaksi', width=100, anchor=tk.CENTER)
        self.tree_status.column('ID Pelanggan', width=100, anchor=tk.CENTER)
        self.tree_status.column('Nama Pelanggan', width=150, anchor=tk.W)
        self.tree_status.column('Tanggal', width=120, anchor=tk.CENTER)
        self.tree_status.column('Berat (kg)', width=100, anchor=tk.CENTER)
        self.tree_status.column('Harga (Rp)', width=120, anchor=tk.CENTER)
        self.tree_status.column('Status', width=80, anchor=tk.CENTER)

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
        self.stat_pendapatan = self.create_stat_card(stats_frame, "Pendapatan Selesai", "Rp 0", "#8e44ad")

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

        columns_p = ('ID', 'Nama', 'WhatsApp')
        self.tree_laporan_pelanggan = ttk.Treeview(pelanggan_frame, columns=columns_p, show='headings')
        for col in columns_p:
            self.tree_laporan_pelanggan.heading(col, text=col)
        self.tree_laporan_pelanggan.column('ID', width=100, anchor=tk.CENTER)
        self.tree_laporan_pelanggan.column('Nama', width=250, anchor=tk.W)
        self.tree_laporan_pelanggan.heading('WhatsApp', text='Nomor WhatsApp')
        self.tree_laporan_pelanggan.column('WhatsApp', width=180, anchor=tk.CENTER)

        scroll_p = ttk.Scrollbar(pelanggan_frame, orient=tk.VERTICAL, command=self.tree_laporan_pelanggan.yview)
        self.tree_laporan_pelanggan.configure(yscrollcommand=scroll_p.set)
        self.tree_laporan_pelanggan.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_p.pack(side=tk.RIGHT, fill=tk.Y)

        # Keuangan / Transaksi Selesai preview
        keuangan_frame = ttk.Frame(preview_notebook, padding=10)
        preview_notebook.add(keuangan_frame, text="Catatan Keuangan (Selesai)")

        columns_k = ('ID Transaksi', 'Tanggal', 'Nama Pelanggan', 'Berat (kg)', 'Pendapatan (Rp)')
        self.tree_laporan_keuangan = ttk.Treeview(keuangan_frame, columns=columns_k, show='headings')
        for col in columns_k:
            self.tree_laporan_keuangan.heading(col, text=col)
        self.tree_laporan_keuangan.column('ID Transaksi', width=100, anchor=tk.CENTER)
        self.tree_laporan_keuangan.column('Tanggal', width=120, anchor=tk.CENTER)
        self.tree_laporan_keuangan.column('Nama Pelanggan', width=180, anchor=tk.W)
        self.tree_laporan_keuangan.column('Berat (kg)', width=100, anchor=tk.CENTER)
        self.tree_laporan_keuangan.column('Pendapatan (Rp)', width=150, anchor=tk.E)

        scroll_k = ttk.Scrollbar(keuangan_frame, orient=tk.VERTICAL, command=self.tree_laporan_keuangan.yview)
        self.tree_laporan_keuangan.configure(yscrollcommand=scroll_k.set)
        self.tree_laporan_keuangan.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_k.pack(side=tk.RIGHT, fill=tk.Y)

        # Transaksi preview
        transaksi_frame = ttk.Frame(preview_notebook, padding=10)
        preview_notebook.add(transaksi_frame, text="Semua Transaksi")

        columns_t = ('ID Transaksi', 'ID Pelanggan', 'Nama Pelanggan', 'Tanggal', 'Berat (kg)', 'Harga (Rp)', 'Status')
        self.tree_laporan_transaksi = ttk.Treeview(transaksi_frame, columns=columns_t, show='headings')
        for col in columns_t:
            self.tree_laporan_transaksi.heading(col, text=col)
        self.tree_laporan_transaksi.column('ID Transaksi', width=100, anchor=tk.CENTER)
        self.tree_laporan_transaksi.column('ID Pelanggan', width=100, anchor=tk.CENTER)
        self.tree_laporan_transaksi.column('Nama Pelanggan', width=130, anchor=tk.W)
        self.tree_laporan_transaksi.column('Tanggal', width=120, anchor=tk.CENTER)
        self.tree_laporan_transaksi.column('Berat (kg)', width=100, anchor=tk.CENTER)
        self.tree_laporan_transaksi.column('Harga (Rp)', width=120, anchor=tk.CENTER)
        self.tree_laporan_transaksi.column('Status', width=80, anchor=tk.CENTER)

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
            nomor_wa = self.normalize_phone_number(self.entry_nomor_wa.get())
            self.execute_query(
                "INSERT INTO pelanggan (nama, nomor_wa) VALUES (?, ?)",
                (nama, nomor_wa)
            )
            self.entry_nama.delete(0, tk.END)
            self.entry_nomor_wa.delete(0, tk.END)
            self.entry_nomor_wa.insert(0, "08")
            self.refresh_pelanggan()
            messagebox.showinfo("Sukses", f"Pelanggan '{nama}' berhasil ditambahkan!")
        except ValueError as e:
            messagebox.showerror("Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menambah pelanggan: {e}")

    def edit_pelanggan(self):
        selection = self.tree_pelanggan.selection()
        if not selection:
            messagebox.showwarning("Peringatan", "Pilih pelanggan yang ingin diedit!")
            return

        values = self.tree_pelanggan.item(selection[0], 'values')
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Pelanggan")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Nama pelanggan:").grid(row=0, column=0, padx=10, pady=8, sticky=tk.W)
        nama_entry = ttk.Entry(dialog, width=35)
        nama_entry.insert(0, values[1])
        nama_entry.grid(row=0, column=1, padx=10, pady=8)

        ttk.Label(dialog, text="Nomor WhatsApp:").grid(row=1, column=0, padx=10, pady=8, sticky=tk.W)
        nomor_entry = ttk.Entry(dialog, width=35)
        nomor_entry.insert(0, values[2])
        nomor_entry.grid(row=1, column=1, padx=10, pady=8)

        def simpan():
            nama = nama_entry.get().strip()
            if not nama:
                messagebox.showwarning("Peringatan", "Nama pelanggan tidak boleh kosong!", parent=dialog)
                return
            try:
                nomor_wa = self.normalize_phone_number(nomor_entry.get())
                self.execute_query(
                    "UPDATE pelanggan SET nama = ?, nomor_wa = ? WHERE id_pelanggan = ?",
                    (nama, nomor_wa, values[0])
                )
                dialog.destroy()
                self.refresh_all_data()
            except ValueError as e:
                messagebox.showerror("Error", str(e), parent=dialog)
            except Exception as e:
                messagebox.showerror("Error", f"Gagal mengedit pelanggan: {e}", parent=dialog)

        ttk.Button(dialog, text="Simpan", command=simpan).grid(row=2, column=1, padx=10, pady=10, sticky=tk.E)
        ttk.Button(dialog, text="Batal", command=dialog.destroy).grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        nama_entry.focus_set()
        dialog.wait_window()


    def hapus_pelanggan(self):
        selection = self.tree_pelanggan.selection()
        if not selection:
            messagebox.showwarning("Peringatan", "Pilih pelanggan yang ingin dihapus!")
            return

        item = self.tree_pelanggan.item(selection[0])
        id_pelanggan = item['values'][0]
        nama = item['values'][1]

        if not messagebox.askyesno(
            "Konfirmasi",
            f"Hapus pelanggan '{nama}' (ID: {id_pelanggan})?\n"
            "Semua transaksi terkait juga akan terhapus."
        ):
            return

        try:
            self.execute_query(
                "DELETE FROM transaksi WHERE id_pelanggan = ?",
                (id_pelanggan,)
            )
            self.execute_query(
                "DELETE FROM pelanggan WHERE id_pelanggan = ?",
                (id_pelanggan,)
            )
            if self.fetch_one("SELECT COUNT(*) FROM pelanggan")[0] == 0:
                self.execute_query("DELETE FROM sqlite_sequence WHERE name = 'pelanggan'")
            self.refresh_all_data()
            messagebox.showinfo("Sukses", "Pelanggan dan transaksi terkait berhasil dihapus!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menghapus pelanggan: {e}")

    def refresh_pelanggan(self):
        # Clear tree
        for item in self.tree_pelanggan.get_children():
            self.tree_pelanggan.delete(item)

        # Fetch and populate
        data = self.fetch_all("SELECT id_pelanggan, nama, nomor_wa FROM pelanggan ORDER BY id_pelanggan")
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

        # Get berat and harga
        try:
            berat_baju = float(self.entry_berat_baju.get() or 0)
            harga = float(self.entry_harga.get() or 0)
        except ValueError:
            messagebox.showerror("Error", "Berat atau harga harus berupa angka!")
            return

        try:
            self.execute_query(
                "INSERT INTO transaksi (id_pelanggan, tanggal, status, berat_baju, harga) VALUES (?, ?, ?, ?, ?)",
                (id_pelanggan, tanggal, "PROSES", berat_baju, harga)
            )
            self.entry_tanggal.delete(0, tk.END)
            self.entry_tanggal.insert(0, datetime.now().strftime('%Y-%m-%d'))
            self.entry_berat_baju.delete(0, tk.END)
            self.entry_harga.delete(0, tk.END)
            self.refresh_transaksi()
            messagebox.showinfo("Sukses", "Transaksi baru berhasil dibuat dengan status PROSES!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuat transaksi: {e}")

    def edit_transaksi(self):
        selection = self.tree_transaksi.selection()
        if not selection:
            messagebox.showwarning("Peringatan", "Pilih transaksi yang ingin diedit!")
            return

        values = self.tree_transaksi.item(selection[0], 'values')
        tanggal = simpledialog.askstring(
            "Edit Transaksi", "Tanggal (YYYY-MM-DD):", initialvalue=values[3]
        )
        if tanggal is None:
            return
        try:
            datetime.strptime(tanggal.strip(), '%Y-%m-%d')
            self.execute_query(
                "UPDATE transaksi SET tanggal = ? WHERE id_transaksi = ?",
                (tanggal.strip(), values[0])
            )
            self.refresh_all_data()
        except ValueError:
            messagebox.showerror("Error", "Format tanggal salah! Gunakan YYYY-MM-DD")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengedit transaksi: {e}")

    def hapus_transaksi(self):
        selection = self.tree_transaksi.selection()
        if not selection:
            messagebox.showwarning("Peringatan", "Pilih transaksi yang ingin dihapus!")
            return

        item = self.tree_transaksi.item(selection[0])
        vals = item['values']
        # values format: (No, id_transaksi, id_pelanggan, nama, tanggal, berat, harga, status)
        id_transaksi = vals[1]
        nama = vals[3]

        if not messagebox.askyesno(
            "Konfirmasi",
            f"Hapus transaksi ID {id_transaksi} milik '{nama}'?"
        ):
            return

        try:
            self.execute_query(
                "DELETE FROM transaksi WHERE id_transaksi = ?",
                (id_transaksi,)
            )
            self.refresh_all_data()
            messagebox.showinfo("Sukses", "Transaksi berhasil dihapus!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menghapus transaksi: {e}")

    def refresh_transaksi(self):
        for item in self.tree_transaksi.get_children():
            self.tree_transaksi.delete(item)

        data = self.fetch_all('''
            SELECT t.id_transaksi, t.id_pelanggan, p.nama, t.tanggal, t.berat_baju, t.harga, t.status
            FROM transaksi t
            JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
            ORDER BY t.id_transaksi DESC
        ''')
        for nomor, row in enumerate(data, start=1):
            self.tree_transaksi.insert('', tk.END, values=(nomor, *row), tags=(row[6],))

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
            old_status = self.fetch_one(
                "SELECT status FROM transaksi WHERE id_transaksi = ?",
                (id_transaksi,)
            )
            self.execute_query("UPDATE transaksi SET status = ? WHERE id_transaksi = ?", (new_status, id_transaksi))
            self.refresh_status_tab()
            if old_status and old_status[0] != "SELESAI" and new_status == "SELESAI":
                self.open_whatsapp_message(id_transaksi)
            messagebox.showinfo("Sukses", f"Status transaksi ID {id_transaksi} berhasil diupdate ke {new_status}!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal update status: {e}")

    def refresh_status_tab(self):
        for item in self.tree_status.get_children():
            self.tree_status.delete(item)

        data = self.fetch_all('''
            SELECT t.id_transaksi, t.id_pelanggan, p.nama, t.tanggal, t.berat_baju, t.harga, t.status
            FROM transaksi t
            JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
            ORDER BY t.id_transaksi DESC
        ''')
        for row in data:
            self.tree_status.insert('', tk.END, values=row, tags=(row[6],))

        self.update_transaksi_combobox()

    def open_whatsapp_message(self, id_transaksi):
        data = self.fetch_one('''
            SELECT p.nomor_wa, p.nama, t.tanggal
            FROM transaksi t
            JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
            WHERE t.id_transaksi = ?
        ''', (id_transaksi,))
        if not data or not data[0]:
            messagebox.showwarning("WhatsApp", "Nomor WhatsApp pelanggan belum tersedia.")
            return

        nomor_wa, nama, tanggal = data
        pesan = (
            f"Halo {nama}, laundry Anda pada tanggal {tanggal} "
            "sudah selesai dan dapat diambil. Terima kasih."
        )
        url = f"https://web.whatsapp.com/send?phone={nomor_wa}&text={quote(pesan)}"
        webbrowser.open(url)

    # --- Laporan & Export ---
    def refresh_laporan(self):
        # Stats
        total_pelanggan = self.fetch_one("SELECT COUNT(*) FROM pelanggan")[0]
        total_transaksi = self.fetch_one("SELECT COUNT(*) FROM transaksi")[0]
        proses = self.fetch_one("SELECT COUNT(*) FROM transaksi WHERE status = 'PROSES'")[0]
        selesai = self.fetch_one("SELECT COUNT(*) FROM transaksi WHERE status = 'SELESAI'")[0]
        batal = self.fetch_one("SELECT COUNT(*) FROM transaksi WHERE status = 'BATAL'")[0]
        total_pendapatan = self.fetch_one("SELECT SUM(harga) FROM transaksi WHERE status = 'SELESAI'")[0] or 0

        self.stat_total_pelanggan.config(text=str(total_pelanggan))
        self.stat_total_transaksi.config(text=str(total_transaksi))
        self.stat_proses.config(text=str(proses))
        self.stat_selesai.config(text=str(selesai))
        self.stat_batal.config(text=str(batal))
        self.stat_pendapatan.config(text=f"Rp {total_pendapatan:,.0f}")

        # Pelanggan table
        for item in self.tree_laporan_pelanggan.get_children():
            self.tree_laporan_pelanggan.delete(item)
        pelanggan_data = self.fetch_all("SELECT id_pelanggan, nama, nomor_wa FROM pelanggan ORDER BY id_pelanggan")
        for row in pelanggan_data:
            self.tree_laporan_pelanggan.insert('', tk.END, values=row)

        # Keuangan (Selesai) table
        for item in self.tree_laporan_keuangan.get_children():
            self.tree_laporan_keuangan.delete(item)
        keuangan_data = self.fetch_all('''
            SELECT t.id_transaksi, t.tanggal, p.nama, t.berat_baju, t.harga
            FROM transaksi t
            JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
            WHERE t.status = 'SELESAI'
            ORDER BY t.id_transaksi DESC
        ''')
        for row in keuangan_data:
            formatted_harga = f"Rp {row[4]:,.0f}" if row[4] else "Rp 0"
            self.tree_laporan_keuangan.insert('', tk.END, values=(row[0], row[1], row[2], row[3], formatted_harga))

        # Transaksi table
        for item in self.tree_laporan_transaksi.get_children():
            self.tree_laporan_transaksi.delete(item)
        transaksi_data = self.fetch_all('''
            SELECT t.id_transaksi, t.id_pelanggan, p.nama, t.tanggal, t.berat_baju, t.harga, t.status
            FROM transaksi t
            JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
            ORDER BY t.id_transaksi DESC
        ''')
        for row in transaksi_data:
            self.tree_laporan_transaksi.insert('', tk.END, values=row, tags=(row[6],))

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
            ws_pelanggan.append(["ID Pelanggan", "Nama", "Nomor WhatsApp"])
            pelanggan_data = self.fetch_all("SELECT id_pelanggan, nama, nomor_wa FROM pelanggan ORDER BY id_pelanggan")
            for p in pelanggan_data:
                ws_pelanggan.append([p[0], p[1], p[2]])

            # Sheet Transaksi
            ws_transaksi = wb.create_sheet("Transaksi")
            ws_transaksi.append(["ID Transaksi", "ID Pelanggan", "Nama Pelanggan", "Tanggal", "Berat Baju (kg)", "Harga (Rp)", "Status"])
            transaksi_data = self.fetch_all('''
                SELECT t.id_transaksi, t.id_pelanggan, p.nama, t.tanggal, t.berat_baju, t.harga, t.status
                FROM transaksi t
                JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
                ORDER BY t.id_transaksi DESC
            ''')
            for t in transaksi_data:
                ws_transaksi.append([t[0], t[1], t[2], t[3], t[4], t[5], t[6]])

            # Sheet Keuangan (Selesai)
            ws_keuangan = wb.create_sheet("Keuangan Selesai")
            ws_keuangan.append(["ID Transaksi", "Tanggal", "Nama Pelanggan", "Berat Baju (kg)", "Pendapatan (Rp)"])
            keuangan_data = self.fetch_all('''
                SELECT t.id_transaksi, t.tanggal, p.nama, t.berat_baju, t.harga
                FROM transaksi t
                JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan
                WHERE t.status = 'SELESAI'
                ORDER BY t.id_transaksi DESC
            ''')
            for k in keuangan_data:
                ws_keuangan.append([k[0], k[1], k[2], k[3], k[4]])

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