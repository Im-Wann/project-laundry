import sqlite3


conn = sqlite3.connect('laundry.db')

cursor = conn.cursor()

#fitur foreign key
cursor.execute('PRAGMA foreign_keys = ON')

#buat tabel pelanggan dan transaksi jika belum ada
cursor.execute('''
    CREATE TABLE IF NOT EXISTS pelanggan (
        id_pelanggan INTEGER PRIMARY KEY AUTOINCREMENT,
        nama TEXT NOT NULL
    )''')
print("Tabel pelanggan berhasil dibuat.")

cursor.execute('''
    CREATE TABLE IF NOT EXISTS transaksi (
        id_transaksi INTEGER PRIMARY KEY AUTOINCREMENT,
        id_pelanggan INTEGER NOT NULL,
        tanggal TEXT NOT NULL,
        status TEXT NOT NULL,
        FOREIGN KEY (id_pelanggan) REFERENCES pelanggan (id_pelanggan)
    )''')
print("Tabel transaksi berhasil dibuat.")

def layanan():
    print('''
============================================
        Selamat Datang di Laundry Kami
    Silakan Pilih Layanan yang tersedia:
============================================
|   No  |           Layanan Laundry         | 
---------------------------------------------
|   1.  |        Pelanggan Baru             |
|   2.  |        Transaksi Baru             |
|   3.  |        Update Status Laundry      |
|   4.  |        Lihat Laporan Laundry      |
|   5.  |        Keluar                     |
=============================================
''')


while True:
    layanan()
    Pilihan = int(input("Masukkan Pilihan Anda (1-5): "))
    
    if Pilihan == 1:
        print('''
============================================
    Anda Memilih Layanan Pelanggan Baru
============================================''')
        Pelanggan_baru = str(input("Masukkan Nama Pelanggan: "))
        cursor.execute("INSERT INTO pelanggan (nama) VALUES (?)", (Pelanggan_baru,))
        conn.commit()
        #conn.commit untuk menyimpan perubahan ke database
        print("Pelanggan baru berhasil ditambahkan.")
    elif Pilihan == 2:
        print('''
============================================
    Anda Memilih Layanan Transaksi Baru
============================================''')
        cursor.execute("SELECT id_pelanggan, nama FROM pelanggan")
        list_pelanggan = cursor.fetchall()
        if not list_pelanggan:
            print("Belum ada pelanggan yang terdaftar. Silakan tambahkan pelanggan terlebih dahulu.")
            continue
        for pelanggan in list_pelanggan:
            print(f"ID: {pelanggan[0]}, Nama: {pelanggan[1]}")
        id_pelanggan = int(input("Masukkan ID Pelanggan: "))

        #cek id pelanggan, jika tidak ada maka akan menampilkan pesan error dan meminta input ulang
        if id_pelanggan not in [pelanggan[0] for pelanggan in list_pelanggan]:
            print("ID Pelanggan tidak valid. Silakan coba lagi.")
            continue
        tanggal = str(input("Masukkan Tanggal Transaksi (YYYY-MM-DD): "))
        status = "PROSES"
        cursor.execute("INSERT INTO transaksi (id_pelanggan, tanggal, status) VALUES (?, ?, ?)", (id_pelanggan, tanggal, status))
        conn.commit()
        print("Transaksi baru berhasil ditambahkan.")
        
    elif Pilihan == 3:
        print('''
===================================================
    Anda Memilih Layanan Update Status Laundry
===================================================''')
        cursor.execute('''
            SELECT 
                transaksi.id_transaksi, 
                transaksi.id_pelanggan, 
                pelanggan.nama, 
                transaksi.tanggal, 
                transaksi.status 
            FROM transaksi
            JOIN pelanggan ON transaksi.id_pelanggan = pelanggan.id_pelanggan
        ''')   
        list_transaksi = cursor.fetchall()
        if not list_transaksi:
            print("ID Transaksi kosong atau belum ada.")
            continue
        for transaksi in list_transaksi:
            print(f"ID Transaksi: {transaksi[0]}, ID Pelanggan: {transaksi[1]}, Nama Pelanggan: {transaksi[2]}, Tanggal: {transaksi[3]}, Status: {transaksi[4]}")
            id_transaksi = int(input("Masukkan ID Transaksi yang ingin di Update: "))
            if id_transaksi not in [transaksi[0] for transaksi in list_transaksi]:
                print("ID Transaksi tidak valid. Silakan coba lagi.")
                continue
            print('''
Pilih Status Laundry:
1. SELESAI
2. BATAL''')
            while True:
                update_status = int(input("Masukkan Pilihan Status (1-2):"))
                if update_status == 1:
                    cursor.execute("UPDATE transaksi SET status = 'SELESAI' WHERE id_transaksi = ?", (id_transaksi,))
                    conn.commit()
                    print("Status transaksi berhasil diupdate.")
                    break
                elif update_status == 2:
                    cursor.execute("UPDATE transaksi SET status = 'BATAL' WHERE id_transaksi = ?", (id_transaksi,))
                    conn.commit()
                    print("Status transaksi berhasil dibatalkan.")  
                    break
                else:
                    print("Pilihan tidak valid. Silakan pilih antara 1 hingga 2.")
    elif Pilihan == 4:
        print('''''')
    elif Pilihan == 5:
        print("Terima kasih telah menggunakan layanan Laundry Kami. Sampai jumpa!")
        break
    else:
        print("Pilihan tidak valid. Silakan pilih antara 1 hingga 5.")
        




        