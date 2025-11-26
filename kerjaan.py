import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --------------------------------------------
# GENERASI FUNCTION
# --------------------------------------------
def generasi_dari_usia(age):
    tahun_lahir = 2025 - age
    if tahun_lahir >= 1997:
        return "Gen Z"
    elif 1981 <= tahun_lahir <= 1996:
        return "Milenial"
    elif 1965 <= tahun_lahir <= 1980:
        return "Gen X"
    elif 1946 <= tahun_lahir <= 1964:
        return "Baby Boomer"
    else:
        return "Lainnya"

# --------------------------------------------
# PROSES DATA
# --------------------------------------------
def proses_data():
    file_path = entry_file.get()
    if not file_path or not os.path.exists(file_path):
        messagebox.showerror("Error", "File tidak ditemukan!")
        return

    try:
        global data
        data = pd.read_excel(file_path, engine="openpyxl")

        # PROSES DATA
        data["Generasi"] = data["age"].apply(generasi_dari_usia)
        data["Sisa Waktu Luang"] = 24 - data["daily_social_media_time"] - data["work_hours_per_day"]
        data["Rasio Sosmed/Kerja"] = (data["daily_social_media_time"] / data["work_hours_per_day"]).replace([np.inf, -np.inf], np.nan)

        # Simpan output Excel
        output_file = "Data Sosial Media vs Produktifitas.xlsx"
        data.to_excel(output_file, index=False)

        label_status.config(text=f"✔ Data berhasil diproses!\nDisimpan sebagai: {output_file}")
        messagebox.showinfo("Sukses", "Data berhasil diproses dan disimpan!")

        # Tampilkan grafik default
        tampilkan_grafik()

    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {e}")

# --------------------------------------------
# TAMPILKAN GRAFIK
# --------------------------------------------
def tampilkan_grafik():
    global canvas_widget
    # Hapus canvas lama jika ada
    if 'canvas_widget' in globals():
        canvas_widget.get_tk_widget().destroy()

    fig, ax = plt.subplots(figsize=(10,4.5))
    chart_type = grafik_var.get()

    if chart_type == "Bar Chart":
        rasio_per_generasi = data.groupby("Generasi")["Rasio Sosmed/Kerja"].mean()
        rasio_per_generasi.plot(kind="bar", ax=ax, color="skyblue")
        ax.set_title("Rata-rata Rasio Sosmed/Kerja per Generasi")
        ax.set_ylabel("Rasio Sosmed/Kerja")
        ax.set_xlabel("Generasi")
        ax.set_xticklabels(rasio_per_generasi.index, rotation=45, ha='right')
        ax.grid(axis="y", linestyle="--", alpha=0.7)

    elif chart_type == "Pie Chart":
        jumlah_per_generasi = data["Generasi"].value_counts()
        ax.pie(jumlah_per_generasi, labels=jumlah_per_generasi.index, autopct="%1.1f%%", startangle=140)
        ax.set_title("Distribusi Jumlah Responden per Generasi")


    elif chart_type == "Line Chart":
        rata2_waktu = data.groupby("Generasi")["Sisa Waktu Luang"].mean()
        rata2_waktu.plot(kind="line", marker='o', ax=ax, color="green")
        ax.set_title("Rata-rata Sisa Waktu Luang per Generasi")
        ax.set_ylabel("Sisa Waktu Luang (jam)")
        ax.set_xlabel("Generasi")
        ax.grid(True, linestyle="--", alpha=0.7)

    elif chart_type == "Heatmap":
        # Ambil kolom numerik untuk korelasi
        cols = ["age", "daily_social_media_time", "work_hours_per_day", 
                "Sisa Waktu Luang", "Rasio Sosmed/Kerja"]
        
        corr = data[cols].corr()

        # Heatmap manual tanpa seaborn
        cax = ax.matshow(corr, cmap="viridis")

        fig.colorbar(cax)

        ax.set_xticks(range(len(cols)))
        ax.set_yticks(range(len(cols)))

        ax.set_xticklabels(cols, rotation=45, ha="left")
        ax.set_yticklabels(cols)

        ax.set_title("Heatmap Korelasi Variabel", pad=20)

    elif chart_type == "Sosmed per Pekerjaan":
        # Hitung rata-rata per pekerjaan & urutkan dari yang terbesar
        job_sosmed = data.groupby("job_type")["daily_social_media_time"].mean().sort_values()
        
        # plot(kind='barh') untuk horizontal
        job_sosmed.plot(kind="barh", ax=ax, color="teal")
        
        ax.set_title("Rata-rata Waktu Sosmed per Jenis Pekerjaan")
        ax.set_xlabel("Rata-rata Waktu (Jam)")
        ax.set_ylabel("Pekerjaan")
        ax.grid(axis="x", linestyle="--", alpha=0.7)


    plt.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.draw()
    canvas_widget = canvas
    canvas.get_tk_widget().pack(pady=10)

# --------------------------------------------
# SIMPAN GRAFIK SEBAGAI GAMBAR
# --------------------------------------------
def simpan_grafik():
    if 'canvas_widget' not in globals():
        messagebox.showwarning("Warning", "Tidak ada grafik untuk disimpan!")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".png",
        filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("All Files", "*.*")],
        title="Simpan Grafik Sebagai"
    )
    if file_path:
        try:
            canvas_widget.figure.savefig(file_path)
            messagebox.showinfo("Sukses", f"Grafik berhasil disimpan sebagai:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan grafik:\n{e}")

# --------------------------------------------
# PILIH FILE
# --------------------------------------------
def browse_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if filepath:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, filepath)

# --------------------------------------------
# GUI TKINTER
# --------------------------------------------
root = tk.Tk()
root.title("Data Sosial Media vs Produktivitas")
root.geometry("1280x720")
root.resizable(False, False)

# Label instruksi
label_instruksi = tk.Label(root, text="Pilih file Excel untuk diproses:")
label_instruksi.pack(pady=10)

# Frame untuk input file
frame_file = tk.Frame(root)
frame_file.pack()
entry_file = tk.Entry(frame_file, width=50)
entry_file.pack(side=tk.LEFT, padx=5)
btn_browse = tk.Button(frame_file, text="Browse", command=browse_file)
btn_browse.pack(side=tk.LEFT)

# Tombol proses
btn_proses = tk.Button(root, text="Proses Data", width=20, command=proses_data)
btn_proses.pack(pady=10)

# Pilihan grafik
grafik_var = tk.StringVar(value="Bar Chart")
options = [
    "Bar Chart", "Pie Chart", "Scatter Plot", "Line Chart", 
    "Heatmap", # Heatmap preferensi sosmed yang sebelumnya
    "Stress vs Sosmed", 
    "Sosmed per Pekerjaan", 
    "Korelasi Fitur"
]
label_grafik = tk.Label(root, text="Pilih jenis grafik:")
label_grafik.pack()
dropdown_grafik = tk.OptionMenu(root, grafik_var, *options)
dropdown_grafik.pack(pady=5)

btn_update_grafik = tk.Button(root, text="Tampilkan Grafik", command=tampilkan_grafik)
btn_update_grafik.pack(pady=5)

# Tombol simpan grafik
btn_simpan_grafik = tk.Button(root, text="Simpan Gambar Grafik", command=simpan_grafik)
btn_simpan_grafik.pack(pady=5)

# Status label
label_status = tk.Label(root, text="", fg="green")
label_status.pack()

root.mainloop()
