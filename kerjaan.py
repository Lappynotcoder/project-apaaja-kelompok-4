import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --------------------------------------------
# FUNGSI GENERASI & PROSES DATA
# --------------------------------------------
def generasi_dari_usia(age):
    tahun_lahir = 2025 - age
    if tahun_lahir >= 1997: return "Gen Z"
    elif 1981 <= tahun_lahir <= 1996: return "Milenial"
    elif 1965 <= tahun_lahir <= 1980: return "Gen X"
    elif 1946 <= tahun_lahir <= 1964: return "Baby Boomer"
    else: return "Lainnya"

def proses_data():
    file_path = entry_file.get()
    if not file_path or not os.path.exists(file_path):
        messagebox.showerror("Error", "File tidak ditemukan!")
        return

    try:
        global data
        data = pd.read_excel(file_path, engine="openpyxl")
        
        # Deteksi nama kolom agar fleksibel (sesuai file yang kamu upload)
        # Kita cari kolom yang mirip 'social_platform' atau 'preferred'
        global col_pref
        if 'social_platform_preference' in data.columns:
            col_pref = 'social_platform_preference'
        elif 'preferred_social_media_platform' in data.columns:
            col_pref = 'preferred_social_media_platform'
        else:
            # Fallback kalau kolom tidak ketemu, pakai kolom ke-4 (index 3)
            col_pref = data.columns[3] 

        col_sosmed = 'daily_social_media_time'
        col_work = 'work_hours_per_day'
        
        # PROSES DATA
        data["Generasi"] = data["age"].apply(generasi_dari_usia)
        data["Sisa Waktu Luang"] = 24 - data[col_sosmed] - data[col_work]
        
        # Simpan output
        output_file = "Data Final.xlsx"
        data.to_excel(output_file, index=False)
        
        label_status.config(text=f"✔ Data Siap! Disimpan ke: {output_file}")
        messagebox.showinfo("Sukses", "Data berhasil diproses!")
        tampilkan_grafik()

    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {e}\nPastikan file Excel benar.")

# --------------------------------------------
# TAMPILKAN GRAFIK
# --------------------------------------------
def tampilkan_grafik():
    global canvas_widget
    if 'canvas_widget' in globals():
        canvas_widget.get_tk_widget().destroy()

    # Ukuran grafik
    fig, ax = plt.subplots(figsize=(10, 5))
    chart_type = grafik_var.get()

    # 1. RATA-RATA SOSMED PER GENERASI
    if chart_type == "Sosmed per Generasi":
        avg_gen = data.groupby("Generasi")["daily_social_media_time"].mean()
        bars = avg_gen.plot(kind="bar", ax=ax, color="skyblue", rot=0)
        
        ax.set_title("Rata-rata Waktu Main Sosmed per Generasi")
        ax.set_ylabel("Durasi (Jam/Hari)")
        ax.set_xlabel("")
        ax.grid(axis="y", linestyle="--", alpha=0.5)
        
        # Label angka di atas batang
        for p in ax.patches:
            ax.annotate(f'{p.get_height():.1f} Jam', 
                        (p.get_x() + p.get_width() / 2., p.get_height()), 
                        ha='center', va='center', xytext=(0, 5), textcoords='offset points')

    # 2. RATA-RATA SOSMED PER PEKERJAAN
    elif chart_type == "Sosmed per Pekerjaan":
        avg_job = data.groupby("job_type")["daily_social_media_time"].mean().sort_values()
        avg_job.plot(kind="barh", ax=ax, color="salmon")
        ax.set_title("Pekerjaan dengan Waktu Sosmed Tertinggi")
        ax.set_xlabel("Rata-rata Durasi (Jam)")
        ax.grid(axis="x", linestyle="--", alpha=0.5)

    # 3. LINE CHART SISA WAKTU LUANG
    elif chart_type == "Line: Sisa Waktu Luang":
        rata2_waktu = data.groupby("Generasi")["Sisa Waktu Luang"].mean()
        rata2_waktu.plot(kind="line", marker='o', ax=ax, color="green", linewidth=2)
        ax.set_title("Tren Sisa Waktu Luang per Generasi")
        ax.set_ylabel("Jam Luang")
        ax.grid(True, linestyle="--", alpha=0.5)

    # 4. STRESS VS SOSMED (SIMPEL)
    elif chart_type == "Stress vs Sosmed":
        avg_stress = data.groupby("stress_level")["daily_social_media_time"].mean()
        avg_stress.plot(kind="bar", ax=ax, color="orange", rot=0)
        ax.set_title("Hubungan Tingkat Stress & Lama Main Sosmed")
        ax.set_ylabel("Rata-rata Waktu Sosmed (Jam)")
        ax.set_xlabel("Tingkat Stress (1-10)")
        ax.grid(axis="y", linestyle="--", alpha=0.5)

    # 5. KOMPARASI KERJA VS SOSMED
    elif chart_type == "Komparasi: Kerja vs Sosmed":
        komparasi = data.groupby("Generasi")[["work_hours_per_day", "daily_social_media_time"]].mean()
        komparasi.plot(kind="bar", ax=ax, color=["#2c3e50", "#e74c3c"], rot=0)
        ax.set_title("Perbandingan: Jam Kerja vs Jam Sosmed")
        ax.set_ylabel("Durasi (Jam)")
        ax.legend(["Jam Kerja", "Jam Sosmed"])
        ax.grid(axis="y", linestyle="--", alpha=0.5)

    # 6. PIE CHART: GENERASI (RESPONDEN)
    elif chart_type == "Pie Chart: Generasi":
        jumlah = data["Generasi"].value_counts()
        ax.pie(jumlah, labels=jumlah.index, autopct="%1.1f%%", startangle=140, colors=plt.cm.Pastel1.colors)
        ax.set_title("Komposisi Responden per Generasi")

    # 7. PIE CHART: PLATFORM SOSMED (YANG KAMU MINTA)
    elif chart_type == "Pie Chart: Platform Sosmed":
        # Hitung jumlah pengguna per platform
        platform_counts = data[col_pref].value_counts()
        
        # Gambar Pie Chart
        # autopct='%1.1f%%' fungsinya menampilkan persentase (1 desimal)
        ax.pie(platform_counts, labels=platform_counts.index, autopct="%1.1f%%", 
               startangle=90, colors=plt.cm.Set3.colors, pctdistance=0.85)
        
        # Tambahkan lingkaran putih di tengah biar jadi Donut Chart (lebih modern)
        centre_circle = plt.Circle((0,0),0.70,fc='white')
        fig.gca().add_artist(centre_circle)
        
        ax.set_title("Distribusi Pengguna per Platform Sosmed")
        ax.axis('equal')  # Agar lingkaran sempurna

    plt.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.draw()
    canvas_widget = canvas
    canvas.get_tk_widget().pack(pady=10)

# --------------------------------------------
# FUNGSI LAINNYA
# --------------------------------------------
def simpan_grafik():
    if 'canvas_widget' not in globals(): return
    file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png")])
    if file_path: canvas_widget.figure.savefig(file_path)

def browse_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if filepath:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, filepath)

# GUI SETUP
root = tk.Tk()
root.title("Analisis Produktivitas & Sosmed")
root.geometry("1280x720")
root.resizable(False, False)

tk.Label(root, text="Analisis Data: Sosmed vs Produktivitas", font=("Arial", 16, "bold")).pack(pady=10)

frame_input = tk.Frame(root)
frame_input.pack(pady=5)
entry_file = tk.Entry(frame_input, width=50)
entry_file.pack(side=tk.LEFT, padx=5)
tk.Button(frame_input, text="Pilih File Excel", command=browse_file).pack(side=tk.LEFT)
tk.Button(root, text="PROSES DATA", bg="lightblue", command=proses_data).pack(pady=5)

# MENU PILIHAN GRAFIK
grafik_var = tk.StringVar(value="Sosmed per Generasi")
options = [
    "Sosmed per Generasi", 
    "Sosmed per Pekerjaan", 
    "Line: Sisa Waktu Luang",
    "Stress vs Sosmed",
    "Komparasi: Kerja vs Sosmed",
    "Pie Chart: Generasi",
    "Pie Chart: Platform Sosmed"
]

tk.Label(root, text="Pilih Visualisasi:").pack()
tk.OptionMenu(root, grafik_var, *options).pack(pady=5)
tk.Button(root, text="Tampilkan Grafik", command=tampilkan_grafik).pack(pady=5)
tk.Button(root, text="Simpan Gambar", command=simpan_grafik).pack(pady=5)
label_status = tk.Label(root, text="", fg="green")
label_status.pack()

root.mainloop()