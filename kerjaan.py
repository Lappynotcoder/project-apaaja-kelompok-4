import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt

# ================================
# Import Excel
# ================================
file_path = r"F:\cobaLagi\project-apaaja-kelompok-4\smvsprd.xlsx"

try:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' tidak ada!")

    data = pd.read_excel(file_path, engine="openpyxl")

    # ================================
    # Data Engineering
    # ================================
    data.columns = data.columns.str.strip()

    # Bersihkan kolom sleep_hours
    data["sleep_hours"] = pd.to_numeric(data["sleep_hours"], errors="coerce")
    data["sleep_hours"] = data["sleep_hours"].fillna(0).astype(int)

    # Pastikan jadi string
    data["social_platform_preference"] = data["social_platform_preference"].astype(str)
    data["job_type"] = data["job_type"].astype(str)

    # ================================
    # VISUAL 1 — Scatter Plot Platform
    # ================================

    # Encode job_type agar bisa dipakai sebagai nilai numerik di scatter plot
    data["job_code"] = data["job_type"].astype("category").cat.codes

    platforms = data["social_platform_preference"].unique()

    colors = {
        "Facebook": "blue",
        "Twitter": "orange",
        "Telegram": "green",
        "TikTok": "red",
        "Instagram": "purple"
    }

    plt.figure(figsize=(12, 7))

    for plat in platforms:
        subset = data[data["social_platform_preference"] == plat]
        plt.scatter(
            subset["age"],
            subset["job_code"],
            label=plat,
            color=colors.get(plat, "gray"),
            alpha=0.7,
            s=80
        )

    # Ganti angka job_code menjadi label pekerjaan
    plt.yticks(
        ticks=data["job_code"].unique(),
        labels=data["job_type"].astype('category').cat.categories
    )

    plt.title("Hubungan Pekerjaan, Umur, dan Platform Sosial Media")
    plt.xlabel("Umur")
    plt.ylabel("Jenis Pekerjaan")
    plt.legend(title="Platform")
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.tight_layout()
    plt.show()

    
    # ================================
    # VISUAL 3 — Rata-rata Durasi Sosmed per Pekerjaan
    # ================================
    avg_sosmed = (
        data.groupby("job_type")["daily_social_media_time"]
        .mean()
        .sort_values()
    )

    plt.figure(figsize=(10, 6))
    plt.bar(avg_sosmed.index, avg_sosmed.values, color="orange")
    plt.title("Rata-rata Durasi Penggunaan Sosial Media per Pekerjaan")
    plt.ylabel("Jam per Hari")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    # ================================
    # VISUAL 4 — Pie Chart Platform Sosmed
    # ================================
    platform_counts = data["social_platform_preference"].value_counts()
    platform_counts.index = platform_counts.index.astype(str)

    plt.figure(figsize=(8, 8))
    plt.pie(
        platform_counts,
        labels=platform_counts.index,
        autopct="%1.1f%%",
        startangle=90
    )

    plt.title("Distribusi Pengguna Berdasarkan Platform Sosial Media")
    plt.tight_layout()
    plt.show()

    # ================================
    # VISUAL 5 — BAR CHART RENTANG UMUR vs LEVEL STRES
    # ================================
    bins = [10, 20, 30, 40, 50, 60, 70]
    labels = ["10-20", "21-30", "31-40", "41-50", "51-60", "61-70"]

    data["umur_kelompok"] = pd.cut(data["age"], bins=bins, labels=labels, right=False)

    stress_by_age = data.groupby("umur_kelompok")["stress_level"].mean()

    plt.figure(figsize=(10, 6))
    plt.bar(stress_by_age.index.astype(str), stress_by_age.values, color="skyblue")

    plt.title("Rata-rata Level Stres Berdasarkan Rentang Umur")
    plt.xlabel("Rentang Umur")
    plt.ylabel("Level Stres")
    plt.grid(axis="y", linestyle="--", alpha=0.5)
    plt.tight_layout()
    plt.show()

except FileNotFoundError as e:
    print(f"❌ Error: {e}")

except Exception as e:
    print(f"❌ Terjadi kesalahan tak terduga: {e}")
