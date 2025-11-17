import pandas as pd

#1. Import data
df = pd.read_excel(r"C:\Users\ACER\OneDrive\Documents\DASAR PEMROGRAMAN\BARU\project-apaaja-kelompok-4\smvsprd.xlsx",engine="openpyxl")

#2. Rename kolom menjadi konsisten
df = df.rename(columns={
    "age": "age",
    "Age": "age",
    "job": "job_type",
    "job type": "job_type",
    "job_type": "job_type",
    "daily social media time": "daily_social_media_time",
    "Daily social media time": "daily_social_media_time",
    "social platform preference": "social_platform_preference",
    "preferred social platform": "social_platform_preference",
    "work hours per day": "work_hours_per_day",
    "hours_worked": "work_hours_per_day",
    "stress level": "stress_level",
    "stress_level": "stress_level"
})

#3. mengambil kolom yang hanya dibutuhkan
required_columns = [
    "age",
    "job_type",
    "daily_social_media_time",
    "social_platform_preference",
    "work_hours_per_day",
    "stress_level"
]
df = df[required_columns]

#4. Hapus baris yang ada data kosong ---
df = df.dropna()

#5. Simpan sebagai file Excel baru ---
df.to_excel("data_fix.xlsx", index=False)

print("Proses selesai! File berhasil dibuat sebagai data_fix.xlsx")