import pandas as pd
import os
import numpy as np
import matplotlib as plt

#import excel
file_path=r"smvsprd.xlsx"
output_file="Data Sosial Media vs Produktifitas"

try:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' tidak ada!")
    
    data=pd.read_excel(file_path,engine="openpyxl")

    #Data Engineer mulai disini



    #Programmer mulai disini



    #Visual mulai disini
    


except FileNotFoundError as e:
    print(f"❌ Error: {e}")

except KeyError as e:
    print(f"❌ Error: Kolom tidak ditemukan: {e}")

except ValueError as e:
    print(f"❌ Error: {e}")

except Exception as e:
    print(f"❌ Terjadi kesalahan tak terduga: {e}")