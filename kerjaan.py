import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import pandas as pd
import numpy as np
import threading
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

# Translator
# pip install googletrans==4.0.0-rc1
try:
    from googletrans import Translator
    translator = Translator()
except Exception:
    translator = None

# ----------------------------
# Constants & language list
# ----------------------------
LANG_LIST = [
    ("Afrikaans","af"),("Albanian","sq"),("Amharic","am"),("Arabic","ar"),
    ("Armenian","hy"),("Azerbaijani","az"),("Basque","eu"),("Belarusian","be"),
    ("Bengali","bn"),("Bosnian","bs"),("Bulgarian","bg"),("Catalan","ca"),
    ("Cebuano","ceb"),("Chinese Simplified","zh-cn"),("Chinese Traditional","zh-tw"),
    ("Corsican","co"),("Croatian","hr"),("Czech","cs"),("Danish","da"),
    ("Dutch","nl"),("English","en"),("Esperanto","eo"),("Estonian","et"),
    ("Finnish","fi"),("French","fr"),("Frisian","fy"),("Galician","gl"),
    ("Georgian","ka"),("German","de"),("Greek","el"),("Gujarati","gu"),
    ("Haitian Creole","ht"),("Hausa","ha"),("Hawaiian","haw"),("Hebrew","iw"),
    ("Hindi","hi"),("Hmong","hmn"),("Hungarian","hu"),("Icelandic","is"),
    ("Igbo","ig"),("Indonesian","id"),("Irish","ga"),("Italian","it"),
    ("Japanese","ja"),("Javanese","jv"),("Kannada","kn"),("Kazakh","kk"),
    ("Khmer","km"),("Kinyarwanda","rw"),("Korean","ko"),("Kurdish","ku"),
    ("Kyrgyz","ky"),("Lao","lo"),("Latin","la"),("Latvian","lv"),
    ("Lithuanian","lt"),("Luxembourgish","lb"),("Macedonian","mk"),
    ("Malagasy","mg"),("Malay","ms"),("Malayalam","ml"),("Maltese","mt"),
    ("Maori","mi"),("Marathi","mr"),("Mongolian","mn"),("Myanmar","my"),
    ("Nepali","ne"),("Norwegian","no"),("Nyanja","ny"),("Odia","or"),
    ("Pashto","ps"),("Persian","fa"),("Polish","pl"),("Portuguese","pt"),
    ("Punjabi","pa"),("Romanian","ro"),("Russian","ru"),("Samoan","sm"),
    ("Scots Gaelic","gd"),("Serbian","sr"),("Sesotho","st"),("Shona","sn"),
    ("Sindhi","sd"),("Sinhala","si"),("Slovak","sk"),("Slovenian","sl"),
    ("Somali","so"),("Spanish","es"),("Sundanese","su"),("Swahili","sw"),
    ("Swedish","sv"),("Tagalog","tl"),("Tajik","tg"),("Tamil","ta"),
    ("Tatar","tt"),("Telugu","te"),("Thai","th"),("Turkish","tr"),
    ("Turkmen","tk"),("Ukrainian","uk"),("Urdu","ur"),("Uyghur","ug"),
    ("Uzbek","uz"),("Vietnamese","vi"),("Welsh","cy"),("Xhosa","xh"),
    ("Yiddish","yi"),("Yoruba","yo"),("Zulu","zu")
]
LANG_DISPLAY = [f"{name} ({code})" for name, code in LANG_LIST]
LANG_CODES = {f"{name} ({code})": code for name, code in LANG_LIST}

# Chart options
CHART_OPTIONS = [
    "Scatter Plot",
    "Line Chart",
    "Bar Chart (Category vs Numeric)",
    "Histogram",
    "Pie Chart",
    "Heatmap (Correlation)",
    "Box Plot",
    "Area Chart",
    "All Charts"
]

# Quick-load sample path (file you uploaded earlier)
SAMPLE_PATH = "/mnt/data/smvsprd.xlsx"

# Globals
current_df = None
last_fig = None
is_dark_mode = False
chart_widgets = []  # Store chart widgets for cleanup

# ----------------------------
# Utility functions
# ----------------------------
def safe_to_str(x):
    return "" if pd.isna(x) else str(x)

def add_waktu_rasio(df):
    if {"daily_social_media_time", "work_hours_per_day"}.issubset(df.columns):
        try:
            df["Waktu Luang"] = 24 - df["daily_social_media_time"] - df["work_hours_per_day"]
        except Exception:
            df["Waktu Luang"] = np.nan
        try:
            df["Rasio Sosmed/Kerja"] = (df["daily_social_media_time"] / df["work_hours_per_day"]).replace([np.inf, -np.inf], np.nan)
        except Exception:
            df["Rasio Sosmed/Kerja"] = np.nan
    else:
        if "Waktu Luang" not in df.columns:
            df["Waktu Luang"] = "N/A"
        if "Rasio Sosmed/Kerja" not in df.columns:
            df["Rasio Sosmed/Kerja"] = "N/A"
    return df

# ----------------------------
# GUI: Data display
# ----------------------------
def tampilkan_dataframe(df):
    for widget in frame_table.winfo_children():
        widget.destroy()

    tree = ttk.Treeview(frame_table, show="headings")
    vsb = ttk.Scrollbar(frame_table, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(frame_table, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    cols = list(df.columns)
    tree["columns"] = cols
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor="center")

    for i, row in df.iterrows():
        values = [safe_to_str(v) for v in list(row)]
        tree.insert("", "end", values=values)

    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    frame_table.grid_rowconfigure(0, weight=1)
    frame_table.grid_columnconfigure(0, weight=1)

    update_column_selectors()

# ----------------------------
# Load Excel
# ----------------------------
def load_excel():
    global current_df
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not filepath:
        return
    try:
        df = pd.read_excel(filepath, engine="openpyxl")
        df = add_waktu_rasio(df)
        current_df = df
        tampilkan_dataframe(df)
        label_status.config(text=f"✔ Loaded: {filepath}", fg= "green")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal memuat file:\n{e}")
        label_status.config(text="✘ Gagal memuat file", fg="red")

def quick_load():
    global current_df
    try:
        df = pd.read_excel(SAMPLE_PATH, engine="openpyxl")
        df = add_waktu_rasio(df)
        current_df = df
        tampilkan_dataframe(df)
        label_status.config(text=f"✔ Sample loaded: {SAMPLE_PATH}", fg="green")
    except Exception as e:
        messagebox.showerror("Error", f"Quick load gagal:\n{e}")

# ----------------------------
# Translate entire DataFrame -> new columns
# ----------------------------
def translate_text_cell(text, dest):
    if not isinstance(text, str) or translator is None:
        return text
    try:
        res = translator.translate(text, dest=dest)
        return res.text
    except Exception:
        return text

def translate_entire_df_thread():
    threading.Thread(target=translate_entire_df, daemon=True).start()

def translate_entire_df():
    global current_df
    if current_df is None:
        messagebox.showwarning("Peringatan", "Muat data dulu.")
        return
    lang = combo_lang.get()
    if lang not in LANG_CODES:
        messagebox.showwarning("Peringatan", "Pilih bahasa target.")
        return
    code = LANG_CODES[lang]
    df = current_df.copy()

    text_cols = [c for c in df.columns if df[c].dtype == object]
    if not text_cols:
        messagebox.showinfo("Info", "Tidak ada kolom teks untuk diterjemahkan.")
        return

    btn_translate.config(state=tk.DISABLED)
    label_status.config(text="🕗 Mentranslate... tunggu", fg="orange")
    root.update_idletasks()

    try:
        for col in text_cols:
            new_col = f"{col}_translated({code})"
            mask = df[col].apply(lambda x: isinstance(x, str))
            uniques = df.loc[mask, col].astype(str).unique()
            translations = {}
            for u in uniques:
                translations[u] = translate_text_cell(u, code)
            df[new_col] = df[col].apply(lambda x: translations.get(x, x) if isinstance(x, str) else x)
        current_df = df
        tampilkan_dataframe(df)
        label_status.config(text=f"✔ Translate selesai ({code})", fg="green")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal translate:\n{e}")
        label_status.config(text="✘ Translate gagal", fg="red")
    finally:
        btn_translate.config(state=tk.NORMAL)

# ----------------------------
# Save / Save as
# ----------------------------
def save_file():
    if current_df is None:
        messagebox.showwarning("Peringatan", "Tidak ada data untuk disimpan.")
        return
    name = entry_filename.get().strip()
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    try:
        current_df.to_excel(name, index=False)
        label_status.config(text=f"✔ Disimpan: {name}", fg="green")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal menyimpan:\n{e}")

def save_as_file():
    if current_df is None:
        messagebox.showwarning("Peringatan", "Tidak ada data untuk disimpan.")
        return
    path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files","*.xlsx")])
    if not path:
        return
    try:
        current_df.to_excel(path, index=False)
        label_status.config(text=f"✔ Save As sukses: {path}", fg="green")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal menyimpan:\n{e}")

# ----------------------------
# Column selectors update
# ----------------------------
def update_column_selectors():
    if current_df is None:
        cols = []
    else:
        cols = list(current_df.columns)

    combo_x['values'] = ["--Auto--"] + cols
    combo_y['values'] = ["--Auto--"] + cols
    combo_cat['values'] = ["--Auto--"] + cols
    
    if combo_x.get() not in combo_x['values']:
        combo_x.set("--Auto--")
    if combo_y.get() not in combo_y['values']:
        combo_y.set("--Auto--")
    if combo_cat.get() not in combo_cat['values']:
        combo_cat.set("--Auto--")

# ----------------------------
# Clear visuals
# ----------------------------
def clear_visuals():
    global chart_widgets
    for widget in chart_widgets:
        try:
            widget.destroy()
        except:
            pass
    chart_widgets = []

# ----------------------------
# Chart creation helpers
# ----------------------------
def save_current_figure():
    global last_fig
    if last_fig is None:
        messagebox.showwarning("Peringatan", "Belum ada grafik untuk disimpan.")
        return
    path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG Image","*.png")])
    if not path:
        return
    try:
        last_fig.savefig(path, dpi=150, bbox_inches='tight', facecolor=last_fig.get_facecolor())
        messagebox.showinfo("Sukses", f"Grafik tersimpan: {path}")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal menyimpan grafik:\n{e}")

def make_scatter(ax, df, xcol, ycol):
    ax.scatter(df[xcol], df[ycol], alpha=0.6, s=50)
    ax.set_xlabel(xcol, fontsize=10)
    ax.set_ylabel(ycol, fontsize=10)
    ax.set_title(f"Scatter: {xcol} vs {ycol}", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)

def make_line(ax, df, xcol, ycol):
    try:
        ax.plot(df[xcol], df[ycol], linewidth=2)
    except Exception:
        ax.plot(df[ycol], linewidth=2)
    ax.set_xlabel(xcol, fontsize=10)
    ax.set_ylabel(ycol, fontsize=10)
    ax.set_title(f"Line: {ycol} over {xcol}", fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3)

def make_hist(ax, df, col):
    ax.hist(df[col].dropna(), bins=15, edgecolor='black', alpha=0.7)
    ax.set_title(f"Histogram: {col}", fontsize=11, fontweight='bold')
    ax.set_xlabel(col, fontsize=10)
    ax.set_ylabel("Frequency", fontsize=10)
    ax.grid(True, alpha=0.3, axis='y')

def make_pie(ax, df, col):
    counts = df[col].value_counts()
    colors = plt.cm.Set3(range(len(counts)))
    ax.pie(counts, labels=counts.index.astype(str), autopct="%1.1f%%", 
           startangle=90, colors=colors)
    ax.set_title(f"Pie Chart: {col}", fontsize=11, fontweight='bold')

def make_bar_category_numeric(ax, df, cat_col, num_col):
    try:
        avg = df.groupby(cat_col)[num_col].mean().sort_values()
        bars = ax.bar(range(len(avg)), avg.values, edgecolor='black', alpha=0.7)
        ax.set_xticks(range(len(avg)))
        ax.set_xticklabels(avg.index.astype(str), rotation=45, ha='right')
        ax.set_title(f"Avg {num_col} per {cat_col}", fontsize=11, fontweight='bold')
        ax.set_ylabel(f"Average {num_col}", fontsize=10)
        ax.grid(True, alpha=0.3, axis='y')
    except Exception as e:
        ax.text(0.5, 0.5, f"Error: {str(e)}", ha='center', va='center')

def make_heatmap(ax, df):
    corr = df.select_dtypes(include=[np.number]).corr()
    if corr.empty:
        ax.text(0.5, 0.5, "Tidak ada kolom numerik untuk korelasi", ha='center', va='center')
        return
    im = ax.imshow(corr, aspect='auto', cmap='coolwarm', vmin=-1, vmax=1)
    ax.set_xticks(range(len(corr.columns)))
    ax.set_yticks(range(len(corr.columns)))
    ax.set_xticklabels(corr.columns, rotation=45, ha='right', fontsize=9)
    ax.set_yticklabels(corr.columns, fontsize=9)
    ax.set_title("Correlation Heatmap", fontsize=11, fontweight='bold')
    plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)

def make_box(ax, df, col):
    data = df[col].dropna()
    bp = ax.boxplot([data], patch_artist=True)
    for patch in bp['boxes']:
        patch.set_facecolor('lightblue')
    ax.set_xticklabels([col])
    ax.set_title(f"Boxplot: {col}", fontsize=11, fontweight='bold')
    ax.set_ylabel("Values", fontsize=10)
    ax.grid(True, alpha=0.3, axis='y')

def make_area(ax, df, xcol, ycol):
    try:
        ax.fill_between(range(len(df)), df[ycol], alpha=0.5)
        ax.plot(range(len(df)), df[ycol], linewidth=2)
        ax.set_title(f"Area Chart: {ycol}", fontsize=11, fontweight='bold')
        ax.set_xlabel("Index", fontsize=10)
        ax.set_ylabel(ycol, fontsize=10)
        ax.grid(True, alpha=0.3)
    except Exception:
        ax.text(0.5, 0.5, "Tidak bisa membuat area chart", ha='center', va='center')

# ----------------------------
# Main function to render charts in grid layout
# ----------------------------
def render_charts():
    global last_fig, is_dark_mode, chart_widgets
    clear_visuals()
    
    if current_df is None:
        messagebox.showwarning("Peringatan", "Muat data dulu.")
        return
    
    df = current_df.copy()
    sel = combo_chart.get()
    
    # Determine which charts to draw
    if sel == "All Charts":
        charts_to_draw = CHART_OPTIONS[:-1]
    else:
        charts_to_draw = [sel]

    # Get column selections
    xsel = combo_x.get()
    ysel = combo_y.get()
    catsel = combo_cat.get()

    cols = list(df.columns)
    numeric_cols = [c for c in cols if pd.api.types.is_numeric_dtype(df[c])]
    object_cols = [c for c in cols if pd.api.types.is_object_dtype(df[c])]

    def pick_x():
        if xsel != "--Auto--" and xsel in cols:
            return xsel
        return numeric_cols[0] if numeric_cols else (cols[0] if cols else None)

    def pick_y():
        if ysel != "--Auto--" and ysel in cols:
            return ysel
        return numeric_cols[1] if len(numeric_cols) > 1 else (numeric_cols[0] if numeric_cols else None)

    def pick_cat():
        if catsel != "--Auto--" and catsel in cols:
            return catsel
        return object_cols[0] if object_cols else None

    xcol = pick_x()
    ycol = pick_y()
    catcol = pick_cat()

    # Set theme
    bg = "#2e2e2e" if is_dark_mode else "white"
    fg = "white" if is_dark_mode else "black"
    
    # Calculate grid layout
    num_charts = len(charts_to_draw)
    cols_per_row = 2
    
    # Create charts
    for idx, chart in enumerate(charts_to_draw):
        # Create a frame for each chart
        chart_frame = tk.Frame(frame_visual, relief=tk.RIDGE, borderwidth=2, bg=bg)
        chart_frame.grid(row=idx//cols_per_row, column=idx%cols_per_row, 
                        padx=10, pady=10, sticky="nsew")
        chart_widgets.append(chart_frame)
        
        # Create figure with proper sizing
        fig = plt.Figure(figsize=(6, 4.5), facecolor=bg, dpi=100)
        ax = fig.add_subplot(111, facecolor=bg)
        
        # Style the axis
        ax.tick_params(colors=fg, labelsize=9)
        for spine in ax.spines.values():
            spine.set_edgecolor(fg)
        
        # Draw appropriate chart
        try:
            if chart == "Scatter Plot":
                if xcol and ycol and xcol in cols and ycol in cols:
                    if pd.api.types.is_numeric_dtype(df[xcol]) and pd.api.types.is_numeric_dtype(df[ycol]):
                        make_scatter(ax, df, xcol, ycol)
                    else:
                        ax.text(0.5, 0.5, "Scatter membutuhkan 2 kolom numerik", 
                               ha="center", va="center", color=fg)
                else:
                    ax.text(0.5, 0.5, "Kolom X/Y tidak tersedia", 
                           ha="center", va="center", color=fg)

            elif chart == "Line Chart":
                if xcol and ycol and xcol in cols and ycol in cols:
                    make_line(ax, df, xcol, ycol)
                else:
                    ax.text(0.5, 0.5, "Kolom X/Y tidak tersedia", 
                           ha="center", va="center", color=fg)

            elif chart == "Bar Chart (Category vs Numeric)":
                if catcol and xcol and catcol in cols and xcol in cols:
                    make_bar_category_numeric(ax, df, catcol, xcol)
                else:
                    ax.text(0.5, 0.5, "Butuh 1 kolom kategori dan 1 numerik", 
                           ha="center", va="center", color=fg)

            elif chart == "Histogram":
                if xcol and xcol in cols and pd.api.types.is_numeric_dtype(df[xcol]):
                    make_hist(ax, df, xcol)
                elif numeric_cols:
                    make_hist(ax, df, numeric_cols[0])
                else:
                    ax.text(0.5, 0.5, "Tidak ada kolom numerik", 
                           ha="center", va="center", color=fg)

            elif chart == "Pie Chart":
                if catcol and catcol in cols:
                    make_pie(ax, df, catcol)
                elif object_cols:
                    make_pie(ax, df, object_cols[0])
                else:
                    ax.text(0.5, 0.5, "Tidak ada kolom kategorikal", 
                           ha="center", va="center", color=fg)

            elif chart == "Heatmap (Correlation)":
                make_heatmap(ax, df)

            elif chart == "Box Plot":
                target = xcol if xcol and xcol in cols else (numeric_cols[0] if numeric_cols else None)
                if target and pd.api.types.is_numeric_dtype(df[target]):
                    make_box(ax, df, target)
                else:
                    ax.text(0.5, 0.5, "Tidak ada kolom numerik", 
                           ha="center", va="center", color=fg)

            elif chart == "Area Chart":
                if ycol and ycol in cols:
                    make_area(ax, df, xcol, ycol)
                else:
                    ax.text(0.5, 0.5, "Kolom Y tidak tersedia", 
                           ha="center", va="center", color=fg)
        except Exception as e:
            ax.text(0.5, 0.5, f"Error: {str(e)}", ha='center', va='center', color=fg)
        
        # Apply color theme to text elements
        if ax.get_title():
            ax.title.set_color(fg)
        if ax.get_xlabel():
            ax.xaxis.label.set_color(fg)
        if ax.get_ylabel():
            ax.yaxis.label.set_color(fg)
        
        # Tight layout
        fig.tight_layout()
        
        # Embed figure
        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(fill="both", expand=True, padx=5, pady=5)
        chart_widgets.append(canvas_widget)
        
        last_fig = fig
    
    # Configure grid weights for responsive layout
    for i in range((num_charts + cols_per_row - 1) // cols_per_row):
        frame_visual.grid_rowconfigure(i, weight=1)
    for i in range(cols_per_row):
        frame_visual.grid_columnconfigure(i, weight=1)
    
    # Update scroll region
    frame_visual.update_idletasks()
    canvas_visual.configure(scrollregion=canvas_visual.bbox("all"))
    
    label_status.config(text=f"✔ {len(charts_to_draw)} grafik ditampilkan", fg="green")

# ----------------------------
# Dark mode toggle
# ----------------------------
def toggle_dark_mode():
    global is_dark_mode
    is_dark_mode = not is_dark_mode
    bg = "#2e2e2e" if is_dark_mode else "SystemButtonFace"
    fg = "white" if is_dark_mode else "black"
    
    root.configure(bg=bg)
    for widget in [frame_top, frame_lang, frame_save, frame_table, frame_chart, frame_actions]:
        widget.configure(bg=bg)
    
    # Update canvas background
    canvas_visual.configure(bg=bg)
    frame_visual.configure(bg=bg)
    
    # Re-render charts if any exist
    if current_df is not None and chart_widgets:
        render_charts()

# ----------------------------
# Build GUI
# ----------------------------
root = tk.Tk()
root.title("Excel Analyzer - Enhanced Visual Layout")
root.geometry("1400x900")

# Top frame: load, quick load
frame_top = tk.Frame(root, bg="SystemButtonFace")
frame_top.pack(fill="x", pady=6)

btn_load = tk.Button(frame_top, text="📁 Load Excel", command=load_excel, font=("Arial", 10, "bold"))
btn_load.grid(row=0, column=0, padx=6, pady=5)

btn_quick = tk.Button(frame_top, text="⚡ Quick Load Sample", command=quick_load, font=("Arial", 10))
btn_quick.grid(row=0, column=1, padx=6, pady=5)

# Translate controls
frame_lang = tk.Frame(root, bg="SystemButtonFace")
frame_lang.pack(fill="x", pady=6)

tk.Label(frame_lang, text="Pilih Bahasa Translate:", bg="SystemButtonFace").grid(row=0, column=0, padx=5)
combo_lang = ttk.Combobox(frame_lang, values=LANG_DISPLAY, width=40)
combo_lang.grid(row=0, column=1, padx=5)
combo_lang.set("English (en)")

btn_translate = tk.Button(frame_lang, text="🌐 Translate Seluruh Isi", command=translate_entire_df_thread)
btn_translate.grid(row=0, column=2, padx=6)

# Save controls
frame_save = tk.Frame(root, bg="SystemButtonFace")
frame_save.pack(fill="x", pady=6)

tk.Label(frame_save, text="Nama Output:", bg="SystemButtonFace").grid(row=0, column=0, padx=5)
entry_filename = tk.Entry(frame_save, width=40)
entry_filename.grid(row=0, column=1, padx=5)
entry_filename.insert(0, "output.xlsx")

btn_save = tk.Button(frame_save, text="💾 Save", command=save_file)
btn_save.grid(row=0, column=2, padx=4)
btn_save_as = tk.Button(frame_save, text="💾 Save As...", command=save_as_file)
btn_save_as.grid(row=0, column=3, padx=4)

# Chart selectors
frame_chart = tk.Frame(root, bg="SystemButtonFace")
frame_chart.pack(fill="x", pady=6)

tk.Label(frame_chart, text="Jenis Grafik:", bg="SystemButtonFace").grid(row=0, column=0, padx=5)
combo_chart = ttk.Combobox(frame_chart, values=CHART_OPTIONS, width=28)
combo_chart.grid(row=0, column=1, padx=5)
combo_chart.set("All Charts")

tk.Label(frame_chart, text="Kolom X:", bg="SystemButtonFace").grid(row=0, column=2, padx=5)
combo_x = ttk.Combobox(frame_chart, values=["--Auto--"], width=20)
combo_x.grid(row=0, column=3, padx=5)
combo_x.set("--Auto--")

tk.Label(frame_chart, text="Kolom Y:", bg="SystemButtonFace").grid(row=0, column=4, padx=5)
combo_y = ttk.Combobox(frame_chart, values=["--Auto--"], width=20)
combo_y.grid(row=0, column=5, padx=5)
combo_y.set("--Auto--")

tk.Label(frame_chart, text="Kolom Kategori:", bg="SystemButtonFace").grid(row=0, column=6, padx=5)
combo_cat = ttk.Combobox(frame_chart, values=["--Auto--"], width=20)
combo_cat.grid(row=0, column=7, padx=5)
combo_cat.set("--Auto--")

# Actions frame
frame_actions = tk.Frame(root, bg="SystemButtonFace")
frame_actions.pack(fill="x", pady=6)

btn_render = tk.Button(frame_actions, text="📊 Tampilkan Grafik", command=render_charts, 
                       font=("Arial", 10, "bold"), bg="#4CAF50", fg="white")
btn_render.grid(row=0, column=0, padx=6)

btn_save_chart = tk.Button(frame_actions, text="📥 Download Grafik (PNG)", command=save_current_figure,
                           font=("Arial", 10))
btn_save_chart.grid(row=0, column=1, padx=6)

btn_toggle_dark = tk.Button(frame_actions, text="🌙 Toggle Dark Mode", command=toggle_dark_mode,
                            font=("Arial", 10))
btn_toggle_dark.grid(row=0, column=2, padx=6)

# Status label
label_status = tk.Label(root, text="Siap untuk memuat data", fg="blue", font=("Arial", 10))
label_status.pack(pady=5)

# Visual container with scrollbar
visual_container = tk.Frame(root, height=500, relief=tk.SUNKEN, borderwidth=2)
visual_container.pack(fill="both", expand=True, padx=10, pady=5)
visual_container.pack_propagate(False)

canvas_visual = tk.Canvas(visual_container, bg="white")
scroll_visual = tk.Scrollbar(visual_container, orient="vertical", command=canvas_visual.yview)
scroll_visual.pack(side="right", fill="y")
canvas_visual.pack(side="left", fill="both", expand=True)
canvas_visual.configure(yscrollcommand=scroll_visual.set)

frame_visual = tk.Frame(canvas_visual, bg="white")
canvas_visual.create_window((0, 0), window=frame_visual, anchor="nw")

def on_frame_config(event):
    canvas_visual.configure(scrollregion=canvas_visual.bbox("all"))

frame_visual.bind("<Configure>", on_frame_config)

# Bind mousewheel for scrolling
def on_mousewheel(event):
    canvas_visual.yview_scroll(int(-1*(event.delta/120)), "units")

canvas_visual.bind_all("<MouseWheel>", on_mousewheel)

# Frame table for DataFrame
frame_table = tk.Frame(root, relief=tk.GROOVE, borderwidth=2, height=200)
frame_table.pack(fill="both", expand=False, padx=10, pady=10)
frame_table.pack_propagate(False)

# Initialize column selectors
update_column_selectors()

root.mainloop()