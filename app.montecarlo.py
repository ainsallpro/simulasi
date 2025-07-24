import pandas as pd
import os
import random
import math
import matplotlib.pyplot as plt
import numpy as np
import streamlit as st

# --- Konfigurasi Aplikasi ---
FILE_PATHS = {
    'A': "Prob A.xlsx",
    'B': "Prob B.xlsx",
    'AB': "Prob AB.xlsx",
    'O': "Prob O.xlsx"
}

# Pengaturan Warna Untuk Grafik
PLOT_COLORS = {
    'A': '#1f77b4', # Biru gelap
    'B': '#ff7f0e', # Oranye
    'AB': '#d62728', # Merah
    'O': '#2ca02c', # Hijau
}

# --- Fungsi Pembantu (Helper Functions) ---

def clean_interval_string(interval_str):
    """
    Membersihkan dan menstandarkan format string interval (misal: '103â-130' menjadi '103-130').
    """
    return str(interval_str).replace('â', '-').replace('–', '-').replace('—', '-').replace(' ', '').replace(',', '')

def parse_interval(interval_str):
    """
    Mengurai string interval 'a-b' dan mengembalikan tuple (a, b) dalam bentuk integer.
    Mengembalikan (None, None) jika terjadi error parsing.
    """
    try:
        a, b = map(int, interval_str.split('-'))
        return a, b
    except ValueError:
        return None, None

# --- Fungsi Pemuatan dan Pra-pemrosesan Data ---

@st.cache_data # Menggunakan cache Streamlit untuk mempercepat pemuatan data Excel
def load_distribusi_from_excel(file_path):
    """
    Memuat data distribusi golongan darah dari file Excel.
    """
    try:
        df = pd.read_excel(file_path)
        df = df[df['Interval Kelas '].notna() & df['Probabilitas'].notna()].copy()
        df['Probabilitas'] = df['Probabilitas'].astype(str).str.replace(',', '.').astype(float)
        df['Prob Kumulatif '] = df['Prob Kumulatif '].astype(str).str.replace(',', '.').astype(float)
        df['Prob Kumulatif * 100'] = df['Prob Kumulatif * 100'].astype(str).str.replace(',', '.').astype(float)
        return df[['No', 'Interval Kelas ', 'Frekuensi', 'Probabilitas', 'Prob Kumulatif ', 'Prob Kumulatif * 100']].reset_index(drop=True)
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memuat {file_path}: {e}")
        return pd.DataFrame()

@st.cache_data # Menggunakan cache Streamlit untuk mempercepat pemuatan semua distribusi
def load_all_distributions(_file_paths):
    """
    Memuat data distribusi untuk semua golongan darah yang ditentukan.
    """
    distribusi_dict = {}
    for gol, path in _file_paths.items():
        if os.path.exists(path):
            distribusi_dict[gol] = load_distribusi_from_excel(path)
        else:
            st.warning(f"File untuk golongan darah {gol} tidak ditemukan di: {path}")
    return distribusi_dict

# --- Fungsi Tampilan Data ---

def display_distribution_table(df, golongan):
    """
    Menampilkan tabel distribusi yang diformat untuk golongan darah tertentu di Streamlit.
    """
    st.subheader(f"Tabel Distribusi Golongan Darah {golongan} 📋")
    
    display_df = df.copy()
    
    # Hitung dan tambahkan kolom 'Titik Tengah'
    midpoints = []
    for _, row in display_df.iterrows():
        interval = clean_interval_string(row['Interval Kelas '])
        a, b = parse_interval(interval)
        if a is not None and b is not None:
            midpoints.append(math.ceil((a + b) / 2))
        else:
            midpoints.append(None)
    display_df['Titik Tengah'] = midpoints

    # Siapkan kolom 'Interval Angka Acak'
    display_df['Interval Angka Acak'] = ""
    lower_bound = 0
    for i, row in display_df.iterrows():
        upper_bound = int(round(row['Prob Kumulatif * 100']))
        display_df.loc[i, 'Interval Angka Acak'] = f"{str(lower_bound).zfill(2)} - {str(upper_bound).zfill(2)}"
        lower_bound = upper_bound + 1
    
    # Bersihkan string interval dan ganti nama kolom
    display_df['Interval Kelas '] = display_df['Interval Kelas '].apply(clean_interval_string)
    display_df = display_df.rename(columns={
        'Interval Kelas ': 'Interval Kelas',
        'Prob Kumulatif ': 'Prob Kumulatif'
    })
    
    # Tentukan urutan kolom yang akan ditampilkan
    cols_to_display = [
        'No', 'Interval Kelas', 'Frekuensi', 'Probabilitas', 'Prob Kumulatif',
        'Prob Kumulatif * 100', 'Titik Tengah', 'Interval Angka Acak'
    ]
    
    st.dataframe(display_df[cols_to_display], hide_index=True)

# --- Logika Simulasi ---

def get_simulation_value(df, random_number):
    """
    Menentukan nilai simulasi berdasarkan angka acak dan distribusi yang diberikan.
    """
    lower_bound = 0
    for _, row in df.iterrows():
        upper_bound = int(round(row['Prob Kumulatif * 100']))
        if lower_bound <= random_number <= upper_bound:
            interval = clean_interval_string(row['Interval Kelas '])
            a, b = parse_interval(interval)
            if a is not None and b is not None:
                return math.ceil((a + b) / 2)
            else:
                return 0
        lower_bound = upper_bound + 1
    return 0

@st.cache_data # Menggunakan cache untuk hasil simulasi agar lebih cepat jika inputnya sama
def run_monte_carlo_simulation(_distribusi_dict, num_simulations):
    """
    Menjalankan simulasi Monte Carlo untuk pemakaian darah.
    """
    simulation_results = []
    progress_text = "Simulasi sedang berjalan. Mohon tunggu. ⏳"
    my_bar = st.progress(0, text=progress_text)

    for i in range(num_simulations):
        a_acak = random.randint(0, 99)
        b_acak = random.randint(0, 99)
        ab_acak = random.randint(0, 99)
        o_acak = random.randint(0, 99)

        sim_a = get_simulation_value(_distribusi_dict.get('A', pd.DataFrame()), a_acak)
        sim_b = get_simulation_value(_distribusi_dict.get('B', pd.DataFrame()), b_acak)
        sim_ab = get_simulation_value(_distribusi_dict.get('AB', pd.DataFrame()), ab_acak)
        sim_o = get_simulation_value(_distribusi_dict.get('O', pd.DataFrame()), o_acak)

        total_sim = sim_a + sim_b + sim_ab + sim_o
        pa = (sim_a / total_sim * 100) if total_sim else 0
        pb = (sim_b / total_sim * 100) if total_sim else 0
        pab = (sim_ab / total_sim * 100) if total_sim else 0
        po = (sim_o / total_sim * 100) if total_sim else 0

        simulation_results.append({
            'Periode': i + 1, 'Angka Acak A': a_acak, 'Angka Acak B': b_acak,
            'Angka Acak AB': ab_acak, 'Angka Acak O': o_acak, 'Simulasi A': sim_a,
            'Simulasi B': sim_b, 'Simulasi AB': sim_ab, 'Simulasi O': sim_o,
            'Total Simulasi': total_sim, 'Pmk A%': pa, 'Pmk B%': pb,
            'Pmk AB%': pab, 'Pmk O%': po
        })
        my_bar.progress((i + 1) / num_simulations, text=progress_text)
        
    my_bar.empty()
    return pd.DataFrame(simulation_results)

# --- Analisis dan Visualisasi ---

def perform_summary_analysis(sim_df):
    """
    Melakukan dan menampilkan statistik ringkasan dari simulasi.
    """
    summary_stats = sim_df[['Simulasi A', 'Simulasi B', 'Simulasi AB', 'Simulasi O', 'Total Simulasi']].agg(['mean', 'std', 'min', 'max']).transpose()
    summary_stats.columns = ['Rata-rata', 'Standar Deviasi', 'Minimum', 'Maksimum']
    st.markdown("### Statistik Ringkasan Pemakaian Darah (kantong) 📊:")
    st.dataframe(summary_stats.round(2))

    max_total_usage_period = sim_df.loc[sim_df['Total Simulasi'].idxmax()]
    min_total_usage_period = sim_df.loc[sim_df['Total Simulasi'].idxmin()]

    st.markdown(f"### Periode dengan Total Pemakaian Tertinggi 📈:")
    st.markdown(f"Periode **{int(max_total_usage_period['Periode'])}** (Total: **{int(max_total_usage_period['Total Simulasi'])}** kantong)")
    st.markdown(f"Komposisi: A={int(max_total_usage_period['Simulasi A'])}, B={int(max_total_usage_period['Simulasi B'])}, AB={int(max_total_usage_period['Simulasi AB'])}, O={int(max_total_usage_period['Simulasi O'])}")

    st.markdown(f"### Periode dengan Total Pemakaian Terendah 📉:")
    st.markdown(f"Periode **{int(min_total_usage_period['Periode'])}** (Total: **{int(min_total_usage_period['Total Simulasi'])}** kantong)")
    st.markdown(f"Komposisi: A={int(min_total_usage_period['Simulasi A'])}, B={int(min_total_usage_period['Simulasi B'])}, AB={int(min_total_usage_period['Simulasi AB'])}, O={int(min_total_usage_period['Simulasi O'])}")

def provide_decision_insights(sim_df):
    """
    Memberikan wawasan dan rekomendasi untuk pengambil keputusan.
    """
    st.header("💡 WAWASAN & REKOMENDASI UNTUK PENGAMBIL KEPUTUSAN 🎯")
    avg_usages = {
        'A': sim_df['Simulasi A'].mean(), 'B': sim_df['Simulasi B'].mean(),
        'AB': sim_df['Simulasi AB'].mean(), 'O': sim_df['Simulasi O'].mean()
    }
    sorted_avg_usages = sorted(avg_usages.items(), key=lambda item: item[1], reverse=True)
    total_avg_usage = sim_df['Total Simulasi'].mean()

    st.markdown(f"Dari hasil simulasi **{len(sim_df)} periode**:")
    st.markdown(f"**1. Prediksi Kebutuhan Darah Secara Umum: 🩸**")
    st.markdown(f"Rata-rata total pemakaian darah per periode adalah sekitar **{total_avg_usage:.0f} kantong**.")
    st.markdown(f"**2. Golongan Darah Paling Banyak dan Paling Sedikit Digunakan: 📊**")
    for blood_type, avg_usage in sorted_avg_usages:
        st.markdown(f"- **Golongan {blood_type}:** Rata-rata pemakaian sekitar **{avg_usage:.0f} kantong per periode**.")
    st.markdown(f"Permintaan paling tinggi adalah **Golongan {sorted_avg_usages[0][0]}** dan paling rendah **Golongan {sorted_avg_usages[-1][0]}**.")
    st.markdown(f"**3. Strategi Pengelolaan Stok: 📦**")
    st.markdown(f"- **Prioritas Tinggi ({sorted_avg_usages[0][0]} & {sorted_avg_usages[1][0]}):** Pastikan stok selalu mencukupi.")
    st.markdown(f"- **Prioritas Rendah ({sorted_avg_usages[-1][0]}):** Kelola stok agar tidak menumpuk.")

def plot_blood_usage_bar_chart(sim_df, colors):
    """
    Menggambar grafik batang pemakaian darah per periode.
    """
    fig, ax = plt.subplots(figsize=(20, 8)) # Ukuran disesuaikan
    bar_width = 0.2
    r1 = np.arange(len(sim_df))
    r2 = [x + bar_width for x in r1]
    r3 = [x + bar_width * 2 for x in r1]
    r4 = [x + bar_width * 3 for x in r1]

    ax.bar(r1, sim_df['Simulasi A'], color=colors['A'], width=bar_width, edgecolor='black', label='Simulasi A')
    ax.bar(r2, sim_df['Simulasi B'], color=colors['B'], width=bar_width, edgecolor='black', label='Simulasi B')
    ax.bar(r3, sim_df['Simulasi AB'], color=colors['AB'], width=bar_width, edgecolor='black', label='Simulasi AB')
    ax.bar(r4, sim_df['Simulasi O'], color=colors['O'], width=bar_width, edgecolor='black', label='Simulasi O')

    ax.set_xlabel('Periode Simulasi', fontweight='bold', fontsize=12)
    ax.set_ylabel('Jumlah Pemakaian (kantong)', fontweight='bold', fontsize=12)
    ax.set_title('GRAFIK SIMULASI PEMAKAIAN DARAH PER PERIODE', fontweight='bold', fontsize=16)
    
    if len(sim_df) > 0:
        ax.set_xticks([r + bar_width * 1.5 for r in range(len(sim_df))], labels=sim_df['Periode'], rotation=90, fontsize=8)
    
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    ax.legend()
    fig.tight_layout()
    st.pyplot(fig)

def plot_average_usage_pie_chart(sim_df, colors):
    """
    Menggambar grafik pai rata-rata pemakaian darah.
    """
    fig, ax = plt.subplots(figsize=(6, 5))
    avg_pemakaian = [
        sim_df['Simulasi A'].mean(), sim_df['Simulasi B'].mean(),
        sim_df['Simulasi AB'].mean(), sim_df['Simulasi O'].mean()
    ]
    labels = ['PEMAKAIAN A', 'PEMAKAIAN B', 'PEMAKAIAN AB', 'PEMAKAIAN O']
    pie_colors = [colors['A'], colors['B'], colors['AB'], colors['O']] 
    
    ax.pie(avg_pemakaian, labels=labels, autopct='%1.1f%%', startangle=140, colors=pie_colors,
           wedgeprops={'edgecolor': 'black'}, textprops={'fontsize': 10})
    ax.set_title('RATA-RATA PEMAKAIAN SIMULASI DARAH PER GOLONGAN', fontweight='bold', fontsize=12)
    ax.axis('equal')
    st.pyplot(fig)

# --- Eksekusi Program Utama (Main Execution) ---

if __name__ == "__main__":
    st.set_page_config(layout="wide", page_title="Simulasi Monte Carlo")
    
    # Custom CSS
    st.markdown("""
        <style>
            .main { background-color: #f0f2f6; }
            h1, h2, h3 { text-align: center; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1>🩸 Simulasi Pemakaian Darah di UTD RSUD dr. Soekardjo Kota Tasikmalaya pada 2021-2023 dengan Simulasi Monte Carlo 🩸</h1>", unsafe_allow_html=True)
    st.markdown("---")

    # Tombol untuk membersihkan cache
    if st.button("🔄 Hapus Cache & Muat Ulang Data dari Excel"):
        st.cache_data.clear()
        st.success("Cache berhasil dihapus! Data terbaru dari Excel akan dimuat pada proses selanjutnya.")
        st.rerun()

    # 1. Pemuatan Data Distribusi
    st.header("1. 📊 Data Distribusi Golongan Darah")
    distribusi_data = load_all_distributions(FILE_PATHS)

    tab_titles = list(FILE_PATHS.keys())
    tabs = st.tabs(tab_titles)
    for i, gol in enumerate(FILE_PATHS.keys()):
        with tabs[i]:
            if gol in distribusi_data and not distribusi_data[gol].empty:
                display_distribution_table(distribusi_data[gol], gol)
            else:
                st.warning(f"Data distribusi untuk Golongan Darah {gol} tidak tersedia. ⚠️")

    # 2. Menjalankan Simulasi Monte Carlo
    st.markdown("---")
    st.header("2. 🧪 Hasil Simulasi Monte Carlo")
    
    num_simulations_str = st.text_input(
        "Pilih Jumlah Periode Simulasi:", value="",
        help="Masukkan jumlah periode simulasi (misal: 84, 120, 365)."
    )
    
    if st.button('Jalankan Simulasi ▶️'):
        if num_simulations_str:
            try:
                num_simulations_input = int(num_simulations_str)
                if num_simulations_input > 0:
                    with st.spinner(f'Menjalankan simulasi untuk {num_simulations_input} periode...'):
                        simulation_df = run_monte_carlo_simulation(distribusi_data, num_simulations_input)
                    
                    st.subheader("Tabel Hasil Simulasi Per Periode:")
                    st.dataframe(simulation_df.round(2), hide_index=True)
                    
                    st.markdown("---")
                    st.header("3. 📈 Analisis dan Visualisasi")
                    
                    col1, col2 = st.columns([3, 2]) # Buat 2 kolom, grafik batang lebih lebar
                    with col1:
                        st.subheader("Grafik Pemakaian per Periode")
                        plot_blood_usage_bar_chart(simulation_df, PLOT_COLORS)
                    with col2:
                        st.subheader("Grafik Rata-Rata Pemakaian")
                        plot_average_usage_pie_chart(simulation_df, PLOT_COLORS)
                    
                    st.markdown("---")
                    perform_summary_analysis(simulation_df)
                    st.markdown("---")
                    provide_decision_insights(simulation_df)
                else:
                    st.error("Jumlah periode simulasi harus lebih besar dari 0. ❌")
            except ValueError:
                st.error("Input tidak valid. Harap masukkan angka bulat. 🔢")
        else:
            st.warning("Harap masukkan jumlah periode simulasi terlebih dahulu. ⚠️")
