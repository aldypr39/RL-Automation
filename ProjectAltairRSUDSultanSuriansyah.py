# %%
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time
import threading

# --- Variabel Global ---
file_path = ""
chrome_driver_path = ""

# Fungsi untuk memilih file Excel
def select_file():
    global file_path
    file_path = filedialog.askopenfilename(title="Pilih File Excel", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
    if file_path:
        file_name = file_path.split("/")[-1]
        file_label.configure(text=f"File Terpilih: {file_name}")
    else:
        file_label.configure(text="Belum ada file yang dipilih")

# Fungsi untuk memilih path ChromeDriver
def select_driver_path():
    global chrome_driver_path
    chrome_driver_path = filedialog.askopenfilename(title="Pilih File ChromeDriver", filetypes=[("Executable files", "*.exe"), ("All files", "*.*")])
    if chrome_driver_path:
        driver_name = chrome_driver_path.split("/")[-1]
        driver_label.configure(text=f"Driver Manual: {driver_name}")
    else:
        driver_label.configure(text="Belum ada driver manual yang dipilih")

# --- Fungsi Baru: Update Log di UI ---
def update_log(message):
    log_textbox.configure(state="normal")
    log_textbox.insert("end", message + "\n")
    log_textbox.configure(state="disabled")
    log_textbox.see("end") # Auto-scroll to the bottom

# --- Fungsi Baru: Update Progress Bar ---
def update_progress(value):
    progress_bar.set(value)

# --- Fungsi Baru: Hapus Log ---
def clear_log():
    log_textbox.configure(state="normal")
    log_textbox.delete("1.0", "end")
    log_textbox.configure(state="disabled")

# --- Fungsi untuk menampilkan/menyembunyikan password ---
def toggle_password():
    if show_password_var.get():
        password_entry.configure(show="")
    else:
        password_entry.configure(show="*")

# Fungsi utama yang menjalankan proses Selenium
def run_selenium_process():
    global chrome_driver_path
    driver = None # Inisialisasi driver
    try:
        # Mengambil email dan password dari entry
        email = email_entry.get()
        password = password_entry.get()
        selected_rl = rl_choice.get() # Ambil pilihan RL di awal
        
        update_log("Membaca file Excel...")
        # --- PENYESUAIAN HEADER BERDASARKAN RL ---
        if selected_rl == "RL 5.1":
            df = pd.read_excel(file_path, header=4)
            update_log("File untuk RL 5.1 dibaca dengan header baris ke-5.")
        else: # Asumsi default atau RL 4.1
            df = pd.read_excel(file_path, header=3)
            update_log("File untuk RL 4.1 dibaca dengan header baris ke-4.")
        # --- AKHIR PENYESUAIAN ---

        total_rows = len(df)
        update_log(f"Ditemukan {total_rows} baris data untuk diproses.")

        # --- LOGIKA BARU: Coba otomatis, jika gagal pakai manual ---
        try:
            update_log("Mencoba setup ChromeDriver otomatis...")
            service = ChromeService(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service)
            update_log("ChromeDriver otomatis berhasil disiapkan.")
        except Exception as e_auto:
            update_log(f"(!) Gagal setup otomatis: {e_auto}")
            if chrome_driver_path:
                update_log("Mencoba menggunakan driver manual yang dipilih...")
                service = ChromeService(executable_path=chrome_driver_path)
                driver = webdriver.Chrome(service=service)
                update_log("Driver manual berhasil digunakan.")
            else:
                update_log("ERROR: Setup otomatis gagal dan tidak ada driver manual yang dipilih.")
                update_log("Silakan pilih driver manual di tab Pengaturan dan coba lagi.")
                # Mengaktifkan kembali tombol setelah selesai
                start_button.configure(state="normal")
                return # Hentikan eksekusi jika tidak ada driver
        # --- AKHIR LOGIKA BARU ---
        
        driver.get("https://akun-yankes.kemkes.go.id/beranda")
        driver.maximize_window()
        driver.implicitly_wait(15)

        update_log("Melakukan login...")
        driver.find_element(By.ID, "floatingInput").send_keys(email)
        driver.find_element(By.ID, "floatingPassword").send_keys(password)
        time.sleep(1)
        driver.find_element(By.CLASS_NAME, "btn-outline-success").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//*[@id='root']/div/div[2]/a").click()
        driver.implicitly_wait(20)
        update_log("Login berhasil.")

        update_log(f"Memilih menu {selected_rl}...")

        if selected_rl == "RL 4.1":
            try:
                rl4_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@role='button' and text()='RL.4']")))
                driver.execute_script("arguments[0].scrollIntoView(true);", rl4_element)
                driver.execute_script("arguments[0].click();", rl4_element)
                time.sleep(2)
                dropdown_item = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'RL 4.1 Morbiditas Pasien Rawat Inap')]")))
                dropdown_item.click()
            except Exception:
                update_log("Gagal klik RL.4, mencoba lagi...")
                time.sleep(2)
                rl4_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@role='button' and text()='RL.4']")))
                driver.execute_script("arguments[0].click();", rl4_element)
                dropdown_item = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'RL 4.1 Morbiditas Pasien Rawat Inap')]")))
                dropdown_item.click()

        elif selected_rl == "RL 5.1":
            try:
                rl5_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@role='button' and text()='RL.5']")))
                driver.execute_script("arguments[0].scrollIntoView(true);", rl5_element)
                driver.execute_script("arguments[0].click();", rl5_element)
                time.sleep(2)
                dropdown_item = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'RL 5.1 Mobiditas Pasien Rawat Jalan')]")))
                dropdown_item.click()
            except Exception:
                update_log("Gagal klik RL.5, mencoba lagi...")
                time.sleep(2)
                rl5_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@role='button' and text()='RL.5']")))
                driver.execute_script("arguments[0].click();", rl5_element)
                dropdown_item = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'RL 5.1 Mobiditas Pasien Rawat Jalan')]")))
                dropdown_item.click()
        
        update_log("Menu berhasil dipilih.")
        time.sleep(2)
        button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'tambah')]"))
        )
        driver.execute_script("arguments[0].click();", button)

        month_mapping = {"January": 2, "February": 3, "March": 4, "April": 5, "May": 6, "June": 7, "July": 8, "August": 9, "September": 10, "October": 11, "November": 12, "December": 13}
        selected_month = month_choice.get()
        month_index = month_mapping[selected_month]

        update_log("Memulai proses input data per baris...")
        for i, row in df.iterrows():
            progress = (i + 1) / total_rows
            update_progress(progress)
            
            original_icd = str(row[1])
            icd_data = original_icd
            
            if '.' in icd_data:
                parts = icd_data.split('.')
                if len(parts) > 1 and len(parts[1]) > 1:
                    icd_data = f"{parts[0]}.{parts[1][0]}"
                    update_log(f"INFO: Kode ICD '{original_icd}' disederhanakan menjadi '{icd_data}' untuk pencarian.")

            update_log(f"({i+1}/{total_rows}) Memproses ICD: {icd_data}")

            input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@name='caripenyakit']")))
            driver.execute_script("arguments[0].scrollIntoView(true);", input_element)
            driver.execute_script("arguments[0].focus();", input_element)
            input_element.clear()
            input_element.send_keys(icd_data)
            input_element.send_keys(Keys.RETURN)

            try:
                # Coba temukan dan klik tombol 'Tambah'
                add_button = driver.find_element(By.XPATH, "//*[@id='root']/div/div[2]/div/div/div/div[2]/table/tbody/tr[1]/td[4]/button")
                add_button.click()
                time.sleep(2)
                
            except NoSuchElementException:
                # Jika tombol 'Tambah' tidak ditemukan, artinya ICD tidak ada
                update_log(f"PERINGATAN: ICD '{icd_data}' tidak ditemukan di website. Baris ini dilewati.")
                
                # Kosongkan lagi input field untuk persiapan ICD berikutnya
                input_element.clear() 
                
                # Lanjutkan ke iterasi/baris berikutnya di file Excel
                continue
            

            try:
                month_dropdown_xpath = f"//*[@id='bulan']/option[{month_index}]"
                month_dropdown_element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, month_dropdown_xpath)))
                month_dropdown_element.click()
                time.sleep(2)
            except Exception as e:
                update_log(f"ERROR: Gagal memilih bulan untuk {icd_data}. Menghentikan proses.")
                raise e

            # =====================================================================
            # === MULAI PERUBAHAN: Looping pengisian data Laki-Laki & Perempuan ===
            # =====================================================================
            for j in range(25):
                # Ambil data dari excel
                male_data = row.iloc[3 + 2 * j]
                female_data = row.iloc[4 + 2 * j]
                
                # Definisikan XPATH untuk baris saat ini
                male_xpath = f"//*[@id='root']/div/div[2]/div[2]/div/div/form/div[2]/div[2]/table/tbody/tr[{j + 1}]/td[3]/input"
                female_xpath = f"//*[@id='root']/div/div[2]/div[2]/div/div/form/div[2]/div[2]/table/tbody/tr[{j + 1}]/td[4]/input"
                
                # Proses input untuk kolom Laki-Laki (jika aktif)
                try:
                    male_element = driver.find_element(By.XPATH, male_xpath)
                    if male_element.is_enabled():
                        driver.execute_script("arguments[0].scrollIntoView(true);", male_element)
                        male_element.clear()
                        male_element.send_keys(str(male_data))
                except Exception:
                    # Jika elemen tidak ditemukan atau tidak bisa diisi, lewati saja.
                    pass
                    
                # Proses input untuk kolom Perempuan (jika aktif)
                try:
                    female_element = driver.find_element(By.XPATH, female_xpath)
                    if female_element.is_enabled():
                        driver.execute_script("arguments[0].scrollIntoView(true);", female_element)
                        female_element.clear()
                        female_element.send_keys(str(female_data))
                except Exception:
                    # Jika elemen tidak ditemukan atau tidak bisa diisi, lewati saja.
                    pass

            # --- PENYESUAIAN ILOC BERDASARKAN RL UNTUK BARIS TERAKHIR ---
            if selected_rl == "RL 4.1":
                male_last_index = 55
                female_last_index = 56
            else: # Asumsi RL 5.1
                male_last_index = 56 
                female_last_index = 57
            
            male_last_data = row.iloc[male_last_index]
            female_last_data = row.iloc[female_last_index]
            
            # Definisikan XPATH untuk baris terakhir (baris ke-26)
            male_last_xpath = "//*[@id='root']/div/div[2]/div[2]/div/div/form/div[2]/div[2]/table/tbody/tr[26]/td[3]/input"
            female_last_xpath = "//*[@id='root']/div/div[2]/div[2]/div/div/form/div[2]/div[2]/table/tbody/tr[26]/td[4]/input"

            # Proses input untuk kolom Laki-Laki terakhir (jika aktif)
            try:
                male_last_element = driver.find_element(By.XPATH, male_last_xpath)
                if male_last_element.is_enabled():
                    driver.execute_script("arguments[0].scrollIntoView(true);", male_last_element)
                    male_last_element.clear()
                    male_last_element.send_keys(str(male_last_data))
            except Exception:
                pass
            
            # Proses input untuk kolom Perempuan terakhir (jika aktif)
            try:
                female_last_element = driver.find_element(By.XPATH, female_last_xpath)
                
                if female_last_element.is_enabled():
                    driver.execute_script("arguments[0].scrollIntoView(true);", female_last_element)
                    female_last_element.clear()
                    female_last_element.send_keys(str(female_last_data))
            except Exception:
                pass
            # =====================================================================
            # === AKHIR PERUBAHAN =================================================
            # =====================================================================

            time.sleep(2)
            tombol_simpan = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Simpan')]"))
            )

            driver.execute_script("arguments[0].click();", tombol_simpan)
            time.sleep(4)
            button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'tambah')]"))
            )
            driver.execute_script("arguments[0].click();", button)
        
        update_log("\n✓✓✓ SEMUA PROSES SELESAI ✓✓✓")

    except Exception as e:
        update_log(f"\nERROR: Terjadi kesalahan. Proses dihentikan. Detail: {e}")
        update_log("-" * 50)
        update_log("SARAN PENANGANAN:")
        update_log("1. Tutup dan jalankan ulang aplikasi ini.")
        update_log("2. Buka file Excel Anda, lalu HAPUS semua baris data yang sudah berhasil diinput.")
        update_log("3. Simpan file Excel tersebut.")
        update_log("4. Coba jalankan kembali proses ini menggunakan file Excel yang sudah diperbarui.")
        update_log("-" * 50)
    finally:
        if driver:
            driver.quit()
        # Mengaktifkan kembali tombol setelah selesai
        start_button.configure(state="normal")

# Fungsi untuk memulai proses di thread terpisah
def start_process():
    global file_path
    
    if not file_path or not password_entry.get():
        if not file_path:
            file_label.configure(text="Harap pilih file Excel!", text_color="red")
        if not password_entry.get():
            update_log("ERROR: Harap masukkan password Anda sebelum memulai proses.")
        return
    
    start_button.configure(state="disabled")
    clear_log()
    update_progress(0)
    
    tab_view.set("Log")
    
    process_thread = threading.Thread(target=run_selenium_process)
    process_thread.daemon = True
    process_thread.start()

# --- GUI Setup ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Project Altair")
root.geometry("600x780")

main_frame = ctk.CTkFrame(master=root)
main_frame.pack(pady=20, padx=20, fill="both", expand=True)

title_label = ctk.CTkLabel(master=main_frame, text="SIRS INPUT AUTOMATION", font=ctk.CTkFont(size=20, weight="bold"))
title_label.pack(pady=(10, 20), padx=20)

tab_view = ctk.CTkTabview(master=main_frame)
tab_view.pack(padx=20, pady=10, fill="both", expand=True)

tab_utama = tab_view.add("Proses Utama")
tab_pengaturan = tab_view.add("Pengaturan")
tab_log = tab_view.add("Log")

# --- Konten untuk Tab Proses Utama ---
login_frame = ctk.CTkFrame(master=tab_utama)
login_frame.pack(pady=10, padx=10, fill="x")
login_header = ctk.CTkLabel(master=login_frame, text="Informasi Login", font=ctk.CTkFont(weight="bold"))
login_header.pack(pady=(10, 5), padx=10, anchor="w")

email_entry = ctk.CTkEntry(master=login_frame)
email_entry.insert(0, "rsudsultansuriansyah@gmail.com")
email_entry.configure(state="disabled")
email_entry.pack(pady=5, padx=10, fill="x")

password_entry = ctk.CTkEntry(master=login_frame, placeholder_text="Masukkan Password Akun", show="*")
password_entry.pack(pady=(5, 5), padx=10, fill="x")

show_password_var = ctk.BooleanVar() 
show_password_checkbox = ctk.CTkCheckBox(master=login_frame, text="Tampilkan Password", variable=show_password_var, command=toggle_password)
show_password_checkbox.pack(pady=(0, 10), padx=10, anchor="w")

process_frame = ctk.CTkFrame(master=tab_utama)
process_frame.pack(pady=10, padx=10, fill="x")
process_header = ctk.CTkLabel(master=process_frame, text="Pengaturan Proses", font=ctk.CTkFont(weight="bold"))
process_header.pack(pady=(10, 10), padx=10, anchor="w")
file_button = ctk.CTkButton(master=process_frame, text="Pilih File Excel", command=select_file)
file_button.pack(pady=10, padx=10, fill="x")
file_label = ctk.CTkLabel(master=process_frame, text="Belum ada file yang dipilih", wraplength=350, justify="center")
file_label.pack(pady=(0, 10), padx=10)
rl_label = ctk.CTkLabel(master=process_frame, text="Jenis RL")
rl_label.pack(pady=(10, 0), padx=10, anchor="w")
rl_choice = ctk.StringVar(value="RL 4.1")
rl_dropdown = ctk.CTkComboBox(master=process_frame, values=["RL 4.1", "RL 5.1"], variable=rl_choice)
rl_dropdown.pack(pady=(0, 10), padx=10, fill="x")
month_label = ctk.CTkLabel(master=process_frame, text="Bulan Pelaporan")
month_label.pack(pady=(10, 0), padx=10, anchor="w")
month_choice = ctk.StringVar(value="January")
month_dropdown = ctk.CTkComboBox(
    master=process_frame,
    values=["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    variable=month_choice
)
month_dropdown.pack(pady=(0, 15), padx=10, fill="x")

start_button = ctk.CTkButton(master=tab_utama, text="Mulai Proses", command=start_process, height=40, font=ctk.CTkFont(size=14, weight="bold"))
start_button.pack(pady=20, padx=100, fill="x", side="bottom")

# --- Konten untuk Tab Pengaturan ---
driver_frame = ctk.CTkFrame(master=tab_pengaturan)
driver_frame.pack(pady=10, padx=10, fill="both", expand=True)
driver_header = ctk.CTkLabel(master=driver_frame, text="Pengaturan Driver (Cadangan)", font=ctk.CTkFont(weight="bold"))
driver_header.pack(pady=(10, 10), padx=10, anchor="w")
driver_info = ctk.CTkLabel(master=driver_frame, text="Gunakan ini HANYA jika setup driver otomatis gagal. Pilih file chromedriver.exe yang sesuai.", wraplength=400, justify="left")
driver_info.pack(pady=(0, 10), padx=10, anchor="w")
driver_button = ctk.CTkButton(master=driver_frame, text="Pilih Path Driver Manual", command=select_driver_path)
driver_button.pack(pady=10, padx=10, fill="x")
driver_label = ctk.CTkLabel(master=driver_frame, text="Belum ada driver manual yang dipilih", wraplength=350, justify="center")
driver_label.pack(pady=(0, 15), padx=10)

# --- Konten untuk Tab Log ---
status_frame = ctk.CTkFrame(master=tab_log)
status_frame.pack(pady=10, padx=10, fill="both", expand=True)

progress_container = ctk.CTkFrame(master=status_frame, fg_color="transparent")
progress_container.pack(pady=10, padx=10, fill="x")

progress_bar = ctk.CTkProgressBar(master=progress_container)
progress_bar.set(0)
progress_bar.pack(side="left", fill="x", expand=True, padx=(0, 10))

clear_log_button = ctk.CTkButton(master=progress_container, text="Hapus Log", command=clear_log, width=100)
clear_log_button.pack(side="left")

log_textbox = ctk.CTkTextbox(master=status_frame, state="disabled", wrap="word")
log_textbox.pack(pady=(0, 10), padx=10, fill="both", expand=True)

# --- Label Nama Pembuat & Versi ---
creator_label = ctk.CTkLabel(master=main_frame, text="Developed by: Aldy Pradana, A.Md.RMIK |  v2.0", font=ctk.CTkFont(size=10))
creator_label.pack(pady=(10, 5), padx=20, side="bottom", anchor="e")

# --- Pesan Sambutan Awal di Log ---
update_log("Selamat datang! Silakan isi semua pengaturan dan klik 'Mulai Proses'.")

root.mainloop()



