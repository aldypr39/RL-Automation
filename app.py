# %%
import sys
import os
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
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException
import pandas as pd
import time
import threading
from PIL import Image

# --- Variabel Global ---
file_path = ""
chrome_driver_path = ""

def resource_path(relative_path):
    """ Mendapatkan path absolut ke resource, bekerja untuk dev dan PyInstaller """
    try:
        # PyInstaller membuat folder temp dan menyimpan path di _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# --- Fungsi GUI (Tidak Diubah) ---
def select_file():
    global file_path
    file_path = filedialog.askopenfilename(title="Pilih File Excel", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
    if file_path:
        file_name = file_path.split("/")[-1]
        file_label.configure(text=f"File Terpilih: {file_name}")
    else:
        file_label.configure(text="Belum ada file yang dipilih")
def select_driver_path():
    global chrome_driver_path
    chrome_driver_path = filedialog.askopenfilename(title="Pilih File ChromeDriver", filetypes=[("Executable files", "*.exe"), ("All files", "*.*")])
    if chrome_driver_path:
        driver_name = chrome_driver_path.split("/")[-1]
        driver_label.configure(text=f"Driver Manual: {driver_name}")
    else:
        driver_label.configure(text="Belum ada driver manual yang dipilih")
def update_log(message):
    log_textbox.configure(state="normal")
    log_textbox.insert("end", message + "\n")
    log_textbox.configure(state="disabled")
    log_textbox.see("end")
def update_progress(value):
    progress_bar.set(value)
def clear_log():
    log_textbox.configure(state="normal")
    log_textbox.delete("1.0", "end")
    log_textbox.configure(state="disabled")
def toggle_password():
    if show_password_var.get():
        password_entry.configure(show="")
    else:
        password_entry.configure(show="*")

# --- Fungsi Bantuan untuk Recovery (Tidak Diubah) ---
def recover_and_re_navigate(driver, selected_rl):
    update_log("--> Error terdeteksi. Mencoba recovery dengan navigasi ulang...")
    try:
        update_log("--> Memaksa kembali ke halaman menu utama SIRS...")
        sirs_home_url = "https://sirs6.kemkes.go.id/v3/"
        driver.get(sirs_home_url)
        time.sleep(3) 

        if selected_rl == "RL 4.1":
            rl_element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@role='button' and text()='RL.4']")))
            driver.execute_script("arguments[0].click();", rl_element)
            time.sleep(1.5)
            dropdown_item = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'RL 4.1 Morbiditas Pasien Rawat Inap')]")))
            dropdown_item.click()
        elif selected_rl == "RL 5.1":
            rl_element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@role='button' and text()='RL.5']")))
            driver.execute_script("arguments[0].click();", rl_element)
            time.sleep(1.5)
            dropdown_item = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'RL 5.1 Mobiditas Pasien Rawat Jalan')]")))
            dropdown_item.click()
        
        time.sleep(2)
        button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'tambah')]")))
        driver.execute_script("arguments[0].click();", button)
        update_log("--> Recovery berhasil. Melanjutkan proses pada baris yang sama.")
    except Exception as e:
        update_log(f"(!) Gagal melakukan recovery. Error: {e}")
        raise e

# --- Fungsi Bantuan untuk Input (Tidak Diubah) ---
def robust_clear_and_send_keys(driver, element_xpath, text_to_send):
    try:
        element = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, element_xpath)))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(0.5)
        element.click()
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)
        time.sleep(0.5)
        element.send_keys(text_to_send)
    except Exception as e:
        update_log(f"(!) Gagal membersihkan atau mengirim teks ke elemen. Error: {e}")
        raise e

# Fungsi utama yang menjalankan proses Selenium
def run_selenium_process():
    global chrome_driver_path
    driver = None
    try:
        email = email_entry.get()
        password = password_entry.get()
        selected_rl = rl_choice.get() 
        
        # Setup Excel dan Driver (Tidak Diubah)
        update_log("Membaca file Excel...")
        if selected_rl == "RL 5.1":
            df = pd.read_excel(file_path, header=4)
        else:
            df = pd.read_excel(file_path, header=3)
        total_rows = len(df)
        update_log(f"Ditemukan {total_rows} baris data untuk diproses.")
        try:
            update_log("Mencoba setup ChromeDriver otomatis...")
            service = ChromeService(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service)
        except Exception as e_auto:
            update_log(f"(!) Gagal setup otomatis: {e_auto}")
            if chrome_driver_path:
                update_log("Mencoba menggunakan driver manual...")
                service = ChromeService(executable_path=chrome_driver_path)
                driver = webdriver.Chrome(service=service)
            else:
                update_log("ERROR: Setup otomatis gagal, tidak ada driver manual dipilih.")
                start_button.configure(state="normal")
                return
        
        # Login dan Navigasi Awal (Tidak Diubah)
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
        button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'tambah')]")))
        driver.execute_script("arguments[0].click();", button)

        month_mapping = {"January": 2, "February": 3, "March": 4, "April": 5, "May": 6, "June": 7, "July": 8, "August": 9, "September": 10, "October": 11, "November": 12, "December": 13}
        selected_month = month_choice.get()
        month_index = month_mapping[selected_month]

        update_log("Memulai proses input data per baris...")
        for i, row in df.iterrows():
            row_processed_successfully = False
            attempts_for_this_row = 0
            while not row_processed_successfully and attempts_for_this_row < 3:
                try:
                    update_log("-" * 20)
                    update_log(f"Memproses baris {i + 1}/{total_rows}")
                    
                    original_icd = str(row[1])
                    icd_data = original_icd 
                    if '.' in icd_data and len(icd_data.split('.')) > 1 and len(icd_data.split('.')[1]) > 1:
                        parts = icd_data.split('.')
                        icd_data = f"{parts[0]}.{parts[1][0]}"
                    
                    update_log(f"({i+1}/{total_rows}) Mencari ICD: {icd_data}")
                    icd_input_xpath = "//input[@name='caripenyakit']"
                    robust_clear_and_send_keys(driver, icd_input_xpath, icd_data)
                    driver.find_element(By.XPATH, icd_input_xpath).send_keys(Keys.RETURN)
                    
                    # Logika Pengecekan Cerdas (Tidak Diubah)
                    try:
                        # Definisikan XPath untuk kedua kemungkinan hasil
                        first_result_row_xpath = "//*[@id='root']/div/div[2]/div/div/div/div[2]/table/tbody/tr[1]"
                        not_found_message_xpath = "//td[contains(text(), 'Data tidak ditemukan')]"

                        # Tunggu hingga SALAH SATU dari dua elemen ini muncul
                        wait = WebDriverWait(driver, 20)
                        found_element = wait.until(
                            EC.presence_of_element_located(
                                (By.XPATH, f"{first_result_row_xpath} | {not_found_message_xpath}")
                            )
                        )

                        # Periksa elemen mana yang ditemukan berdasarkan tag-nya
                        if found_element.tag_name == 'tr':
                            # Jika yang ditemukan adalah baris (tr), klik tombol Tambah di dalamnya
                            add_button = found_element.find_element(By.XPATH, ".//td[4]/button")
                            add_button.click()
                        else: # Berarti tag-nya adalah 'td'
                            # Jika yang ditemukan adalah sel tabel (td), itu adalah pesan error
                            update_log(f"PERINGATAN: ICD '{icd_data}' tidak valid (Data tidak ditemukan). Baris ini dilewati.")
                            row_processed_successfully = True
                            continue
                            
                    except TimeoutException:
                        # Jika setelah 20 detik TIDAK ADA SATU PUN dari kedua elemen itu yang muncul, 
                        # barulah kita bisa yakin halaman benar-benar macet.
                        update_log(f"(!) Halaman 'stuck' untuk ICD '{icd_data}'. Memicu recovery...")
                        raise TimeoutException("Halaman pencarian macet.")

                    # Sisa proses pengisian form (Tidak Diubah)
                    month_dropdown_xpath = f"//*[@id='bulan']/option[{month_index}]"
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, month_dropdown_xpath))).click()
                    time.sleep(2)
                    for j in range(25):
                        male_data, female_data = row.iloc[3 + 2 * j], row.iloc[4 + 2 * j]
                        male_xpath = f"//*[@id='root']/div/div[2]/div[2]/div/div/form/div[2]/div[2]/table/tbody/tr[{j + 1}]/td[3]/input"
                        female_xpath = f"//*[@id='root']/div/div[2]/div[2]/div/div/form/div[2]/div[2]/table/tbody/tr[{j + 1}]/td[4]/input"
                        try:
                            male_element = driver.find_element(By.XPATH, male_xpath)
                            if male_element.is_enabled(): male_element.send_keys(str(male_data))
                        except: pass
                        try:
                            female_element = driver.find_element(By.XPATH, female_xpath)
                            if female_element.is_enabled(): female_element.send_keys(str(female_data))
                        except: pass
                    male_last_idx, female_last_idx = (55, 56) if selected_rl == "RL 4.1" else (56, 57)
                    male_last_data, female_last_data = row.iloc[male_last_idx], row.iloc[female_last_idx]
                    try: driver.find_element(By.XPATH, "//*[@id='root']/div/div[2]/div[2]/div/div/form/div[2]/div[2]/table/tbody/tr[26]/td[3]/input").send_keys(str(male_last_data))
                    except: pass
                    try: driver.find_element(By.XPATH, "//*[@id='root']/div/div[2]/div[2]/div/div/form/div[2]/div[2]/table/tbody/tr[26]/td[4]/input").send_keys(str(female_last_data))
                    except: pass
                    time.sleep(2)
                    tombol_simpan = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Simpan')]")))
                    driver.execute_script("arguments[0].click();", tombol_simpan)
                    time.sleep(4)
                    button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'tambah')]")))
                    driver.execute_script("arguments[0].click();", button)
                    row_processed_successfully = True
                    update_progress((i + 1) / total_rows)
                    time.sleep(2)
                except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
                    attempts_for_this_row += 1
                    update_log(f"(!) Gagal pada baris {i + 1}. Error: {str(e).splitlines()[0]}")
                    if attempts_for_this_row < 3:
                        recover_and_re_navigate(driver, selected_rl)
                    else:
                        update_log(f"(!) Gagal total memproses baris {i + 1}. Melewati baris ini.")
                        break
        
        update_log("\n✓✓✓ SEMUA PROSES SELESAI ✓✓✓")
    except Exception as e:
        update_log(f"\nERROR: Terjadi kesalahan. Proses dihentikan. Detail: {e}")
        update_log("-" * 50)
        update_log("SARAN PENANGANAN:\n1. Tutup dan jalankan ulang aplikasi ini.\n2. Hapus baris yang sudah berhasil diinput dari file Excel.\n3. Simpan file Excel.\n4. Jalankan kembali proses.")
    finally:
        if driver:
            driver.quit()
        start_button.configure(state="normal")

# --- GUI Setup (Tidak diubah) ---
def start_process():
    global file_path
    if not file_path or not password_entry.get():
        if not file_path: file_label.configure(text="Harap pilih file Excel!", text_color="red")
        if not password_entry.get(): update_log("ERROR: Harap masukkan password Anda.")
        return
    
    start_button.configure(state="disabled")
    clear_log()
    update_progress(0)
    tab_view.set("Log")
    
    process_thread = threading.Thread(target=run_selenium_process)
    process_thread.daemon = True
    process_thread.start()

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")
root = ctk.CTk()
root.title("AUTOMASI LAPORAN TERINTEGRASI SIRS")
root.geometry("600x860")
main_frame = ctk.CTkFrame(master=root)
main_frame.pack(pady=20, padx=20, fill="both", expand=True)
logo_path = resource_path("logoaltair.png")
logo_data = Image.open(logo_path)
# Buat CTkImage
root.logo_image = ctk.CTkImage(dark_image=logo_data, light_image=logo_data, size=(150, 115))

# Buat label dan gunakan gambar yang sudah disimpan di root
logo_label = ctk.CTkLabel(master=main_frame, text="", image=root.logo_image)
logo_label.pack(pady=(20, 20), padx=20)

tab_view = ctk.CTkTabview(master=main_frame)
tab_view.pack(padx=20, pady=10, fill="both", expand=True)
tab_utama = tab_view.add("Proses Utama")
tab_pengaturan = tab_view.add("Pengaturan")
tab_log = tab_view.add("Log")
login_frame = ctk.CTkFrame(master=tab_utama)
login_frame.pack(pady=10, padx=10, fill="x")
login_header = ctk.CTkLabel(master=login_frame, text="Informasi Login", font=ctk.CTkFont(weight="bold"))
login_header.pack(pady=(10, 5), padx=10, anchor="w")
email_entry = ctk.CTkEntry(master=login_frame)
email_entry.insert(0, "rsudulinprovkalsel@gmail.com")
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
month_dropdown = ctk.CTkComboBox(master=process_frame, values=["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"], variable=month_choice)
month_dropdown.pack(pady=(0, 15), padx=10, fill="x")
start_button = ctk.CTkButton(master=tab_utama, text="Mulai Proses", command=start_process, height=40, font=ctk.CTkFont(size=14, weight="bold"))
start_button.pack(pady=20, padx=100, fill="x", side="bottom")
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
creator_label = ctk.CTkLabel(master=main_frame, text="Developed by: Aldy Pradana, A.Md.RMIK |  v3.2", font=ctk.CTkFont(size=10))
creator_label.pack(pady=(10, 5), padx=20, side="bottom", anchor="e")
update_log("Selamat datang! Silakan isi semua pengaturan dan klik 'Mulai Proses'.")
root.mainloop()




