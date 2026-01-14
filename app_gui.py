import os
import sys
import time
import threading
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

# =====================
# FUNGSI PATH (AMAN UNTUK EXE)
# =====================
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# =====================
# KONFIGURASI SELENIUM
# =====================
URL = "https://matchapro.web.bps.go.id/dirgc"
DELAY = 3

COL_IDSBR = 0
COL_LAT = 10
COL_LONG = 11
COL_KEBERADAAN = 12

# =====================
# GUI APP
# =====================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GC Automation - IPDS BPS Kabupaten Buleleng")
        self.geometry("750x520")
        self.resizable(False, False)

        self.file_path = tk.StringVar()
        self.nama_petugas = tk.StringVar()
        self.satker = tk.StringVar(value="BPS Kabupaten Buleleng")

        self.create_widgets()

    def create_widgets(self):
        frame = tk.Frame(self, padx=10, pady=10)
        frame.pack(fill="both")

        tk.Label(frame, text="File SBR (.xlsx)").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.file_path, width=60).grid(row=1, column=0, padx=5)
        tk.Button(frame, text="Browse", command=self.browse_file).grid(row=1, column=1)

        tk.Label(frame, text="Nama Petugas").grid(row=2, column=0, sticky="w", pady=(10, 0))
        tk.Entry(frame, textvariable=self.nama_petugas, width=40).grid(row=3, column=0, sticky="w")

        tk.Label(frame, text="Satker").grid(row=4, column=0, sticky="w", pady=(10, 0))
        tk.Entry(frame, textvariable=self.satker, width=40, state="readonly").grid(row=5, column=0, sticky="w")

        self.btn_start = tk.Button(frame, text="MULAI PROSES", bg="#2ecc71", fg="white",
                                   width=20, command=self.start_thread)
        self.btn_start.grid(row=6, column=0, pady=15, sticky="w")

        tk.Label(frame, text="Log Proses").grid(row=7, column=0, sticky="w")
        self.log = scrolledtext.ScrolledText(frame, width=85, height=15, state="disabled")
        self.log.grid(row=8, column=0, columnspan=2)

    def browse_file(self):
        file = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if file:
            self.file_path.set(file)

    def write_log(self, text):
        self.log.config(state="normal")
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.log.config(state="disabled")

    def start_thread(self):
        if not self.file_path.get():
            messagebox.showerror("Error", "Silakan pilih file SBR.xlsx")
            return
        if not self.nama_petugas.get():
            messagebox.showerror("Error", "Silakan isi Nama Petugas")
            return

        self.btn_start.config(state="disabled")
        threading.Thread(target=self.run_selenium, daemon=True).start()

    def run_selenium(self):
        try:
            df = pd.read_excel(self.file_path.get())
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca Excel:\n{e}")
            self.btn_start.config(state="normal")
            return

        hasil_file = os.path.join(os.path.dirname(self.file_path.get()), "hasil.xlsx")

        if os.path.exists(hasil_file):
            df_hasil = pd.read_excel(hasil_file)
        else:
            df_hasil = pd.DataFrame(
                columns=["idsbr", "waktu", "status", "petugas", "satker"]
            )

        options = webdriver.ChromeOptions()
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")

        service = Service(resource_path("chromedriver.exe"))
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 10)

        driver.get(URL)
        self.write_log("Silakan login terlebih dahulu...")
        time.sleep(40)

        try:
            driver.find_element(By.ID, "toggle-filter").click()
            time.sleep(DELAY)
        except:
            pass

        for _, row in df.iterrows():
            idsbr = str(row.iloc[COL_IDSBR])
            lat = str(row.iloc[COL_LAT])
            lon = str(row.iloc[COL_LONG])
            keberadaan = str(row.iloc[COL_KEBERADAAN])

            status = "SKIP, KEMUNGKINAN SUDAH DILAKUKAN GC"
            waktu = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            self.write_log(f"Proses IDSBR: {idsbr}")

            try:
                driver.execute_script("window.scrollTo(0, 0);")
                time.sleep(DELAY)

                search = wait.until(EC.presence_of_element_located((By.ID, "search-idsbr")))
                search.clear()
                search.send_keys(idsbr)
                time.sleep(DELAY)

                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(DELAY)

                cards = driver.find_elements(By.CLASS_NAME, "usaha-card-header")
                if not cards:
                    raise Exception("Usaha tidak ditemukan")

                cards[0].click()
                time.sleep(DELAY)

                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(DELAY)

                btn = driver.find_elements(By.CLASS_NAME, "btn-tandai")
                if not btn:
                    raise Exception("KEMUNGKINAN SUDAH DILAKUKAN GC")

                btn[0].click()
                time.sleep(DELAY)

                Select(wait.until(
                    EC.presence_of_element_located((By.ID, "tt_hasil_gc"))
                )).select_by_value(keberadaan)

                driver.find_element(By.ID, "tt_latitude_cek_user").clear()
                driver.find_element(By.ID, "tt_latitude_cek_user").send_keys(lat)

                driver.find_element(By.ID, "tt_longitude_cek_user").clear()
                driver.find_element(By.ID, "tt_longitude_cek_user").send_keys(lon)

                driver.find_element(By.ID, "save-tandai-usaha-btn").click()
                time.sleep(DELAY)

                driver.find_element(By.CLASS_NAME, "swal2-confirm").click()
                time.sleep(DELAY)

                status = "BERHASIL"
                self.write_log(f"IDSBR {idsbr} : BERHASIL")

            except Exception as e:
                self.write_log(f"IDSBR {idsbr} : SKIP, {e}")

            df_hasil.loc[len(df_hasil)] = [
                idsbr, waktu, status, self.nama_petugas.get(), self.satker.get()
            ]
            df_hasil.to_excel(hasil_file, index=False)

        driver.quit()
        self.write_log("=== SEMUA PROSES SELESAI ===")
        messagebox.showinfo("Selesai", "Proses DIRGC telah selesai")
        self.btn_start.config(state="normal")


if __name__ == "__main__":
    app = App()
    app.mainloop()
