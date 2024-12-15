import re
import pdfplumber
import pandas as pd
import os
import shutil
from datetime import datetime

def extract_invoice_data(data_folder="Data", output_folder="Result"):
    # Direktori sumber dan backup
    backup_folder = os.path.join(data_folder, "Backup")

    # Pastikan folder backup dan output ada
    os.makedirs(backup_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    # Regex patterns for extraction
    patterns = {
        "Invoice": r"INVOICE\s(\d+)",
        "Invoice_Date": r"Invoice Date\s?:\s?([0-9]{2}\.[A-Za-z]{3}\.[0-9]{4})",
        "Due_Date": r"Due Date\s?:\s?([0-9]{2}\.[A-Za-z]{3}\.[0-9]{4})",
        "Customer_No": r"Customer No\s?:\s?([A-Za-z0-9]+)",
        "Net_Value": r"Net value\s([\d,]+)",
        "Amount_Due": r"Amount Due\s([\d,]+)",
        "FCR": r"FCR\s([A-Za-z0-9]+)",
        "Total_Packages": r"Total Packages\s?:\s?(\d+)",
        "Weight": r"Weight\s?:\s?([\d.]+)",
        "Volume": r"Volume\s?:\s?([\d.]+)",
        "Units": r"Units\s?:\s?(\d+)",
    }

    # List untuk menyimpan hasil ekstraksi dari semua file
    all_data = []

    # Proses semua file di folder Data
    for filename in os.listdir(data_folder):
        if filename.endswith(".pdf"):  # Hanya proses file PDF
            file_path = os.path.join(data_folder, filename)
            
            # Ekstrak data dari file PDF
            data = {}
            try:
                with pdfplumber.open(file_path) as pdf:
                    text = pdf.pages[0].extract_text()
                    for key, pattern in patterns.items():
                        match = re.search(pattern, text)
                        data[key] = match.group(1) if match else None
            except Exception as e:
                print(f"Error membaca file '{filename}': {e}")
                continue

            # Cetak data sebelum konversi
            print(f"\nHasil ekstraksi sebelum konversi dari file '{filename}':")
            for key, value in data.items():
                print(f"{key}: {value}")
            print("=" * 50)

            # Konversi tipe data
            processed_data = {
                "Filename": filename,  # Tambahkan nama file
                "Invoice": data.get("Invoice"),
                "Invoice_Date": pd.to_datetime(data.get("Invoice_Date"), format="%d.%b.%Y") if data.get("Invoice_Date") else None,
                "Due_Date": pd.to_datetime(data.get("Due_Date"), format="%d.%b.%Y") if data.get("Due_Date") else None,
                "Customer_No": data.get("Customer_No"),
                "Net_Value": float(data.get("Net_Value").replace(",", "")) if data.get("Net_Value") else None,
                "Amount_Due": float(data.get("Amount_Due").replace(",", "")) if data.get("Amount_Due") else None,
                "FCR": data.get("FCR"),
                "Total_Packages": int(data.get("Total_Packages")) if data.get("Total_Packages") else None,
                "Weight": float(data.get("Weight")) if data.get("Weight") else None,
                "Volume": float(data.get("Volume").replace(".", "").replace(",", ".")) if data.get("Volume") else None,
                "Units": int(data.get("Units")) if data.get("Units") else None,
            }
            
            # Cetak data setelah konversi
            print(f"\nHasil ekstraksi setelah konversi dari file '{filename}':")
            for key, value in processed_data.items():
                print(f"{key}: {value}")
            print("=" * 50)

            # Tambahkan ke list semua data
            all_data.append(processed_data)
            
            # Pindahkan file ke folder backup
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"executed_{os.path.splitext(filename)[0]}_{timestamp}.pdf"
            shutil.move(file_path, os.path.join(backup_folder, backup_filename))

    # Jika ada data, simpan ke file Excel
    if all_data:
        # Buat DataFrame dari semua data
        df = pd.DataFrame(all_data)
        
        # Simpan ke Excel dengan nama file berbasis timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(output_folder, f"result_{timestamp}.xlsx")
        df.to_excel(output_file, index=False)
        print(f"\nData berhasil disimpan ke {output_file}")
    else:
        print("Tidak ada file PDF yang ditemukan di folder Data.")

if __name__ == "__main__":
    extract_invoice_data()
