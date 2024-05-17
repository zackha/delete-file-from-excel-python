import openpyxl
import os
import tkinter as tk
from tkinter import filedialog

# Tkinter'i başlatın ve dosya seçme penceresi açın
root = tk.Tk()
root.withdraw()

# Excel dosyasını seçmek için dosya seçme penceresi açın
excel_dosyasi_yolu = filedialog.askopenfilename(
    title="Excel dosyasını seçin",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

# Eğer kullanıcı Excel dosyası seçmezse, işlem sonlandırılır
if not excel_dosyasi_yolu:
    print("Excel dosyası seçilmedi. İşlem iptal edildi.")
else:
    # Klasör seçmek için dosya seçme penceresi açın
    klasor_yolu = filedialog.askdirectory(title="Klasörü seçin")

    # Eğer kullanıcı klasör seçmezse, işlem sonlandırılır
    if not klasor_yolu:
        print("Klasör seçilmedi. İşlem iptal edildi.")
    else:
        # Excel dosyasını açın
        wb = openpyxl.load_workbook(excel_dosyasi_yolu)
        ws = wb.active

        # A sütunundaki dosya isimlerini alın
        dosya_isimleri = []
        for row in ws.iter_rows(min_row=1, min_col=1, max_col=1, values_only=True):
            dosya_isimleri.append(row[0])

        # Dosyaları klasörden silin
        for dosya_adi in dosya_isimleri:
            if dosya_adi:  # Dosya adı boş değilse işlem yap
                dosya_yolu = os.path.join(klasor_yolu, dosya_adi)
                if os.path.exists(dosya_yolu):
                    os.remove(dosya_yolu)
                    print(f"{dosya_adi} silindi.")
                else:
                    print(f"{dosya_adi} bulunamadı.")

        print("İşlem tamamlandı.")
