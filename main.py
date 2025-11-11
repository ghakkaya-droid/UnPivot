import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog


def select_excel_file():
    """Kullanıcıya dosya seçtiren basit file dialog."""
    root = tk.Tk()
    root.withdraw()  # pencereyi gizle
    root.attributes('-topmost', True)
    root.update()

    file_path = filedialog.askopenfilename(
        title="Unpivot yapılacak Excel dosyasını seçin",
        filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsm")]
    )

    root.destroy()
    return file_path


def write_unpivot_to_new_sheet(wb, df_unpivot):
    """
    Aynı workbook içinde yeni bir sheet açar,
    Unpivot sonucunu A1'den itibaren bu sayfaya yazar.
    """
    # Önce var olan sheet isimlerini toplayalım
    sheet_names = [ws.Name for ws in wb.Worksheets]

    base_name = "Unpivot"
    new_name = base_name
    i = 1
    while new_name in sheet_names:
        i += 1
        new_name = f"{base_name}{i}"

    # Yeni sheet oluştur (en sona ekle)
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = new_name

    # DataFrame'i 2D listeye çevir (başlık + satırlar)
    data = [df_unpivot.columns.tolist()] + df_unpivot.values.tolist()
    n_rows = len(data)
    n_cols = len(data[0]) if n_rows > 0 else 0

    if n_rows == 0 or n_cols == 0:
        return ws  # yazacak bir şey yoksa boş sayfa bırakırız

    # A1'den başlayarak tabloyu yaz
    ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols)).Value = data

    return ws


def main():
    print("=== UNPIVOT ARACI ===")

    file_path = select_excel_file()
    if not file_path:
        print("Dosya seçilmedi, işlem iptal edildi.")
        return

    print(f"Seçilen dosya: {file_path}")

    # Excel'i aç
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True  # Kullanıcı Excel'i görsün

    try:
        wb = excel.Workbooks.Open(file_path)
    except Exception as e:
        print("Excel dosyası açılamadı:", e)
        excel.Quit()
        return

    print("\nExcel açıldı.")
    print("Lütfen unpivot yapılacak alanı Excel içinde fare ile seçin.")
    print("Seçimi yaptıktan sonra buraya dönüp Enter'a basın...")
    input()  # Kullanıcı Enter'a basana kadar bekle

    selection = excel.Selection

    # Seçimdeki hücreleri al (2D tuple)
    values = selection.Value

    if values is None:
        print("Herhangi bir alan seçilmedi veya seçim boş.")
        # Burada Excel'i açık bırakalım, kullanıcı kendi kapatır
        return

    # Tek hücre seçildiyse values tek tuple olabilir, onu 2D forma sokalım
    if not isinstance(values, (tuple, list)) or not isinstance(values[0], (tuple, list)):
        values = ((values,),)

    # İlk satır başlık, kalanlar veri
    header = values[0]
    data_rows = values[1:]

    # DataFrame oluştur
    df = pd.DataFrame(data_rows, columns=header)

    print("\nSeçilen aralık DataFrame'e alındı.")
    print("Kolonlar:")
    for i, col in enumerate(df.columns):
        print(f"{i}: {col}")

    # Kullanıcıdan soldan kaç kolonun sabit (id_vars) olacağını soralım
    while True:
        try:
            n_id = int(input("\nSoldan kaç kolon sabit kalsın? (Departman, Personel vb.): "))
            if 0 < n_id < len(df.columns):
                break
            else:
                print(f"Lütfen 1 ile {len(df.columns) - 1} arasında bir sayı girin.")
        except ValueError:
            print("Lütfen sayı girin.")

    id_vars = list(df.columns[:n_id])
    value_vars = list(df.columns[n_id:])

    print("\nSabit kalacak kolonlar (id_vars):", id_vars)
    print("Unpivot yapılacak kolonlar (value_vars):", value_vars)

    # Unpivot (melt)
    df_unpivot = df.melt(
        id_vars=id_vars,
        value_vars=value_vars,
        var_name="Tarih",
        value_name="Deger"
    )

    print("\nUnpivot işlemi tamamlandı. İlk birkaç satır:")
    print(df_unpivot.head())

    # Sonucu aynı workbook içinde yeni bir sheet'e yaz
    ws_unpivot = write_unpivot_to_new_sheet(wb, df_unpivot)

    print(f"\nSonuç, {wb.Name} dosyasında '{ws_unpivot.Name}' sayfasına yazıldı.")
    print("Excel şu anda açık. İsterseniz Ctrl+S ile dosyayı kaydedebilirsiniz.")

    # Excel'i kapatmıyoruz; kullanıcı isterse inceleyip kaydeder/kapatır.
    # wb.Save() otomatik kaydetmek istersen buraya eklenebilir.


if __name__ == "__main__":
    main()
