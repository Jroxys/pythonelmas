"""Telefon ve e-posta bilgilerini bir Excel'den diğerine TCKN + ad + soyad eşleşmesiyle aktarır.

Kullanım:
    python excel_birlestir.py                      (dosya seçme ekranları açılır)
    python excel_birlestir.py KAYNAK HEDEF [-o CIKTI.xlsx]   (.xlsx veya eski .xls)

KAYNAK : telefon ve e-posta dolu olan dosya
HEDEF  : ad, soyad, TC ve doğum tarihi dolu olan dosya (formatı korunarak doldurulur)
"""
import argparse
import re
import sys

from openpyxl import load_workbook

# Türkçe harfler ASCII karşılıklarıyla aynı sayılır (İ/ı/i/I->I, Ş->S, Ç->C, Ğ->G, Ö->O, Ü->U)
HARF_ESLE = str.maketrans("iİıIşŞçÇğĞöÖüÜ", "IIIISSCCGGOOUU")


def ac(yol, **kw):
    """Excel'i açar. Eski .xls ise xlrd ile okuyup bellekte .xlsx çalışma kitabına çevirir."""
    if not yol.lower().endswith(".xls"):
        return load_workbook(yol, **kw)
    try:
        import xlrd
    except ImportError:
        sys.exit("Eski .xls dosyası için önce şunu çalıştırın: pip install xlrd")
    from openpyxl import Workbook

    kitap = xlrd.open_workbook(yol)
    sayfa = kitap.sheet_by_index(0)
    wb = Workbook()
    ws = wb.active
    for r in range(sayfa.nrows):
        for c in range(sayfa.ncols):
            hucre = sayfa.cell(r, c)
            deger = hucre.value
            if hucre.ctype == xlrd.XL_CELL_DATE:
                d = xlrd.xldate_as_datetime(deger, kitap.datemode)
                deger = d.strftime("%d.%m.%Y")
            elif hucre.ctype == xlrd.XL_CELL_NUMBER and float(deger).is_integer():
                deger = int(deger)
            elif hucre.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                continue
            ws.cell(r + 1, c + 1, deger)
    return wb


def norm_ad(deger):
    """Boşlukları sadeleştirir, büyük harfe çevirir; Türkçe harfler ASCII karşılığıyla aynı sayılır."""
    if deger is None:
        return ""
    s = " ".join(str(deger).split())
    return s.translate(HARF_ESLE).upper()


def norm_tc(deger):
    if deger is None:
        return ""
    s = re.sub(r"\D", "", str(deger).split(".")[0] if isinstance(deger, float) else str(deger))
    return s.zfill(11) if s else ""


def norm_tel(deger):
    """Sadece rakam bırakır, başındaki 0'ı atar (örn. 530XXXXXX7)."""
    if deger is None:
        return None
    s = re.sub(r"\D", "", str(deger).split(".")[0] if isinstance(deger, float) else str(deger))
    if s.startswith("90") and len(s) == 12:
        s = s[2:]
    return s.lstrip("0") or None


def baslik_bul(ws):
    """Başlık satırını ve sütun numaralarını bulur (başlıklar uzun açıklamalı olduğundan öneke bakılır)."""
    anahtarlar = {
        "tc": "TCKN",
        "ad": "ÇALIŞAN ADI",
        "soyad": "ÇALIŞAN SOYADI",
        "tel": "TELEFONU",
        "eposta": "E-POSTA",
    }
    for satir in ws.iter_rows(min_row=1, max_row=15):
        sutunlar = {}
        for hucre in satir:
            if isinstance(hucre.value, str):
                metin = norm_ad(hucre.value)
                for ad, onek in anahtarlar.items():
                    if ad not in sutunlar and metin.startswith(norm_ad(onek)):
                        sutunlar[ad] = hucre.column
        if "tc" in sutunlar and "ad" in sutunlar:
            return satir[0].row, sutunlar
    sys.exit(f"Başlık satırı bulunamadı: {ws.title}")


def dosya_sec():
    """Sırayla kaynak ve hedef Excel'i seçtirir, çıktı için kayıt yeri sorar."""
    import tkinter as tk
    from tkinter import filedialog

    pencere = tk.Tk()
    pencere.withdraw()
    pencere.attributes("-topmost", True)
    tipler = [("Excel dosyaları", "*.xlsx *.xlsm *.xls"), ("Tüm dosyalar", "*.*")]
    kaynak = filedialog.askopenfilename(title="1) Telefon/e-posta DOLU Excel'i seçin", filetypes=tipler)
    if not kaynak:
        sys.exit("Kaynak dosya seçilmedi.")
    hedef = filedialog.askopenfilename(title="2) Bilgilerin GİRİLECEĞİ Excel'i seçin", filetypes=tipler)
    if not hedef:
        sys.exit("Hedef dosya seçilmedi.")
    cikti = filedialog.asksaveasfilename(
        title="Sonuç nereye kaydedilsin?", defaultextension=".xlsx",
        initialfile="birlesik.xlsx", filetypes=[("Excel", "*.xlsx")])
    if not cikti:
        sys.exit("Kayıt yeri seçilmedi.")
    return kaynak, hedef, cikti, pencere


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("kaynak", nargs="?", help="Telefon/e-posta dolu Excel")
    p.add_argument("hedef", nargs="?", help="Ad, TC, doğum tarihi dolu Excel")
    p.add_argument("-o", "--cikti", help="Çıktı dosyası (varsayılan: hedef dosyanın yanına *_dolu.xlsx)")
    a = p.parse_args()

    pencere = None
    if not (a.kaynak and a.hedef):
        a.kaynak, a.hedef, a.cikti, pencere = dosya_sec()

    kws = ac(a.kaynak, data_only=True).active
    k_satir, k_sut = baslik_bul(kws)

    # (tc, ad, soyad) -> (telefon, eposta)
    kaynak = {}
    for r in range(k_satir + 1, kws.max_row + 1):
        tc = norm_tc(kws.cell(r, k_sut["tc"]).value)
        if not tc:
            continue
        ad = norm_ad(kws.cell(r, k_sut["ad"]).value)
        soyad = norm_ad(kws.cell(r, k_sut["soyad"]).value) if "soyad" in k_sut else ""
        tel = norm_tel(kws.cell(r, k_sut["tel"]).value) if "tel" in k_sut else None
        eposta = kws.cell(r, k_sut["eposta"]).value if "eposta" in k_sut else None
        eposta = str(eposta).strip() if eposta not in (None, "") else None
        kaynak[(tc, ad, soyad)] = (tel, eposta)

    hwb = ac(a.hedef)
    hws = hwb.active
    h_satir, h_sut = baslik_bul(hws)
    for gerekli in ("tel", "eposta"):
        if gerekli not in h_sut:
            sys.exit(f"Hedef dosyada '{gerekli}' sütunu yok.")

    eslesen = eslesmeyen = 0
    eslesmeyenler = []
    for r in range(h_satir + 1, hws.max_row + 1):
        tc = norm_tc(hws.cell(r, h_sut["tc"]).value)
        if not tc:
            continue
        ad = norm_ad(hws.cell(r, h_sut["ad"]).value)
        soyad = norm_ad(hws.cell(r, h_sut["soyad"]).value) if "soyad" in h_sut else ""
        bilgi = kaynak.get((tc, ad, soyad))
        if bilgi is None:
            eslesmeyen += 1
            eslesmeyenler.append((r, tc, ad, soyad))
            continue
        tel, eposta = bilgi
        if tel:
            hws.cell(r, h_sut["tel"]).value = tel
        if eposta:
            hws.cell(r, h_sut["eposta"]).value = eposta
        eslesen += 1

    cikti = a.cikti or re.sub(r"\.xls[xm]?$", "", a.hedef, flags=re.I) + "_dolu.xlsx"
    hwb.save(cikti)
    mesaj = f"Eşleşen: {eslesen} | Eşleşmeyen: {eslesmeyen} | Çıktı: {cikti}"
    ayrinti = "".join(
        f"\n  Satır {r}: {tc} {ad} {soyad} -> kaynakta bulunamadı" for r, tc, ad, soyad in eslesmeyenler)
    print(mesaj + ayrinti)
    if pencere:
        from tkinter import messagebox
        messagebox.showinfo("Tamamlandı", mesaj + ayrinti[:1500])


if __name__ == "__main__":
    main()
