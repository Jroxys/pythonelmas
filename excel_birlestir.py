"""Telefon ve e-posta bilgilerini bir Excel'den diğerine aktarır.

Kullanım:
    python excel_birlestir.py                      (dosya seçme ekranları açılır)
    python excel_birlestir.py KAYNAK HEDEF [-o CIKTI]   (.xlsx veya eski .xls)

KAYNAK : telefon ve e-posta dolu olan dosya
HEDEF  : ad, soyad, TC ve doğum tarihi dolu olan dosya
Çıktı, HEDEF dosyanın birebir kopyasıdır (aynı biçim, aynı uzantı); yalnızca
telefon ve e-posta hücreleri doldurulur.

Eşleştirme: önce TC + ad + soyad; bulunamazsa TC ile (isim farkı raporlanır).
Gerekli paketler: openpyxl; .xls için ayrıca xlrd, xlwt ve xlutils.
"""
import argparse
import re
import sys

from openpyxl import load_workbook

# Türkçe harfler ASCII karşılıklarıyla aynı sayılır (İ/ı/i/I->I, Ş->S, Ç->C, Ğ->G, Ö->O, Ü->U)
HARF_ESLE = str.maketrans("iİıIşŞçÇğĞöÖüÜ", "IIIISSCCGGOOUU")


def norm_ad(deger):
    """Boşlukları sadeleştirir, büyük harfe çevirir; Türkçe harfler ASCII karşılığıyla aynı sayılır."""
    if deger is None:
        return ""
    s = " ".join(str(deger).split())
    return s.translate(HARF_ESLE).upper()


def sayi_metni(deger):
    if isinstance(deger, float) and deger.is_integer():
        return str(int(deger))
    return str(deger)


def norm_tc(deger):
    if deger is None:
        return ""
    s = re.sub(r"\D", "", sayi_metni(deger))
    return s.zfill(11) if s else ""


def norm_tel(deger):
    """Sadece rakam bırakır, başındaki 0'ı atar (örn. 530XXXXXX7)."""
    if deger is None:
        return None
    s = re.sub(r"\D", "", sayi_metni(deger))
    if s.startswith("90") and len(s) == 12:
        s = s[2:]
    return s.lstrip("0") or None


def xls_mi(yol):
    return yol.lower().endswith(".xls")


def satirlari_oku(yol):
    """Dosyanın ilk sayfasını değer listeleri olarak döndürür (0 tabanlı)."""
    if xls_mi(yol):
        xlrd = xls_paketleri()[0]
        sayfa = xlrd.open_workbook(yol).sheet_by_index(0)
        return [[sayfa.cell_value(r, c) if sayfa.cell_type(r, c) not in (0, 6) else None
                 for c in range(sayfa.ncols)] for r in range(sayfa.nrows)]
    ws = load_workbook(yol, data_only=True).active
    return [list(r) for r in ws.iter_rows(values_only=True)]


def xls_paketleri():
    try:
        import xlrd
        import xlutils.copy
    except ImportError:
        sys.exit("Eski .xls dosyası için önce şunu çalıştırın: pip install xlrd xlwt xlutils")
    return xlrd, xlutils.copy


def baslik_bul(satirlar, ad_):
    """Başlık satırını ve sütun numaralarını bulur (başlıklar uzun açıklamalı olduğundan öneke bakılır)."""
    anahtarlar = {
        "tc": "TCKN",
        "ad": "ÇALIŞAN ADI",
        "soyad": "ÇALIŞAN SOYADI",
        "tel": "TELEFONU",
        "eposta": "E-POSTA",
    }
    for i, satir in enumerate(satirlar[:15]):
        sutunlar = {}
        for j, deger in enumerate(satir):
            if isinstance(deger, str):
                metin = norm_ad(deger)
                for ad, onek in anahtarlar.items():
                    if ad not in sutunlar and metin.startswith(norm_ad(onek)):
                        sutunlar[ad] = j
        if "tc" in sutunlar and "ad" in sutunlar:
            return i, sutunlar
    sys.exit(f"Başlık satırı bulunamadı: {ad_}")


class Hedef:
    """Hedef dosyayı biçimini bozmadan düzenler ve kaydeder (.xlsx veya .xls)."""

    def __init__(self, yol):
        self.xls = xls_mi(yol)
        if self.xls:
            xlrd, xlutils_copy = xls_paketleri()
            okunan = xlrd.open_workbook(yol, formatting_info=True)
            self.wb = xlutils_copy.copy(okunan)
            self.ws = self.wb.get_sheet(0)
            self.ws._cell_overwrite_ok = True
        else:
            self.wb = load_workbook(yol)
            self.ws = self.wb.active

    def yaz(self, r, c, deger):
        if not self.xls:
            self.ws.cell(r + 1, c + 1).value = deger
            return
        satir = self.ws._Worksheet__rows.get(r)
        eski = satir._Row__cells.get(c) if satir else None
        xf = getattr(eski, "xf_idx", None)
        self.ws.write(r, c, deger)
        if xf is not None:
            self.ws._Worksheet__rows[r]._Row__cells[c].xf_idx = xf  # hücre biçimini koru

    def kaydet(self, yol):
        self.wb.save(yol)


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
    uzanti = ".xls" if xls_mi(hedef) else ".xlsx"
    cikti = filedialog.asksaveasfilename(
        title="Sonuç nereye kaydedilsin?", defaultextension=uzanti,
        initialfile="birlesik" + uzanti, filetypes=[("Excel", "*" + uzanti)])
    if not cikti:
        sys.exit("Kayıt yeri seçilmedi.")
    return kaynak, hedef, cikti, pencere


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("kaynak", nargs="?", help="Telefon/e-posta dolu Excel")
    p.add_argument("hedef", nargs="?", help="Ad, TC, doğum tarihi dolu Excel")
    p.add_argument("-o", "--cikti", help="Çıktı dosyası (varsayılan: hedef dosyanın yanına *_dolu)")
    a = p.parse_args()

    pencere = None
    if not (a.kaynak and a.hedef):
        a.kaynak, a.hedef, a.cikti, pencere = dosya_sec()

    # Kaynak: tam anahtar (tc, ad, soyad) ve yalnızca TC ile arama tabloları
    satirlar = satirlari_oku(a.kaynak)
    k_satir, k_sut = baslik_bul(satirlar, a.kaynak)

    def hucre(satir, anahtar):
        j = k_sut.get(anahtar)
        return satir[j] if j is not None and j < len(satir) else None

    tam, tc_ile = {}, {}
    for satir in satirlar[k_satir + 1:]:
        tc = norm_tc(hucre(satir, "tc"))
        if not tc:
            continue
        ad, soyad = norm_ad(hucre(satir, "ad")), norm_ad(hucre(satir, "soyad"))
        eposta = hucre(satir, "eposta")
        bilgi = (norm_tel(hucre(satir, "tel")), str(eposta).strip() if eposta not in (None, "") else None)
        tam[(tc, ad, soyad)] = bilgi
        tc_ile[tc] = (ad, soyad, bilgi)

    hedef = Hedef(a.hedef)
    h_satirlar = satirlari_oku(a.hedef)
    h_satir, h_sut = baslik_bul(h_satirlar, a.hedef)
    for gerekli in ("tel", "eposta"):
        if gerekli not in h_sut:
            sys.exit(f"Hedef dosyada '{gerekli}' sütunu yok.")

    def hh(satir, anahtar):
        j = h_sut.get(anahtar)
        return satir[j] if j is not None and j < len(satir) else None

    tam_eslesen = 0
    tc_eslesen, bulunamayan = [], []
    for r in range(h_satir + 1, len(h_satirlar)):
        satir = h_satirlar[r]
        tc = norm_tc(hh(satir, "tc"))
        if not tc:
            continue
        ad, soyad = norm_ad(hh(satir, "ad")), norm_ad(hh(satir, "soyad"))
        bilgi = tam.get((tc, ad, soyad))
        if bilgi is not None:
            tam_eslesen += 1
        elif tc in tc_ile:
            k_ad, k_soyad, bilgi = tc_ile[tc]
            tc_eslesen.append(f"Satır {r + 1}: {tc} hedefte '{ad} {soyad}' / kaynakta '{k_ad} {k_soyad}'")
        else:
            bulunamayan.append(f"Satır {r + 1}: {tc} {ad} {soyad}")
            continue
        tel, eposta = bilgi
        if tel:
            hedef.yaz(r, h_sut["tel"], tel)
        if eposta:
            hedef.yaz(r, h_sut["eposta"], eposta)

    cikti = a.cikti or re.sub(r"(\.xls[xm]?)$", r"_dolu\1", a.hedef, flags=re.I)
    try:
        hedef.kaydet(cikti)
    except PermissionError:
        # Dosya genelde Excel'de açıktır; kapatmadan devam edebilmek için yeni bir adla kaydet
        import time
        yeni = re.sub(r"(\.xls[xm]?)$", time.strftime("_%H%M%S") + r"\1", cikti, flags=re.I)
        print(f"'{cikti}' yazılamadı (Excel'de açık olabilir). Şuraya kaydediyorum: {yeni}")
        cikti = yeni
        hedef.kaydet(cikti)

    satirlar_ = [f"Tam eşleşen: {tam_eslesen} | Sadece TC ile eşleşen: {len(tc_eslesen)} "
                 f"| Bulunamayan: {len(bulunamayan)} | Çıktı: {cikti}"]
    if tc_eslesen:
        satirlar_.append("\nTC aynı ama isim farklı (doldurdum, kontrol edin):")
        satirlar_ += ["  " + s for s in tc_eslesen]
    if bulunamayan:
        satirlar_.append("\nKaynakta TC'si hiç yok (boş bıraktım):")
        satirlar_ += ["  " + s for s in bulunamayan]
    metin = "\n".join(satirlar_)
    print(metin)
    if pencere:
        from tkinter import messagebox
        messagebox.showinfo("Tamamlandı", metin[:1800])


if __name__ == "__main__":
    main()
