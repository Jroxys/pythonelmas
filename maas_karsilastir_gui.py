#!/usr/bin/env python3
"""Maaş Karşılaştırma - pencereli (görsel) arayüz.

Çalıştırma:  python maas_karsilastir_gui.py
"""
import os
import sys
import tkinter as tk
from collections import defaultdict
from tkinter import filedialog, messagebox, ttk

import maas_karsilastir as mk

OKUL_ADI = "Sivas Cumhuriyet Üniversitesi Vakıf Okulları"


def kaynak_yolu(ad):
    """Program klasöründeki (veya .exe içine gömülü) dosyanın yolu."""
    taban = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(taban, ad)


class Uygulama(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Maaş Karşılaştırma")
        self.geometry("720x620")
        self.minsize(640, 560)
        self.configure(bg="#f3f5f9")
        self.dosya1 = tk.StringVar()
        self.dosya2 = tk.StringVar()
        self.elden = tk.BooleanVar(value=False)
        self.tolerans = tk.StringVar(value="0")
        self.son_sonuc = None
        mk.SECICI = self.sutun_sor
        self.logo = self._logo_yukle()
        self._stil()
        self._arayuz()

    def _logo_yukle(self):
        """logo.png varsa pencere simgesi ve başlık görseli olarak kullanır."""
        yol = kaynak_yolu("logo.png")
        if not os.path.isfile(yol):
            return None
        try:
            ham = tk.PhotoImage(file=yol)
            self.iconphoto(True, ham)
            kucult = max(1, -(-ham.height() // 80))  # başlıkta en çok ~80 piksel yükseklik
            return ham.subsample(kucult, kucult)
        except tk.TclError:
            return None

    def _stil(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("TFrame", background="#f3f5f9")
        st.configure("Kart.TLabelframe", background="#ffffff")
        st.configure("Kart.TLabelframe.Label", background="#f3f5f9", foreground="#1f4e78",
                     font=("Segoe UI", 10, "bold"))
        st.configure("TLabel", background="#ffffff", font=("Segoe UI", 10))
        st.configure("TCheckbutton", background="#ffffff", font=("Segoe UI", 10))
        st.configure("Sec.TButton", font=("Segoe UI", 10), padding=6)
        st.configure("Ana.TButton", font=("Segoe UI", 13, "bold"), padding=12,
                     background="#1f4e78", foreground="#ffffff")
        st.map("Ana.TButton", background=[("active", "#2a6aa3"), ("disabled", "#9db3c8")])

    def _arayuz(self):
        if self.logo:
            tk.Label(self, image=self.logo, bg="#f3f5f9").pack(pady=(12, 0))
        ttk.Label(self, text=OKUL_ADI, font=("Segoe UI", 10),
                  background="#f3f5f9", foreground="#555").pack(pady=(12 if not self.logo else 4, 0))
        ttk.Label(self, text="Maaş Karşılaştırma", font=("Segoe UI", 18, "bold"),
                  background="#f3f5f9", foreground="#1f4e78").pack()
        ttk.Label(self, text="İki Excel dosyasındaki maaşları TC no / isme göre karşılaştırır",
                  background="#f3f5f9", foreground="#555").pack(pady=(0, 10))

        kart = ttk.Frame(self)
        kart.pack(fill="x", padx=18)
        self._dosya_satiri(kart, "1) Sayfalı Excel (anaokulu, ilkokul, lise ...)", self.dosya1)
        self._dosya_satiri(kart, "2) Tek listeli Excel", self.dosya2)

        ayar = ttk.LabelFrame(self, text=" Ayarlar ", style="Kart.TLabelframe", padding=10)
        ayar.pack(fill="x", padx=18, pady=(8, 0))
        ttk.Checkbutton(ayar, text="2. dosyada MAAŞ + ELDEN toplamını karşılaştır",
                        variable=self.elden).grid(row=0, column=0, sticky="w", columnspan=3)
        ttk.Label(ayar, text="Tolerans (TL) - bu tutara kadar farkı 'aynı' say:").grid(
            row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Spinbox(ayar, from_=0, to=100000, width=8, textvariable=self.tolerans).grid(
            row=1, column=1, sticky="w", padx=8, pady=(8, 0))


        self.btn = ttk.Button(self, text="MAAŞLARI KARŞILAŞTIR", style="Ana.TButton", command=self.calistir)
        self.btn.pack(fill="x", padx=18, pady=12)

        self.bar = ttk.Progressbar(self, mode="indeterminate")
        self.bar.pack(fill="x", padx=18)

        self.log = tk.Text(self, height=9, state="disabled", bg="#ffffff", relief="flat",
                           font=("Consolas", 9), padx=8, pady=6)
        self.log.pack(fill="both", expand=True, padx=18, pady=8)
        self.log.tag_configure("hata", foreground="#c00000")
        self.log.tag_configure("ok", foreground="#2e7d32")

        alt = ttk.Frame(self)
        alt.pack(fill="x", padx=18, pady=(0, 12))
        self.btn_ac = ttk.Button(alt, text="Sonuç dosyasını aç", style="Sec.TButton",
                                 command=self.sonucu_ac, state="disabled")
        self.btn_ac.pack(side="left")
        self.yaz("Hazır. İki dosyayı seçip 'Maaşları Karşılaştır' butonuna basın.")

    def _dosya_satiri(self, ust, baslik, degisken):
        f = ttk.LabelFrame(ust, text=" " + baslik + " ", style="Kart.TLabelframe", padding=8)
        f.pack(fill="x", pady=4)
        ttk.Entry(f, textvariable=degisken, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(f, text="Dosya Seç...", style="Sec.TButton",
                   command=lambda: self.dosya_sec(degisken)).pack(side="left", padx=(8, 0))

    # ------------------------------------------------------------ yardımcılar
    def yaz(self, metin, etiket=None):
        self.log.configure(state="normal")
        self.log.insert("end", metin + "\n", etiket)
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def dosya_sec(self, degisken):
        yol = filedialog.askopenfilename(title="Excel dosyası seçin",
                                         filetypes=[("Excel dosyaları", "*.xlsx *.xlsm")])
        if yol:
            degisken.set(yol)

    def sutun_sor(self, mesaj, secenekler):
        """Birden fazla aday sütun varsa (örn. Brüt/Net maaş) seçim penceresi açar."""
        pen = tk.Toplevel(self)
        pen.title("Sütun seçimi")
        pen.transient(self)
        pen.grab_set()
        pen.resizable(False, False)
        ttk.Label(pen, text=mesaj, font=("Segoe UI", 10, "bold"), background="#f3f5f9").pack(
            padx=16, pady=(14, 6), anchor="w")
        secili = tk.IntVar(value=0)
        for n, (_, etiket) in enumerate(secenekler):
            ttk.Radiobutton(pen, text=etiket, variable=secili, value=n).pack(padx=24, anchor="w")
        ttk.Button(pen, text="Tamam", style="Sec.TButton", command=pen.destroy).pack(pady=14)
        pen.configure(bg="#f3f5f9")
        self.wait_window(pen)
        return secenekler[secili.get()][0]

    # ------------------------------------------------------------ ana işlem
    def calistir(self):
        d1, d2 = self.dosya1.get(), self.dosya2.get()
        if not d1 or not d2:
            messagebox.showwarning("Eksik dosya", "Lütfen iki Excel dosyasını da seçin.")
            return
        try:
            tol = float(self.tolerans.get().replace(",", "."))
        except ValueError:
            messagebox.showwarning("Hatalı tolerans", "Tolerans bir sayı olmalı (örn. 0 veya 1).")
            return
        cikti = filedialog.asksaveasfilename(title="Sonuç dosyasını kaydet", defaultextension=".xlsx",
                                             initialfile="maas_karsilastirma.xlsx",
                                             filetypes=[("Excel", "*.xlsx")])
        if not cikti:
            return
        self.btn.configure(state="disabled")
        self.btn_ac.configure(state="disabled")
        self.bar.start(12)
        self.yaz("\nKarşılaştırılıyor...")
        try:
            self._isle(d1, d2, cikti, tol)
        except PermissionError:
            self.yaz("HATA: Sonuç dosyası açık olabilir. Excel'de kapatıp tekrar deneyin.", "hata")
            messagebox.showerror("Hata", "Sonuç dosyası yazılamadı. Açıksa kapatıp tekrar deneyin.")
        except Exception as e:  # kullanıcıya anlaşılır göster
            self.yaz(f"HATA: {e}", "hata")
            messagebox.showerror("Hata", str(e))
        finally:
            self.bar.stop()
            self.btn.configure(state="normal")

    def _isle(self, d1, d2, cikti, tol):
        for d in (d1, d2):
            if d.lower().endswith(".xls"):
                raise ValueError(f"'{os.path.basename(d)}' eski .xls biçiminde. Excel'de 'Farklı Kaydet' "
                                 "ile .xlsx yapıp tekrar deneyin.")
        k1, u1 = mk.kisileri_oku(d1, etiket="Dosya 1")
        k2, u2 = mk.kisileri_oku(d2, etiket="Dosya 2")
        for u in u1 + u2:
            self.yaz("Uyarı: " + u)
        if not k1 or not k2:
            raise ValueError("Dosyalardan kayıt okunamadı. 'ADI SOYADI' ve 'MAAŞ' başlıklarını kontrol edin.")
        for etiket, k in (("Dosya 1", k1), ("Dosya 2", k2)):
            sayac = defaultdict(int)
            for x in k:
                sayac[x["tc"] or x["anahtar"]] += 1
            tekrar = sum(1 for n in sayac.values() if n > 1)
            if tekrar:
                self.yaz(f"Uyarı: {etiket}'de aynı kişi birden fazla kez geçiyor ({tekrar} kayıt), sırayla eşleştirildi.")
        eslesen, s1, s2 = mk.karsilastir(k1, k2, tol, self.elden.get())
        tekrar = mk.tekrar_tc_satirlari([("Dosya 1", k1), ("Dosya 2", k2)])
        if tekrar:
            self.yaz(f"UYARI: {len({(d, k['tc']) for d, k in tekrar})} TC numarası birden fazla farklı kişide geçiyor! "
                     "Sonuçta 'Tekrarlayan TC' sayfasına bakın; kaynak dosyada TC yanlış olabilir.", "hata")
        mk.rapor_yaz(cikti, d1, d2, eslesen, s1, s2, tol, self.elden.get(), tekrar)
        farkli = sum(1 for e in eslesen if not e["ayni"])
        self.yaz(f"Dosya 1: {len(k1)} kişi ({len({k['sayfa'] for k in k1})} sayfa)  |  Dosya 2: {len(k2)} kişi")
        self.yaz(f"Eşleşen: {len(eslesen)}  |  Maaşı farklı: {farkli}  |  "
                 f"Sadece Dosya 1'de: {len(s1)}  |  Sadece Dosya 2'de: {len(s2)}")
        kontrol = sum(e["kontrol"] for e in eslesen)
        if kontrol:
            self.yaz(f"DİKKAT: {kontrol} kişide TC aynı ama isim farklı! Sonuçta 'İsim Kontrol' sayfasına bakın.", "hata")
        self.yaz("Sonuç kaydedildi: " + cikti, "ok")
        self.son_sonuc = cikti
        self.btn_ac.configure(state="normal")

    def sonucu_ac(self):
        if not self.son_sonuc:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.son_sonuc)
            elif sys.platform == "darwin":
                os.system(f'open "{self.son_sonuc}"')
            else:
                os.system(f'xdg-open "{self.son_sonuc}"')
        except Exception as e:
            messagebox.showerror("Hata", str(e))


if __name__ == "__main__":
    Uygulama().mainloop()
