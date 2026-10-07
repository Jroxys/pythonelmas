# pythonelmas

## Maaş Karşılaştırma Programı (`maas_karsilastir.py`)

İki Excel dosyasındaki maaşları **isme göre** karşılaştırır.

- **Dosya 1:** Sayfalara bölünmüş (Anaokulu, İlkokul, Ortaokul, Lise ...). Tüm sayfalar birleştirilir.
- **Dosya 2:** Tek liste.

### Kurulum (bir kez)
1. [Python](https://www.python.org/downloads/) kurun (kurarken "Add Python to PATH" kutusunu işaretleyin).
2. `pip install openpyxl`

### Kullanım
- **Windows:** `maas_karsilastir.bat` dosyasına çift tıklayın, açılan pencerelerden iki Excel dosyasını seçin.
- **Komut satırı:** `python maas_karsilastir.py dosya1.xlsx dosya2.xlsx -o sonuc.xlsx`

Seçenekler: `--tolerans 1` (1 TL farkı aynı say), `--ad1 B --maas1 F --ad2 A --maas2 C` (sütunları elle belirt).

### Özellikler
- Başlık satırını ve "Ad Soyad" / "Maaş" sütunlarını otomatik bulur (birden fazla maaş sütunu varsa, örn. Brüt/Net, bir kez sorar).
- Büyük/küçük harf, Türkçe karakter (İ/ı, ş, ğ...) ve fazla boşluk farklarını yok sayar.
- `1.234,56` gibi metin olarak yazılmış maaşları da okur; TOPLAM satırlarını atlar.
- Sonuç Excel dosyasında sayfalar: Özet, Tüm Karşılaştırma, Farklı Maaşlar, Sadece Dosya 1'de, Sadece Dosya 2'de.

### .exe yapmak (Python olmayan bilgisayarlar için)
Windows'ta `exe_olustur.bat` dosyasına çift tıklayın. `dist\MaasKarsilastir.exe` oluşur; bu tek dosyayı istediğiniz kişiye verebilirsiniz, Python kurması gerekmez.
