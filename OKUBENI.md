# TJK Python Sistemi — Kullanım

Koşu gününün programını çek → `yeni yer` → `yapılacak yer` üret.

## GÜN MANTIĞI (önemli)
Dört bileşenin DÖRDÜ DE artık **mutlak tarih** hedefliyor: hedef = `gerçek bugün + ofset`.
Varsayılan **ofset = 0 → bugün**. Site hangi günü gösterirse göstersin (TJK bazen dünü
gösterir) program bu mutlak tarihe götürülür. "+1 ileri tık" zamanlama riski KALKTI.

Kural çok basit: **işlemek istediğin koşu gününün sabahı dört komutu da çalıştır.**
- Bugünün koşularını işliyorsan: hiçbir ofset değiştirme (hepsi 0).
- Bir günü ÖNCEDEN hazırlamak istersen: dört dosyada da ofseti aynı yap
  (derece/800: `PROGRAM_GUN_OFFSET`, pipeline: `GUN_OFFSET`; galop `now()` kullanır,
  başka gün için galop'taki `program_tarihi()`'yi de ayarla). En kolayı: o gün çalıştır.

## Çalıştırma sırası (3 scraper + pipeline)
```
cd ~/Desktop/tjk
python3 ERTESI_GUN_MODU_OKA_BASAN_PLUS_FAZ30_STIL_TEK_PY.py     # derece + stil
python3 son800_BUGUN_ILK_AT_STIL_FINAL_TEST.py                  # 800
python3 tjk_tum_akis_main_SELF_HEAL_KOSU_LINK_FIX_YARIN_TAM.py  # galop/son galop -> tjk_ULTRA_MESAFELI_YARIN.xlsx
python3 tjk_yeni_yer.py                                          # -> yeni_yer_SONUC.xlsx
```
İlk 3 scraper veri dosyalarını üretir, son komut hepsini birleştirir. Çok il varsa
(ör. Ankara + Kocaeli) ana program ikisini de tek sayfaya, doğru per-il
orjin/dede/galop ile işler. Mac uyumasın diye komutların başına `caffeinate -i` koy.

## Çalıştırma
```
cd ~/Desktop/tjk
python3 tjk_yeni_yer.py
```
Çıktı: `yeni_yer_SONUC.xlsx` (sayfalar: Sayfa1, Sayfa2, yeni yer, yapılacak yer).

## Klasörde bulunması gereken dosyalar

**Ana program + modüller (9 .py — hepsi aynı klasörde):**
- `tjk_yeni_yer.py`  ← çalıştırılan ana dosya
- `tjk_donustur.py`  (Sayfa1 → Sayfa2)
- `yeni_yer_hesapla.py`  (formül katmanı → yeni yer)
- `derece_donustur.py`  (derece verisi)
- `son800_donustur.py`  (800 verisi)
- `yapilacak_yer.py`  (2. aşama: ana tablo + orjin/dede/galop/son galop yerleştirme)
- `orjin_panel.py`  (orjin/dede paneli üretici — AnalizCalistir karşılığı)
- `baba_dede.py`  (Sayfa1 → Sayfa2 baba/dede ayıklama)
- `sayfa6_hesapla.py`  (Sayfa6 VLOOKUP katmanı: baba/dede → ÇEKİLECEK YER)

**Scraper'lar (veri çeker):**
- `ERTESI_GUN_MODU_OKA_BASAN_PLUS_FAZ30_STIL_TEK_PY.py`  (derece + video/stil)
- `son800_BUGUN_ILK_AT_STIL_FINAL_TEST.py`  (800)

**Referans/girdi dosyaları:**
- `buuuuuuuuuuuuuuuuuuuuuuu.xlsm`  (800/derece/medyan referans DB)
- `tjk_ULTRA_MESAFELI.xlsx`  (galop=xy, son galop=xx — galop üreticinin çıktısı)
- `baba.xlsm`  (orjin/dede paneli için Sayfa5 model + Sayfa6 veri)

## Kurulum (bir kez)
```
pip3 install selenium webdriver-manager pandas openpyxl beautifulsoup4 lxml
```

## Ayarlar (`tjk_yeni_yer.py` en üstte)

Gün hizalama (ÜÇÜ DE AYNI olmalı — scraper'larla):
```
GUN_OFFSET = 1          # 0=bugün, 1=yarın, 2=2 gün sonra  (ŞU AN: yarın)
```

Veri kaynakları:
```
DERECE_KAYNAGI   = "veri"   # "ref" | "veri" | "calistir"
SEKIZYUZ_KAYNAGI = "veri"
```

2. aşama (yapılacak yer) girdileri:
```
YY_GALOP    = ("tjk_ULTRA_MESAFELI_YARIN.xlsx", "xy")
YY_SONGALOP = ("tjk_ULTRA_MESAFELI_YARIN.xlsx", "xx")

# orjin/dede panelini SIFIRDAN üret (önerilen — baba.xlsm formülüne gerek yok):
# Scraped Sayfa1 + referans DB'lerden tüm zinciri Python kurar.
# Referans workbook 'ÇEKİLECEK YER' + 'kısa orta uzun' + Sayfa5 içermeli.
YY_ORJIN_FULL = ("baba.xlsm", "İzmir")   # (referans_dosya, il) — baba modu
YY_DEDE_FULL  = ("baba.xlsm", "İzmir")   # dede modu

# Alternatifler (FULL=None ise):
#   YY_ORJIN_GEN = ("baba.xlsm","İzmir")  -> hazır Sayfa6'dan panel üret
#   YY_ORJIN = ("girdiler.xlsx","orjin")  -> hazır panel sayfasını oku
```
İl'i (İzmir/Bursa/…) çektiğin güne göre buradan ver. Öncelik: FULL > GEN > hazır sayfa.

## Akış (özet)
1. Scraper'ları çalıştır (derece + 800) → veri dosyalarını üretir.
2. galop üreticini çalıştır → `tjk_ULTRA_MESAFELI.xlsx`.
3. orjin/dede için `baba.xlsm`'i güncelle (Sayfa1 yapıştır → Sayfa6 formülleri hesaplar).
4. `python3 tjk_yeni_yer.py` → `yeni_yer_SONUC.xlsx`.

## Doğrulama durumu (hepsi mevcut dosyalarınla SIFIR fark)
- yeni yer formül katmanı ✅
- yapılacak yer (ana tablo + orjin + dede + galop + son galop) ✅ (10.563 hücre)
- orjin/dede panel üretici (AnalizCalistir) ✅ (2.327 hücre)
- Sayfa6 VLOOKUP katmanı ✅ (61 kolon × 68 satır)
- Baba/Dede ayıklama (Sayfa1→Sayfa2) ✅
- TAM orjin/dede zinciri (Sayfa1→Sayfa2→Sayfa6→panel) ✅ (2.313 hücre)

## Tamamen otomatik olan
- yeni yer, yapılacak yer, orjin, dede — hepsi Python (baba.xlsm formülüne gerek yok).
- Sen sadece referans DB'leri güncel tut: 800/derece/medyan, ÇEKİLECEK YER, kısa orta uzun, Sayfa5.
- galop/son galop: kendi py üreticin (tjk_ULTRA_MESAFELI.xlsx).

## Notlar
- Renk/merge gibi görsel biçim henüz uygulanmadı (sen "önce doğru veri+düzen" dedin); değerler
  ve yerleşim birebir. İstenirse renk/merge eklenir.
- Bir girdi dosyası yoksa o blok atlanır, program yine çalışır.
