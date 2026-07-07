# -*- coding: utf-8 -*-
print("ERTESI_GUN_MODU + FAZ30_STIL TEK PY AKTIF")
print("1) Yarının da ertesi günü programındaki atlar çekilecek")
print("2) Bu atların son 6 ay ilk-4 geçmiş satırları Excel olacak")
print("3) Aynı Excel üstüne %45 koşu stili eklenecek")

import re
print("TUM_ILLER_TUM_ATLAR_EKSIKSIZ_ILK4_ONCEKI_HP_ERTESI_GUN.py AKTIF")
import time
import pandas as pd
import requests
from io import StringIO
from math import ceil, floor

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import WebDriverException, NoSuchWindowException
from webdriver_manager.chrome import ChromeDriverManager

from urllib.parse import urlparse, parse_qs, urlencode, urljoin
from datetime import datetime, timedelta


# =========================================
# HTTP SESSION - HIZ / STABILITE
# =========================================

SESSION = requests.Session()



# =========================================
# TARİH
# =========================================

bugun = datetime.today()
# MUTLAK GÜN: hedef = gerçek bugün + PROGRAM_GUN_OFFSET (pipeline/galop ile aynı).
# Değeri TEK YERDEN ayarla: gun_ayar.py içindeki OFFSET. (0=bugün,1=yarın,2=...)
try:
    from gun_ayar import OFFSET as PROGRAM_GUN_OFFSET
except Exception:
    PROGRAM_GUN_OFFSET = 0
try:
    from gun_ayar import hedef_gun as _hedef_gun
    program_tarihi = _hedef_gun()
except Exception:
    program_tarihi = bugun + timedelta(days=PROGRAM_GUN_OFFSET)
alti_ay_once = program_tarihi - timedelta(days=183)


# =========================================
# AYARLAR
# =========================================

ANA_URL = "https://www.tjk.org/TR/yarissever/Info/Page/GunlukYarisProgrami"

TURK_SEHIRLER = [
    "İstanbul",
    "Bursa",
    "İzmir",
    "Kocaeli",
    "Ankara",
    "Adana",
    "Şanlıurfa",
    "Diyarbakır",
    "Elazığ",
    "Antalya"
]

EXCEL_NAME = "TUM_ILLER_TUM_ATLAR_EKSIKSIZ_ILK4_SON6AY_ERTESI_GUN.xlsx"

# Chrome arkada çalışsın. Görmek istersen False yap.
HEADLESS_CHROME = True

# Selenium fallback açıldığında sayfanın render edilmesi için bekleme.
# Çok düşürmek veri kaçırabilir; 3 güvenli hız ayarıdır.
SELENIUM_BEKLEME_SN = 1

# Medyan/HP-Kilo için fazla saplanma.
# 3 denemede gelmezse satır Veri'ye yine yazılır, hata Basarisiz_Atlar'a düşer.
MAX_MEDYAN_DENEME = 3

# =========================================
# TEST MODU - TEK KOŞULUK HIZLI DENEME
# =========================================

TEST_MODU = False

# Günlük programdaki ilk şehir alınır.
TEST_SEHIR_LIMIT = 999
TEST_SEHIR_ADI = ""

# İlk koşuyu hızlı test etmek için günlük programdaki ilk at linkleri alınır.
# Çoğu koşuda 8-12 at olduğu için 12 güvenli test sınırıdır.
TEST_AT_LIMIT = 9999
# İlk atlar bugünkü/henüz koşmamış yarışlara takılıyor; test için ileriden başla.
TEST_AT_BASLANGIC = 0

# Dolu HP yarış bulunca fazla uzatmadan dur.
TEST_BASARILI_SATIR_LIMIT = 8

# Testte sonsuza gitmesin; uygun sonuç bulana kadar günlük illerde toplam bu kadar at dener.
TEST_TOPLAM_AT_DENEME_LIMIT = 60
test_toplam_at_deneme_sayaci = 0


# =========================================
# TEMİZLİK
# =========================================

def siklet_temizle(x):
    x = str(x).strip()
    x = re.sub(r"[^0-9,.]", "", x)

    if x.isdigit() and len(x) == 3:
        return x[:2] + "," + x[2]

    return x.replace(".", ",")


def pist_durumu_temizle(x):
    x = str(x).strip()
    x = re.sub(r"\s+\d+([.,]\d+)?$", "", x)
    return x.strip()


def derece_temizle(value):
    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value == "" or value.lower() == "nan":
        return ""

    value = value.replace(".", "")

    if not value.isdigit():
        return ""

    number = int(value)

    if 6000 <= number <= 19999:
        number = number - 4000
    elif 20000 <= number <= 29999:
        number = number - 8000
    elif 30000 <= number <= 39999:
        number = number - 12000

    return number



def sira_ilk4_mi(value):
    """Bitiriş sırasını sağlam yakalar.

    TJK geçmiş tablosunda S kolonu bazen 1, 1., '1 ', '1/12' gibi okunabilir.
    Son 6 ayda ilk 4'e giren satırı kaçırmak kabul değil; bu yüzden baştaki
    sayıyı alıp 1-4 aralığını kontrol ediyoruz.
    """
    if pd.isna(value):
        return False
    t = str(value).strip()
    m = re.search(r"(\d+)", t)
    if not m:
        return False
    try:
        return int(m.group(1)) in [1, 2, 3, 4]
    except Exception:
        return False



def onceki_kosu_hp_getir(df, row_index):
    """İşlenen ilk-4 satırın HP'sini değil, hemen altındaki önceki koşunun HP'sini döndürür.

    TJK at geçmiş tablosunda genelde en yeni koşu üstte, daha eski koşular alttadır.
    Bu yüzden işlenen satırın bir altındaki satır, tarih olarak bir önceki koşudur.

    Kural:
    - İşlenen satırın kendi HP'si asla kullanılmaz.
    - Bir alt satır varsa onun HP'si alınır; boşsa boş yazılır.
    - Bir alt satır yoksa bu atın hayatının ilk koşusudur; HP = 0 yazılır.
    """
    if df is None or len(df) == 0:
        return "0"

    hp_col = hp_gecmis_kolonu_bul(df)
    if hp_col is None:
        return ""

    try:
        pozisyon = list(df.index).index(row_index)
    except ValueError:
        return ""

    onceki_pozisyon = pozisyon + 1
    if onceki_pozisyon >= len(df):
        return "0"

    onceki_row = df.iloc[onceki_pozisyon]
    hp_raw = onceki_row.get(hp_col, "")

    hp_formatli = hp_deger_formatla(hp_raw)
    if hp_formatli != "":
        return hp_formatli

    # Önceki koşunun HP hücresi gerçekten boş/geçersizse aynen boş bırak.
    t = str(hp_raw).strip()
    if t.lower() in ["nan", "none", "<na>"]:
        return ""
    return t


def onceki_kosu_hp_debug_getir(df, row_index):
    """Kontrol için önceki HP'nin hangi satırdan geldiğini açıklar."""
    if df is None or len(df) == 0:
        return "DF boş; HP=0"
    try:
        pozisyon = list(df.index).index(row_index)
    except ValueError:
        return "Satır index bulunamadı"
    onceki_pozisyon = pozisyon + 1
    if onceki_pozisyon >= len(df):
        return "Hayatının ilk koşusu; altta önceki koşu yok; HP=0"
    onceki_row = df.iloc[onceki_pozisyon]
    tarih = satir_kolon_getir(onceki_row, "Tarih")
    sira = satir_kolon_getir(onceki_row, "S")
    hp_col = hp_gecmis_kolonu_bul(df)
    hp_raw = onceki_row.get(hp_col, "") if hp_col is not None else ""
    return f"HP bir alt/önceki koşudan alındı | Önceki Tarih={tarih} | Önceki Sıra={sira} | Önceki HP Ham={hp_raw}"

def satir_excel_kaydi_olustur(row, at_adi, hp_final, hp_kaynak_notu, ilk_3_hp_serisi, ilk_3_hp_kcinsleri, ilk_3_hp_ham):
    """İlk-4 satırı için temel Excel kaydını oluşturur.

    Bu kayıt medyan/link bulunmasa bile yazılır. Eksik at çekme mantığı kesinlikle yok.
    """
    pist = str(satir_kolon_getir(row, "Pist"))
    pist_turu = pist.split(":")[0].strip()
    pist_durumu = ""

    if ":" in pist:
        pist_durumu = pist_durumu_temizle(pist.split(":", 1)[1].strip())

    if pist_turu == "S":
        pist_turu = "Sentetik"
    elif pist_turu == "K":
        pist_turu = "Kum"
    elif pist_turu == "Ç":
        pist_turu = "Çim"

    grup = str(satir_kolon_getir(row, "Grup")).upper()
    irk = ""
    if grup.endswith("A"):
        irk = "Arap"
    elif grup.endswith("İ") or grup.endswith("I"):
        irk = "İngiliz"

    return {
        "Sıra": satir_kolon_getir(row, "S"),
        "At Adı": at_adi,
        "Sıklet": siklet_temizle(satir_kolon_getir(row, "Sıklet")),
        "Irk": irk,
        "Cinsiyet": "",
        "Tarih": satir_kolon_getir(row, "Tarih"),
        "Şehir": satir_kolon_getir(row, "Şehir"),
        "Pist": pist_turu,
        "Pist Durumu": pist_durumu,
        "Mesafe": satir_kolon_getir(row, "Msf"),
        "Derece": derece_temizle(satir_kolon_getir(row, "Derece")),
        "Koşu Cinsi": satir_kolon_getir(row, "Kcins"),
        "HP": hp_final,
        "HP Kaynak Notu": hp_kaynak_notu,
        "Atın İlk 3 HP": ilk_3_hp_serisi,
        "Atın İlk 3 HP Koşu Cinsleri": ilk_3_hp_kcinsleri,
        "İlk 3 HP Ham": ilk_3_hp_ham,
        "Koşu HP Medyan": "",
        "Koşu HP Dolu Atlar Kilo Medyan": "",
        "Medyan HP Listesi": "",
        "Medyan Kilo Listesi": "",
        "HP-Kilo Eşleşme Listesi": "",
        "Koşu Detay URL": "",
        "Kaynak URL": "",
        "Satır Durumu": "MEDYAN_BEKLIYOR",
        "Hata": "",
    }


def satir_kolon_getir(row, kolon_adi):
    if kolon_adi in row.index:
        return row[kolon_adi]

    hedef = str(kolon_adi).strip().upper()

    for c in row.index:
        if str(c).strip().upper() == hedef:
            return row[c]

    return ""


def fetch_html_url(url):
    """TJK sayfasını User-Agent ile ham HTML olarak alır."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/125.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
        "Referer": ANA_URL,
    }
    r = SESSION.get(url, headers=headers, timeout=30)
    r.raise_for_status()

    html = r.text
    if len(html.strip()) < 200:
        raise Exception("Sayfa boş veya eksik geldi")

    return html


def read_html_url(url):
    """TJK sayfasını pandas yerine önce requests ile alır.
    User-Agent ekler; bazı atlarda çıplak pd.read_html 404/engelleme verebiliyor.
    """
    html = fetch_html_url(url)
    return pd.read_html(StringIO(html))


def at_url_adaylari(at_id, orijinal_href=""):
    """Aynı at için farklı TJK URL formatlarını dener.
    404 alan atlarda tek URL'ye takılmamak için.
    """
    adaylar = []

    if orijinal_href:
        adaylar.append(orijinal_href)

    base1 = "https://www.tjk.org/TR/YarisSever/Query/ConnectedPage/AtKosuBilgileri"
    base2 = "https://www.tjk.org/TR/yarissever/Query/ConnectedPage/AtKosuBilgileri"

    param_setleri = [
        {"1": "1", "QueryParameter_AtId": at_id, "Era": "today"},
        {"QueryParameter_AtId": at_id, "Era": "today"},
        {"1": "1", "QueryParameter_AtId": at_id},
        {"QueryParameter_AtId": at_id},
    ]

    for base in [base1, base2]:
        for params in param_setleri:
            adaylar.append(base + "?" + urlencode(params))

    # tekrarları sırayı bozmadan temizle
    temiz = []
    for u in adaylar:
        if u and u not in temiz:
            temiz.append(u)

    return temiz


def read_html_retry(url, at_adi="", max_deneme=10, bekleme=5):
    """Eski genel fonksiyon. Geriye uyumluluk için bırakıldı."""
    son_hata = None

    for deneme in range(1, max_deneme + 1):
        try:
            return read_html_url(url)
        except Exception as e:
            son_hata = e
            print(f"TEKRAR DENENECEK ({deneme}/{max_deneme}) - AT: {at_adi} - HATA: {e}")
            time.sleep(bekleme)

    raise Exception(f"{at_adi} için {max_deneme} denemeden sonra veri alınamadı. Son hata: {son_hata}")




def html_tablolari_oku(html):
    """HTML içinde tablo varsa döndürür; yoksa kontrollü exception verir."""
    if html is None or len(str(html).strip()) < 200:
        raise Exception("HTML boş veya eksik")
    tables = pd.read_html(StringIO(html))
    if tables is None or len(tables) == 0:
        raise Exception("HTML içinde tablo yok")
    return tables


def read_html_detay_retry(detay_url, max_deneme=2, bekleme=2):
    """Koşu detay sayfasını boş geçmeden okumaya zorlar.

    Requests bazen TJK sonuç tablosunu getirmiyor. Bu durumda Selenium ile gerçek
    render edilmiş sayfa okunur. İlk Selenium denemesinde tablo yoksa refresh +
    ek bekleme yapılır. Burada amaç hız değil doğru medyan; boş medyan yazmak tırt.
    """
    global driver
    son_hata = None

    # 1) Hızlı requests denemeleri
    for deneme in range(1, max_deneme + 1):
        try:
            html = fetch_html_url(detay_url)
            return html_tablolari_oku(html)
        except Exception as e:
            son_hata = e
            time.sleep(bekleme)

    # 2) Zorunlu Selenium fallback: render bekle + gerekirse refresh
    for deneme in range(1, 4):
        try:
            driver = guvenli_get(driver, detay_url, max_deneme=2, bekleme=bekleme)
            time.sleep(2 + deneme)

            # Sayfanın aşağı/yukarı hareketi bazı TJK tablolarını tetikliyor.
            try:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1)
                driver.execute_script("window.scrollTo(0, 0);")
                time.sleep(1)
            except Exception:
                pass

            html = driver.page_source
            return html_tablolari_oku(html)

        except Exception as e:
            son_hata = e
            try:
                driver.refresh()
                time.sleep(3)
                html = driver.page_source
                return html_tablolari_oku(html)
            except Exception as e2:
                son_hata = e2
                time.sleep(bekleme)

    raise Exception(f"Koşu detay tablosu okunamadı. Son hata: {son_hata}")

def read_html_at_retry(at_id, at_adi="", orijinal_href="", max_deneme=3, bekleme=3):
    """At özelinde sağlam okuma.
    Her URL'yi birkaç kez değil, önce alternatif URL'leri dener.
    404 gelirse aynı URL'yi 10 kere denemek tırt; direkt sonraki adaya geçer.
    """
    son_hata = None
    denenenler = []

    for aday_url in at_url_adaylari(at_id, orijinal_href):
        denenenler.append(aday_url)

        for deneme in range(1, max_deneme + 1):
            try:
                html = fetch_html_url(aday_url)
                tables = pd.read_html(StringIO(html))
                print(f"OKUNDU: {at_adi} -> {aday_url}")
                return tables, aday_url, html

            except requests.exceptions.HTTPError as e:
                son_hata = e
                status = getattr(e.response, "status_code", "")
                print(f"URL HATASI ({status}) - AT: {at_adi} - URL: {aday_url}")

                # 404 kalıcıdır; aynı yanlış URL'yi tekrar deneme.
                if status == 404:
                    break

                time.sleep(bekleme)

            except Exception as e:
                son_hata = e
                print(f"TEKRAR DENENECEK ({deneme}/{max_deneme}) - AT: {at_adi} - URL: {aday_url} - HATA: {e}")
                time.sleep(bekleme)

    raise Exception(
        f"{at_adi} için veri alınamadı. Son hata: {son_hata}. "
        f"Denenen URL sayısı: {len(denenenler)}"
    )


# =========================================
# KOŞU İÇİ HP DOLU ATLARIN HP / KİLO MEDYAN REFERANSI
# =========================================

detay_referans_cache = {}        # URL bazlı cache
kosu_kimlik_cache = {}          # Aynı koşu kimliği bazlı güvenli cache
kosu_kimlik_cache_url = {}      # Kontrol için: cache anahtarının ilk görülen URL'si

# --- KALICI SQLite KÜTÜPHANE (opsiyonel hızlandırıcı) ---
# Koşu kimliği bazlı; geçmiş koşuları tekrar TJK'dan çekmemek için. Çökse bile
# ana scraper eskisi gibi (TJK'dan çekerek) çalışır. False yaparsan kapanır.
KUTUPHANE_KULLAN = True
try:
    import tjk_kutuphane as _KUT
    _kut_conn = _KUT.baglan() if KUTUPHANE_KULLAN else None
    if _kut_conn is not None:
        print(f"[KÜTÜPHANE] açık — kayıtlı koşu sayısı: {_KUT.sayac(_kut_conn)}")
except Exception as _kut_e:
    print(f"[KÜTÜPHANE] yüklenemedi (görmezden geliniyor): {_kut_e}")
    _KUT = None
    _kut_conn = None
detay_cache_sayac = {
    "toplam_ihtiyac": 0,
    "kimlik_cache_isabet": 0,
    "url_cache_isabet": 0,
    "gercek_detay_acilis": 0,
    "kimlik_cache_pas_gecildi": 0,
    "kilo_sutun_suspect_pas": 0,
}


def metin_norm(x):
    t = str(x).strip().upper()
    t = re.sub(r"\s+", " ", t)
    return t


def kosu_cache_anahtari(row):
    """Aynı koşuyu URL'den değil, koşunun ortak yarış bilgilerinden yakalar.

    ÖNEMLİ:
    Derece bu anahtardan özellikle çıkarıldı.
    Çünkü derece atın kendi bitiriş derecesidir; aynı koşudaki her atta farklı olur.
    Derece anahtarda kalırsa aynı yarışı koşan atlar farklı yarış sanılır ve kimlik cache 0 vurur.

    Güvenli anahtar:
    Tarih + Şehir + Mesafe + Pist + Koşu Cinsi

    Bu birleşim aynı gün/şehir içinde aynı yarışı ayırt etmek için yeterince sıkıdır.
    Şüpheli/eksik bilgi varsa cache zorlanmaz, normal URL okuma yolu çalışır.
    """
    tarih = metin_norm(satir_kolon_getir(row, "Tarih"))
    sehir = metin_norm(satir_kolon_getir(row, "Şehir"))
    msf = metin_norm(satir_kolon_getir(row, "Msf"))
    pist = metin_norm(satir_kolon_getir(row, "Pist"))
    kcins = metin_norm(satir_kolon_getir(row, "Kcins"))

    zorunlu = [tarih, sehir, msf, pist, kcins]
    if any(x == "" or x == "NAN" for x in zorunlu):
        return ""

    return "|".join(zorunlu)


def sayi_float_temizle(x):
    if pd.isna(x):
        return None

    t = str(x).strip()
    if t == "" or t.lower() == "nan":
        return None

    t = re.sub(r"[^0-9,.-]", "", t)
    if t == "":
        return None

    # 1.234,5 gibi olursa binlik noktayı at, virgülü ondalık yap.
    if "," in t:
        t = t.replace(".", "").replace(",", ".")

    try:
        return float(t)
    except:
        return None




def siklet_float_temizle(x):
    """Yarış sonucu Sıklet değerini temizler.

    KRİTİK: TJK sonuç tablosunda 54+0,13 gibi değerlerde + sonrası
    bindirme/ek bilgidir; kilo medyanına girmeyecek.
    Örnek: 54+0,13 -> 54, 55+1.20 -> 55, 58,5 -> 58.5
    """
    if pd.isna(x):
        return None

    t = str(x).strip()
    if t == "" or t.lower() == "nan":
        return None

    if "+" in t:
        t = t.split("+", 1)[0].strip()

    t = re.sub(r"[^0-9,.-]", "", t)
    if t == "":
        return None

    if "," in t:
        t = t.replace(".", "").replace(",", ".")

    try:
        return float(t)
    except:
        return None


def liste_deger_formatla(x):
    try:
        v = float(x)
    except:
        return ""
    if v.is_integer():
        return str(int(v))
    return str(v).replace(".", ",")

def hp_deger_formatla(x):
    """HP değerini Excel çıktısı için temiz formatlar."""
    v = sayi_float_temizle(x)
    if v is None:
        return ""

    if float(v).is_integer():
        return str(int(v))

    return str(v).replace(".", ",")


def hp_gecmis_kolonu_bul(df):
    """At geçmiş tablosundaki HP kolonunu bulur.

    TJK bazen kolon adını HP gibi gösterse de pandas okurken boşluk/nokta/çoklu başlık
    karışabiliyor. Görünen başlık HP olduğu için normalize edilmiş hali HP olan kolon kabul edilir.
    """
    if df is None:
        return None

    for c in df.columns:
        cn = kolon_adi_norm(c)
        cn_sade = re.sub(r"[^A-Z0-9]", "", cn)
        if cn == "HP" or cn_sade == "HP":
            return c

    return None


def tarih_kolonu_bul(df):
    for c in df.columns:
        cn = kolon_adi_norm(c)
        if cn == "TARIH" or cn.endswith(" TARIH"):
            return c
    return "Tarih" if "Tarih" in df.columns else None


def gercek_gecmis_kosu_satirlari(df):
    """At geçmiş tablosunda sadece gerçek koşu satırlarını bırakır.

    Bu sürüm tablo sırasını bozmaz. TJK geçmiş tablosunda en yeni üstte,
    en eski alttadır. Bu yüzden ilk 3 HP için tarih sıralaması yapmayacağız;
    gerçek koşu satırlarını ayıklayıp en alttaki 3 satırı tersten okuyacağız.
    """
    if df is None or len(df) == 0:
        return pd.DataFrame()

    tarih_col = tarih_kolonu_bul(df)
    if tarih_col is None or tarih_col not in df.columns:
        return pd.DataFrame()

    tmp = df.copy()
    tmp["__ORJ_SIRA__"] = range(len(tmp))
    tmp["__TARIH_PARSED__"] = pd.to_datetime(tmp[tarih_col], dayfirst=True, errors="coerce")
    tmp = tmp.dropna(subset=["__TARIH_PARSED__"])

    msf_col = None
    pist_col = None
    s_col = None
    for c in tmp.columns:
        cn = kolon_adi_norm(c)
        if cn in ["MSF", "MESAFE"]:
            msf_col = c
        if cn == "PIST":
            pist_col = c
        if cn == "S":
            s_col = c

    if msf_col is not None:
        tmp = tmp[tmp[msf_col].astype(str).str.extract(r"(\d{3,4})", expand=False).notna()]

    if pist_col is not None:
        tmp = tmp[tmp[pist_col].astype(str).str.strip().ne("")]

    # Bitiriş sırası olan S kolonunda sayı yoksa gerçek yarış satırı değildir.
    # Derecesiz/koşmadı satırları kırpmasın diye sadece kolon sağlamsa uygula.
    if s_col is not None:
        s_txt = tmp[s_col].astype(str).str.strip()
        if s_txt.str.extract(r"^(\d+)$", expand=False).notna().sum() >= 3:
            tmp = tmp[s_txt.str.extract(r"^(\d+)$", expand=False).notna()]

    return tmp.sort_values("__ORJ_SIRA__", ascending=True, kind="mergesort")


def at_gecmisi_tamamla(df, at_sayfa_html, at_adi="", max_sayfa=6):
    """At geçmiş listesi sunucuda KESİKSE tamamlar.

    TJK at sayfası uzun listelerde ilk ~50 koşuyu verir; altta
    'Daha Fazla Sonuç Göster' formu vardır ('Toplam 62 sonuçtan 50 tanesi
    gösteriliyor'). Bu form aynı adrese PageNumber'lı GET atar. Burada o form
    çözülür, devam sayfaları requests ile çekilir ve satırlar df'e eklenir ->
    EN ESKİ koşular da gelir, İLK 3 HP gerçek değerleriyle hesaplanır.
    HERHANGİ bir hata olursa mevcut df aynen döner (50-güvenlik kuralı yanlış
    ilk-HP'yi zaten engelliyor; bu fonksiyon asla scrape'i bozmaz)."""
    try:
        html = at_sayfa_html
        eklenen = 0
        for _sayfa in range(max_sayfa):
            form = None
            for fm in re.finditer(r"<form\b.*?</form>", html, re.S | re.I):
                if "show-more" in fm.group(0):
                    form = fm.group(0)
                    break
            if form is None:
                break
            act = re.search(r'action="([^"]*)"', form)
            action = (act.group(1) if act else "").replace("&amp;", "&")
            if not action:
                break
            inputs = {}
            for im in re.finditer(r"<input\b[^>]*>", form, re.I):
                tag = im.group(0)
                n = re.search(r'name="([^"]*)"', tag)
                v = re.search(r'value="([^"]*)"', tag)
                if n:
                    inputs[n.group(1)] = (v.group(1) if v else "")
            if action.startswith("/"):
                action = "https://www.tjk.org" + action
            devam_url = action + ("&" if "?" in action else "?") + urlencode(inputs)
            html = fetch_html_url(devam_url)
            time.sleep(1)
            # devam sayfasındaki geçmiş tablosunu bul (df ile en çok ortak kolonlu)
            ek = None
            try:
                for t in pd.read_html(StringIO(html)):
                    ortak = len(set(map(str, t.columns)) & set(map(str, df.columns)))
                    if ek is None or ortak > ek[0] or (ortak == ek[0] and len(t) > len(ek[1])):
                        ek = (ortak, t)
            except Exception:
                ek = None
            if ek is None or ek[0] == 0:
                # fragman <table> etiketi olmadan gelmiş olabilir -> sarmala
                try:
                    tl = pd.read_html(StringIO("<table>" + html + "</table>"))
                    if tl and len(tl[0].columns) == len(df.columns):
                        t0 = tl[0]
                        t0.columns = df.columns
                        ek = (len(df.columns), t0)
                except Exception:
                    pass
            if ek is None or len(ek[1]) == 0:
                break
            df = pd.concat([df, ek[1]], ignore_index=True, sort=False)
            eklenen += len(ek[1])
        if eklenen:
            print(f"      [TAM LİSTE] {at_adi}: 'daha fazla' ile +{eklenen} eski koşu eklendi (toplam {len(df)})")
    except Exception as _e:
        print(f"      [TAM LİSTE] {at_adi}: devam sayfası alınamadı ({str(_e)[:80]}) — mevcut listeyle devam")
    return df


def atin_ilk_3_hp_bilgisi(df):
    """At geçmişinin en altından yukarı doğru ilk 3 DOLU HP'yi alır.

    Kural:
    - TJK geçmiş tablosunda en eski koşular alttadır.
    - En alttan yukarı taranır.
    - HP boş/NAN/geçersizse o satır SAYILMAZ, atlanır.
    - İlk 3 dolu HP bulununca durulur.
    - Aynı 3 HP'nin tarihleri ve koşu cinsleri ayrıca döndürülür.
    - 3 dolu HP yoksa eksik kalanlar baştan 0 ile tamamlanır.
    """
    hp_col = hp_gecmis_kolonu_bul(df)
    tarih_col = tarih_kolonu_bul(df)
    kosular = gercek_gecmis_kosu_satirlari(df)

    if len(kosular) == 0 or hp_col is None:
        return "0-0-0/", "", "", ""

    hp_listesi = []
    tarih_listesi = []
    ham_hp_listesi = []
    kcins_listesi = []

    # En alttan yukarı: önce en eski dolu HP, sonra daha yeni dolu HP'ler.
    for _, r in kosular.iloc[::-1].iterrows():
        hp_raw = r.get(hp_col, "")
        hp_num = sayi_float_temizle(hp_raw)

        # Boş HP artık başlangıç kabul edilmiyor; satır direkt atlanır.
        if hp_num is None or not (0 <= float(hp_num) <= 180):
            continue

        hp_listesi.append(hp_deger_formatla(hp_raw))
        ham_hp_listesi.append(str(hp_raw))

        if tarih_col is not None and tarih_col in r.index:
            tarih_listesi.append(str(r.get(tarih_col, "")))
        else:
            tarih_listesi.append("")

        kcins_listesi.append(str(satir_kolon_getir(r, "Kcins")).strip())

        if len(hp_listesi) >= 3:
            break

    # GÜVENLİK KURALI: gerçek İLK HP en fazla 50 olur. En eski dolu HP 50'nin
    # ÜZERİNDEYSE liste kesiktir (sayfa 'daha fazla göster' ile kısalmış, atın
    # gerçek ilk koşuları HTML'de yok) -> YANLIŞ değer yazma, TAMAMEN BOŞ bırak.
    if hp_listesi:
        _ilk_hp = sayi_float_temizle(hp_listesi[0])
        if _ilk_hp is not None and float(_ilk_hp) > 50:
            return "", "", "", ""

    # 3 dolu HP yoksa eksik kalanlar baştan 0 ile tamamlanır.
    while len(hp_listesi) < 3:
        hp_listesi.insert(0, "0")
        tarih_listesi.insert(0, "")
        ham_hp_listesi.insert(0, "")
        kcins_listesi.insert(0, "")

    return (
        "-".join(hp_listesi[:3]) + "/",
        " | ".join(tarih_listesi[:3]),
        " | ".join(ham_hp_listesi[:3]),
        "-".join(kcins_listesi[:3])
    )


def atin_ilk_3_hp_serisi(df):
    return atin_ilk_3_hp_bilgisi(df)[0]

def kolon_bul(df, adaylar):
    for c in df.columns:
        c_norm = str(c).strip().upper()
        for a in adaylar:
            if a in c_norm:
                return c
    return None


def kolon_adi_norm(c):
    """Pandas/MultiIndex kolon adını tek satır ve normalize metne çevirir."""
    if isinstance(c, tuple):
        c = " ".join([str(x) for x in c if str(x).strip() and str(x).strip().lower() != "nan"])
    t = str(c).strip().upper()
    t = t.replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ü", "U").replace("Ö", "O").replace("Ç", "C")
    t = re.sub(r"\s+", " ", t)
    return t





def metin_norm_at(x):
    t = str(x).strip().upper()
    t = t.replace("İ", "I").replace("İ", "I")
    t = t.replace("Ş", "S").replace("Ğ", "G").replace("Ü", "U").replace("Ö", "O").replace("Ç", "C")
    t = re.sub(r"[^A-Z0-9]+", "", t)
    return t


def html_anchor_yakin_parca(html, detay_url, pencere=90000):
    """#224544 gibi anchor varsa HTML'de o koşu bloğunun yakınını alır.
    Amaç bütün günün bütün koşularını değil, ilgili koşuyu okumak.
    """
    if not html:
        return html

    m = re.search(r"#(\d+)$", str(detay_url))
    if not m:
        return html

    anchor = m.group(1)
    idx = str(html).find(anchor)
    if idx < 0:
        return html

    bas = max(0, idx - pencere // 3)
    son = min(len(html), idx + pencere)
    return str(html)[bas:son]


def detay_html_oku(detay_url, max_deneme=2, bekleme=2):
    """Önce requests, olmazsa Selenium ile HTML döndürür."""
    global driver
    son_hata = None

    for _ in range(max_deneme):
        try:
            return fetch_html_url(detay_url)
        except Exception as e:
            son_hata = e
            time.sleep(bekleme)

    for deneme in range(1, 4):
        try:
            driver = guvenli_get(driver, detay_url, max_deneme=2, bekleme=bekleme)
            time.sleep(2 + deneme)
            try:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1)
                driver.execute_script("window.scrollTo(0, 0);")
                time.sleep(1)
            except Exception:
                pass
            return driver.page_source
        except Exception as e:
            son_hata = e
            time.sleep(bekleme)

    raise Exception(f"Detay HTML okunamadı: {son_hata}")


def hp_kilo_kolonlarini_bul(df):
    """Sadece başlığı TAM HP ve TAM SIKLET olan kolonları kabul eder."""
    hp_col = None
    kilo_col = None

    for c in df.columns:
        cn = kolon_adi_norm(c)
        cn_sade = re.sub(r"[^A-Z0-9]", "", cn)

        if cn_sade == "HP":
            hp_col = c
        if cn_sade == "SIKLET":
            kilo_col = c

    if hp_col is None or kilo_col is None:
        return None

    tmp = df[[hp_col, kilo_col]].copy()
    tmp.columns = ["HP", "Sıklet"]
    tmp["HP_NUM"] = tmp["HP"].apply(sayi_float_temizle)
    tmp["KILO_NUM"] = tmp["Sıklet"].apply(siklet_float_temizle)

    # HP boş olan at medyana girmez. Kilo aynı HP'si dolu atlardan alınır.
    tmp = tmp.dropna(subset=["HP_NUM", "KILO_NUM"])
    tmp = tmp[
        (tmp["HP_NUM"] >= 1) & (tmp["HP_NUM"] <= 180) &
        (tmp["KILO_NUM"] >= 35) & (tmp["KILO_NUM"] <= 75)
    ].copy()

    if len(tmp) < 2:
        return None

    return tmp


def kosu_ici_hp_medyani_referans_hesapla(detay_url, at_adi=""):
    """Başlığı HP olan sütundan HP medyan, başlığı Sıklet olan sütundan kilo medyan.

    Kural basit:
    - Bütün gün sayfası değil, varsa #anchor yakınındaki koşu bloğu okunur.
    - Sadece TAM HP ve TAM SIKLET kolonları.
    - Medyan, HP listesinin min-max dışına çıkarsa direkt hata.
    - Excel'e kontrol için kullanılan HP/Kilo listesi global değişkenle yazılır.
    """
    global son_medyan_hp_liste, son_medyan_kilo_liste, son_medyan_eslesme_liste
    global son_medyan_hp_liste_ham, son_medyan_kilo_liste_ham
    son_medyan_hp_liste = ""
    son_medyan_kilo_liste = ""
    son_medyan_eslesme_liste = ""
    son_medyan_hp_liste_ham = []
    son_medyan_kilo_liste_ham = []

    if not detay_url:
        return "", ""

    if detay_url in detay_referans_cache:
        detay_cache_sayac["url_cache_isabet"] += 1
        return detay_referans_cache[detay_url]

    detay_cache_sayac["gercek_detay_acilis"] += 1

    html = detay_html_oku(detay_url, max_deneme=2, bekleme=2)
    parca = html_anchor_yakin_parca(html, detay_url)

    try:
        tables = pd.read_html(StringIO(parca))
    except Exception:
        tables = pd.read_html(StringIO(html))

    adaylar = []
    for tablo_no, df in enumerate(tables, start=1):
        tmp = hp_kilo_kolonlarini_bul(df)
        if tmp is None:
            detay_cache_sayac["kilo_sutun_suspect_pas"] += 1
            continue

        hp_liste = [float(x) for x in tmp["HP_NUM"].tolist()]
        kilo_liste = [float(x) for x in tmp["KILO_NUM"].tolist()]
        hp_ref = float(pd.Series(hp_liste).median())
        kilo_ref = float(pd.Series(kilo_liste).median())

        # Medyan veri aralığının dışına çıkamaz. Çıkarsa o tablo çöp.
        if hp_ref < min(hp_liste) or hp_ref > max(hp_liste):
            detay_cache_sayac["kilo_sutun_suspect_pas"] += 1
            continue

        adaylar.append({
            "tablo_no": tablo_no,
            "n": len(tmp),
            "hp_ref": hp_ref,
            "kilo_ref": round(kilo_ref, 1),
            "hp_liste": hp_liste,
            "kilo_liste": kilo_liste,
        })

    if not adaylar:
        raise Exception(f"HP/Sıklet medyan tablosu bulunamadı: {at_adi} | {detay_url}")

    # İlgili koşu bloğunda gerçek yarış tablosu en fazla HP+Sıklet satırı olan tablodur.
    secilen = sorted(adaylar, key=lambda x: x["n"], reverse=True)[0]

    son_medyan_hp_liste = "-".join([liste_deger_formatla(x) for x in secilen["hp_liste"]])
    son_medyan_kilo_liste = "-".join([liste_deger_formatla(x) for x in secilen["kilo_liste"]])
    son_medyan_eslesme_liste = "-".join([f"{liste_deger_formatla(hp)}/{liste_deger_formatla(kg)}" for hp, kg in zip(secilen["hp_liste"], secilen["kilo_liste"])])
    son_medyan_hp_liste_ham = list(secilen["hp_liste"])
    son_medyan_kilo_liste_ham = list(secilen["kilo_liste"])

    print(f"MEDYAN OK: {at_adi} | tablo={secilen['tablo_no']} | n={secilen['n']} | HP_LISTE=[{son_medyan_hp_liste}] | HP={secilen['hp_ref']} | KILO={secilen['kilo_ref']}")

    sonuc = (secilen["hp_ref"], secilen["kilo_ref"])
    detay_referans_cache[detay_url] = sonuc
    return sonuc


def kosu_ici_hp_medyani_referans_hesapla_cacheli(row, detay_url, at_adi=""):
    global son_medyan_hp_liste, son_medyan_kilo_liste, son_medyan_eslesme_liste
    global son_medyan_hp_liste_ham, son_medyan_kilo_liste_ham
    detay_cache_sayac["toplam_ihtiyac"] += 1

    # Koşu kimliği (Tarih|Şehir|Msf|Pist|Kcins). Boşsa kütüphane kullanılmaz.
    kimlik = ""
    if _kut_conn is not None:
        try:
            kimlik = kosu_cache_anahtari(row)
        except Exception:
            kimlik = ""

    # 1) KALICI KÜTÜPHANE İSABETİ: geçmiş koşu daha önce çekilmişse TJK'ya GİTME.
    if _kut_conn is not None and kimlik:
        kayit = _KUT.medyan_getir(_kut_conn, kimlik)
        if kayit is not None and kayit.get("hp_medyan") is not None:
            detay_cache_sayac["kimlik_cache_isabet"] += 1
            hp_l = kayit["hp_liste"]
            kilo_l = kayit["kilo_liste"]
            # Kontrol kolonları (Medyan HP/Kilo/Eşleşme Listesi) birebir aynı kalsın.
            son_medyan_hp_liste_ham = list(hp_l)
            son_medyan_kilo_liste_ham = list(kilo_l)
            son_medyan_hp_liste = "-".join(liste_deger_formatla(x) for x in hp_l)
            son_medyan_kilo_liste = "-".join(liste_deger_formatla(x) for x in kilo_l)
            son_medyan_eslesme_liste = "-".join(
                f"{liste_deger_formatla(hp)}/{liste_deger_formatla(kg)}"
                for hp, kg in zip(hp_l, kilo_l))
            return (kayit["hp_medyan"], kayit["kilo_medyan"])

    # 2) URL yoksa eski davranış: hata.
    if not detay_url:
        detay_cache_sayac["kimlik_cache_pas_gecildi"] += 1
        raise Exception(f"Koşu detay URL bulunamadı: {at_adi}")

    # 3) ESKİ YOL: TJK'dan çek. Globaller (ham listeler dahil) burada set edilir.
    sonuc = kosu_ici_hp_medyani_referans_hesapla(detay_url, at_adi=at_adi)

    # 4) KÜTÜPHANEYE YAZ (sonraki çalıştırmalar bir daha çekmesin). Sadece geçerli
    #    medyan ve dolu ham listeler varsa yazılır; boş/şüpheli sonuç yazılmaz.
    if (_kut_conn is not None and kimlik and isinstance(sonuc, tuple)
            and len(sonuc) == 2 and sonuc[0] is not None
            and son_medyan_hp_liste_ham):
        _KUT.medyan_yaz(
            _kut_conn, kimlik,
            hp_medyan=sonuc[0], kilo_medyan=sonuc[1],
            hp_liste=son_medyan_hp_liste_ham,
            kilo_liste=son_medyan_kilo_liste_ham,
            detay_url=detay_url, kaynak="derece")

    return sonuc

def kosu_tarih_linkleri_cikar(at_sayfa_html, base_url):
    """At geçmiş tablosundaki tarih satırlarının tıklanan linklerini çıkarır."""
    if BeautifulSoup is None:
        return []

    soup = BeautifulSoup(at_sayfa_html, "html.parser")
    kayitlar = []

    for tr in soup.find_all("tr"):
        tr_text = " ".join(tr.get_text(" ", strip=True).split())
        if not re.search(r"\b\d{1,2}\.\d{1,2}\.\d{4}\b", tr_text):
            continue

        a_tags = tr.find_all("a", href=True)
        if not a_tags:
            continue

        hrefler = []
        for a in a_tags:
            href = a.get("href", "").strip()
            if not href or href.startswith("javascript"):
                continue
            hrefler.append(urljoin(base_url, href))

        if not hrefler:
            continue

        tarih_match = re.search(r"\b\d{1,2}\.\d{1,2}\.\d{4}\b", tr_text)
        kayitlar.append({
            "tarih": tarih_match.group(0) if tarih_match else "",
            "text": tr_text,
            "hrefler": hrefler,
        })

    return kayitlar


def gecmis_kosu_detay_url_bul(link_kayitlari, row):
    """Geçmiş koşu satırına en uygun detay URL'sini bulur."""
    tarih = str(satir_kolon_getir(row, "Tarih")).strip()
    sehir = str(satir_kolon_getir(row, "Şehir")).strip().upper()
    msf = str(satir_kolon_getir(row, "Msf")).strip()
    kcins = str(satir_kolon_getir(row, "Kcins")).strip().upper()

    adaylar = [x for x in link_kayitlari if x.get("tarih") == tarih]
    if not adaylar:
        return ""

    def puanla(k):
        txt = k.get("text", "").upper()
        puan = 0
        if sehir and sehir in txt:
            puan += 3
        if msf and msf in txt:
            puan += 2
        if kcins and kcins.split("/")[0].strip() and kcins.split("/")[0].strip() in txt:
            puan += 2
        return puan

    adaylar = sorted(adaylar, key=puanla, reverse=True)
    hrefler = adaylar[0].get("hrefler", [])

    # Koşu detayını çağrıştıran link varsa onu öne al.
    for h in hrefler:
        hu = h.upper()
        if "KOSU" in hu or "YARIS" in hu or "RESULT" in hu or "SONUC" in hu:
            return h

    return hrefler[0] if hrefler else ""




# =========================================
# GERCEK YARIS SONUCU LINKI / TABLO SECIMI
# =========================================

def sonuc_linki_mi(href):
    h = str(href or "").upper()
    return "GUNLUKYARISSONUCLARI" in h or "YARISSONUCLARI" in h


def satirdan_sonuc_linki_bul(tr, tarih):
    """Geçmiş satırındaki gerçek yarış sonucu linkini seçer.

    Eski hata: tr içindeki ilk linke basmak tırt. Bazen Video/link/aynı sayfa linki geliyor.
    Yeni kural: href içinde GunlukYarisSonuclari olmayan link yarış detayı değildir.
    """
    links = tr.find_elements(By.TAG_NAME, "a")
    adaylar = []
    for a in links:
        try:
            href = (a.get_attribute("href") or "").strip()
            txt = " ".join((a.text or "").split())
            if not href:
                continue
            if sonuc_linki_mi(href):
                puan = 10
                if tarih and tarih.replace(".", "/") in href:
                    puan += 5
                if tarih and tarih in txt:
                    puan += 3
                adaylar.append((puan, href, a, txt))
        except Exception:
            continue
    if not adaylar:
        return None, None
    adaylar = sorted(adaylar, key=lambda x: x[0], reverse=True)
    return adaylar[0][1], adaylar[0][2]


def at_adi_sade(at_adi):
    """Programdaki/sonuçtaki parantez numaralarını temizler.

    Örnek:
    - ALMUTANABİY (6) -> ALMUTANABİY
    - ALMUTANABİY(4) KG -> ALMUTANABİY
    Buradaki parantezler bazen koşu/program/start numarasıdır; at adı kontrolünde
    bunları kullanmak tırt, doğru tabloyu reddeder.
    """
    t = str(at_adi or "").strip()
    t = re.sub(r"\([^)]*\)", " ", t)
    t = re.sub(r"\b(KG|K|DB|SK|SGKR|GKR|YP|BB|DS|TS|AP|DILBAĞI|KULAKLIK)\b", " ", t, flags=re.IGNORECASE)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def tablo_at_adi_iceriyor_mu(df, at_adi):
    if not at_adi:
        return False
    hedef = metin_norm_at(at_adi_sade(at_adi))
    if hedef == "":
        return False
    try:
        metin = " ".join(df.astype(str).fillna("").values.ravel().tolist())
    except Exception:
        return False
    return hedef in metin_norm_at(metin)


def yaris_sonuc_sayfasi_mi(url, html):
    u = str(url or "").upper()
    h = str(html or "").upper()
    return ("GUNLUKYARISSONUCLARI" in u) or ("YARIŞ SONUÇLARI" in h) or ("YARIS SONUCLARI" in h)

# =========================================
# TARIH LINKINE GERCEKTEN TIKLAYARAK DETAY OKUMA
# =========================================

def row_degeri_norm(row, ad):
    return str(satir_kolon_getir(row, ad)).strip()


def gecmis_kosu_detay_html_tarihe_tikla(at_url, row, at_adi="", max_deneme=1):
    """At geçmişindeki ilgili tarih satırından GERÇEK sonuç href'ini alır ve HREF'i açar.

    Tıklama TJK'da güvenilir değil: link bulunduğu halde Selenium click bazen URL'yi hiç değiştirmiyor.
    Direkt URL açma daha önce tırtladı çünkü #anchor kaybolunca bütün gün tablolarından yanlış koşu seçiliyordu.
    Bu sürümün kilidi şu:
    - href içindeki #anchor ASLA kaybedilmez; detay_url olarak aynen href döner.
    - driver.get(href) ile sonuç sayfası açılır.
    - tablo okuma aşamasında sadece row'daki K. No-K. Adı / anchor bloğu kesilir.
    - AtKosuBilgileri HTML'inden medyan almak yasak.
    """
    global driver

    tarih = row_degeri_norm(row, "Tarih")
    sehir = row_degeri_norm(row, "Şehir").upper()
    msf = row_degeri_norm(row, "Msf")
    kcins = row_degeri_norm(row, "Kcins").upper()

    if not tarih:
        raise Exception("Tarih boş; sonuç linki bulunamaz")

    son_hata = None

    for deneme in range(1, max_deneme + 1):
        try:
            driver = guvenli_get(driver, at_url, max_deneme=2, bekleme=2)
            time.sleep(2)

            satirlar = driver.find_elements(By.TAG_NAME, "tr")
            adaylar = []

            for tr in satirlar:
                try:
                    txt = " ".join(tr.text.split())
                    if tarih not in txt:
                        continue

                    href, link_el = satirdan_sonuc_linki_bul(tr, tarih)
                    if not href:
                        continue

                    puan = 0
                    txt_up = txt.upper()
                    if sehir and sehir in txt_up:
                        puan += 4
                    if msf and msf in txt:
                        puan += 3
                    if kcins and kcins.split("/")[0].strip() and kcins.split("/")[0].strip() in txt_up:
                        puan += 2

                    adaylar.append((puan, txt, href))
                except Exception:
                    continue

            if not adaylar:
                raise Exception(f"Gerçek yarış sonucu linki bulunamadı: {at_adi} | {tarih} | {msf}. Satırda GunlukYarisSonuclari href yok.")

            adaylar = sorted(adaylar, key=lambda x: x[0], reverse=True)
            puan, satir_txt, href = adaylar[0]

            if "#" not in href:
                raise Exception(f"Sonuç href'inde anchor yok; doğru koşu bloğu garanti değil: {href}")

            print(f"TARIH SONUC HREF: {at_adi} | {tarih} | {msf} | puan={puan} | href={href}")
            print("HREF+ANCHOR MODU: tıklama yok; href açılacak, tablo sadece ilgili koşu bloğundan okunacak.")

            driver = guvenli_get(driver, href, max_deneme=2, bekleme=2)

            hedef_ok = False
            for bekle_no in range(1, 21):
                time.sleep(1)
                final_url = driver.current_url or ""
                html = driver.page_source or ""
                html_up = html.upper()
                url_up = final_url.upper()

                if bekle_no in [3, 6, 10, 15, 20]:
                    print(f"BEKLE {bekle_no}/20 | URL={final_url}")

                if (
                    "ATKOSUBILGILERI" not in url_up
                    and ("GUNLUKYARISSONUCLARI" in url_up or "YARIŞ SONUÇLARI" in html_up or "YARIS SONUCLARI" in html_up)
                    and len(str(html).strip()) > 1000
                ):
                    hedef_ok = True
                    break

            html = driver.page_source or ""
            final_url = driver.current_url or ""

            if not hedef_ok:
                raise Exception(f"Href ile yarış sonuç sayfası gelmedi. final_url={final_url}")

            # Render/lazy tetikle.
            for js in [
                "window.scrollTo(0, document.body.scrollHeight);",
                "window.scrollTo(0, 0);",
                "window.scrollTo(0, document.body.scrollHeight/2);",
            ]:
                try:
                    driver.execute_script(js)
                    time.sleep(1)
                except Exception:
                    pass

            html = driver.page_source or html

            if "ATKOSUBILGILERI" in str(final_url).upper():
                raise Exception(f"AtKosuBilgileri sayfasında kaldı; medyan yasak. final_url={final_url}")

            if not yaris_sonuc_sayfasi_mi(href, html):
                raise Exception(f"Yarış sonuç sayfası HTML'i doğrulanamadı. href={href} final_url={final_url}")

            if html is None or len(str(html).strip()) < 1000:
                raise Exception("Href açıldı ama HTML boş/eksik")

            # KRİTİK: final_url değil, href döndürülür. Çünkü final_url bazen #anchor'ı düşürüyor.
            return href, html

        except Exception as e:
            son_hata = e
            print(f"HREF ACMA HATASI ({deneme}/{max_deneme}): {at_adi} | {tarih} | {e}")
            time.sleep(2)

    raise Exception(f"Href ile detay açılamadı: {at_adi} | {tarih} | Son hata: {son_hata}")

def kosu_no_satirdan_al(row):
    """At geçmiş satırındaki K. No-K. Adı kolonundan koşu numarasını alır.

    Örnek: ALMUTANABİY 22.04.2026 satırında bu değer 8.
    Sonuç sayfasında bütün günün tabloları olduğu için doğru bloğu bu numarayla kesmek şart.
    """
    val = str(satir_kolon_getir(row, "K. No-K. Adı")).strip()
    m = re.search(r"(\d+)", val)
    if m:
        return int(m.group(1))
    return None


def html_kosu_blogu_kes(html, row=None, detay_url="", at_adi=""):
    """Yarış sonuç HTML'inden SADECE ilgili koşu bloğunu keser.

    23.05.2026 gibi günlerde eski anchor-yakını mantığı başka koşunun tablosunu
    seçebiliyordu. Bu sürümde anchor yardımcı sinyal; ana kilit K. No-K. Adı
    içindeki koşu numarasıdır.

    Sert kurallar:
    - row içinden koşu no alınabiliyorsa sadece aynı "N. Koşu" bloğu kabul edilir.
    - Seçilen blokta hedef at adı yoksa medyan yazılmaz.
    - Bütün gün sayfasındaki başka tabloya düşmek yasak.
    """
    h = str(html or "")
    if len(h) < 500:
        return h

    kosu_no = kosu_no_satirdan_al(row) if row is not None else None
    tarih = row_degeri_norm(row, "Tarih") if row is not None else ""
    sehir = row_degeri_norm(row, "Şehir").upper() if row is not None else ""
    msf = row_degeri_norm(row, "Msf") if row is not None else ""
    kcins = row_degeri_norm(row, "Kcins").upper() if row is not None else ""

    hedef_at = metin_norm_at(at_adi_sade(at_adi))

    def blok_at_iceriyor_mu(segment):
        if not hedef_at:
            return True
        return hedef_at in metin_norm_at(segment)

    def blok_skor(segment):
        """Aynı koşu no birden fazla kez görünürse doğru sonucu seçmek için puan."""
        seg_up = segment.upper()
        puan = 0
        if blok_at_iceriyor_mu(segment):
            puan += 10000
        if sehir and sehir in seg_up:
            puan += 300
        if msf and str(msf) in segment:
            puan += 200
        if kcins and kcins.split("/")[0].strip() and kcins.split("/")[0].strip() in seg_up:
            puan += 100
        if re.search(r"At\s*İsmi|At\s+Ismi|Sıklet|Siklet|HP", segment, flags=re.IGNORECASE):
            puan += 50
        return puan

    def segment_temizle(segment):
        bitisler = []
        for bp in ["Dağıtılacak", "Dagitilacak", "KURUMSAL", "Bülten", "BULTEN"]:
            i = segment.upper().find(bp.upper())
            if i > 1500:
                bitisler.append(i)
        if bitisler:
            segment = segment[:min(bitisler)]
        return segment

    # 1) Esas yöntem: HTML'i "N. Koşu" başlıklarına böl, row'daki koşu no ile birebir eşleştir.
    # Eski kod burada tırtlıyordu: anchor'ın yakınındaki önceki/sonraki koşuyu alabiliyordu.
    pat_heading = re.compile(r"(\d+)\s*\.\s*Ko[şs]u", flags=re.IGNORECASE)
    headings = [(int(m.group(1)), m.start(), m.group(0)) for m in pat_heading.finditer(h)]

    if headings:
        adaylar = []
        for i, (no, pos, title) in enumerate(headings):
            if kosu_no is not None and no != kosu_no:
                continue
            next_pos = headings[i + 1][1] if i + 1 < len(headings) else min(len(h), pos + 120000)
            segment = segment_temizle(h[pos:next_pos])
            if len(segment) <= 500:
                continue
            if not re.search(r"At\s*İsmi|At\s+Ismi|Sıklet|Siklet|HP", segment, flags=re.IGNORECASE):
                continue
            adaylar.append((blok_skor(segment), no, pos, segment))

        if adaylar:
            adaylar = sorted(adaylar, key=lambda x: x[0], reverse=True)
            skor, no, pos, segment = adaylar[0]
            if blok_at_iceriyor_mu(segment):
                print(f"KOSU BLOK OK: {at_adi} | tarih={tarih} | kosu_no={kosu_no} | secilen={no}. Koşu | skor={skor}")
                return segment
            raise Exception(
                f"{kosu_no}. Koşu bloğu bulundu ama hedef at adı blokta yok; yanlış koşudan medyan yazmak yasak: "
                f"{at_adi} | {detay_url}"
            )

    # 2) Koşu başlığı bulunamazsa anchor fallback. Ama row'da koşu no varsa bu segmentte
    # ilgili at adı ve tablo izi zorunlu; aksi halde boş geç.
    m = re.search(r"#([A-Za-z0-9_-]+)", str(detay_url or ""))
    if m:
        anchor = m.group(1)
        for ap in [
            rf'id=["\\\']{re.escape(anchor)}["\\\']',
            rf'name=["\\\']{re.escape(anchor)}["\\\']',
            re.escape(anchor),
        ]:
            am = re.search(ap, h, flags=re.IGNORECASE)
            if not am:
                continue

            # Anchor çevresi geniş ama sınırlı. Burada koşu no başlığı yoksa bile at adı şart.
            left = max(0, am.start() - 45000)
            right = min(len(h), am.start() + 90000)
            segment = segment_temizle(h[left:right])
            if len(segment) <= 500:
                continue
            if not blok_at_iceriyor_mu(segment):
                continue
            if not re.search(r"At\s*İsmi|At\s+Ismi|Sıklet|Siklet|HP", segment, flags=re.IGNORECASE):
                continue
            print(f"ANCHOR FALLBACK BLOK OK: {at_adi} | tarih={tarih} | kosu_no={kosu_no} | anchor={anchor}")
            return segment

    raise Exception(
        f"İlgili koşu bloğu kesin bulunamadı; medyan yazmak yasak: {at_adi} | "
        f"tarih={tarih} | sehir={sehir} | msf={msf} | kcins={kcins} | kosu_no={kosu_no} | {detay_url}"
    )


def soup_tablodan_hp_kilo_adaylari(html_blok, at_adi=""):
    """Sonuç tablosunu pandas yerine satır satır okur.

    Neden:
    pd.read_html bazı TJK tablolarında ilk veri satırını header gibi yutuyor.
    23.05.2026 PİRNUR 67 / 60,5 tam olarak böyle düşmüştü.
    Bu fonksiyon tr/td/th hücrelerini elle okuyup ilk atı da listeye dahil eder.
    """
    if BeautifulSoup is None:
        return []

    soup = BeautifulSoup(str(html_blok or ""), "html.parser")
    adaylar = []
    hedef = metin_norm_at(at_adi_sade(at_adi))

    for tablo_no, table in enumerate(soup.find_all("table"), start=1):
        rows = []
        for tr in table.find_all("tr"):
            cells = tr.find_all(["th", "td"])
            vals = [" ".join(c.get_text(" ", strip=True).split()) for c in cells]
            vals = [v for v in vals if v is not None]
            if vals:
                rows.append(vals)

        if not rows:
            continue

        tablo_text = " ".join(" ".join(r) for r in rows)
        if hedef and hedef not in metin_norm_at(tablo_text):
            continue

        header_idx = None
        hp_idx = None
        kilo_idx = None
        at_idx = None

        for i, vals in enumerate(rows[:8]):
            norms = [re.sub(r"[^A-Z0-9]", "", kolon_adi_norm(v)) for v in vals]
            # At İsmi / Sıklet / HP başlıklarını aynı satırda arıyoruz.
            possible_hp = [j for j, n in enumerate(norms) if n == "HP"]
            possible_kilo = [j for j, n in enumerate(norms) if n == "SIKLET"]
            possible_at = [j for j, n in enumerate(norms) if n in ["ATISMI", "ATADI", "AT"] or "ATISMI" in n]
            if possible_hp and possible_kilo:
                header_idx = i
                hp_idx = possible_hp[-1]
                kilo_idx = possible_kilo[0]
                at_idx = possible_at[0] if possible_at else 0
                break

        if header_idx is None:
            continue

        hp_liste = []
        kilo_liste = []
        eslesen_atlar = []

        for vals in rows[header_idx + 1:]:
            # Bazı nested/header satırları araya girebilir.
            if len(vals) <= max(hp_idx, kilo_idx):
                continue

            satir_text = " ".join(vals)
            satir_norm = metin_norm_at(satir_text)
            # Header tekrarıysa geç.
            if "ATISMI" in satir_norm and "SIKLET" in satir_norm:
                continue

            hp_raw = vals[hp_idx]
            kilo_raw = vals[kilo_idx]
            hp_num = sayi_float_temizle(hp_raw)
            kilo_num = siklet_float_temizle(kilo_raw)

            if hp_num is None or kilo_num is None:
                continue
            if not (1 <= float(hp_num) <= 180):
                continue
            if not (35 <= float(kilo_num) <= 75):
                continue

            hp_liste.append(float(hp_num))
            kilo_liste.append(float(kilo_num))
            if at_idx is not None and len(vals) > at_idx:
                eslesen_atlar.append(vals[at_idx])
            else:
                eslesen_atlar.append(vals[0] if vals else "")

        n = len(hp_liste)
        # KRİTİK: Küçük koşuları reddetme. 4 atlı hatta daha küçük koşuda da medyan hesaplanır.
        # Eski n < 5 filtresi gerçek ilk-4 kayıtlarını gereksiz MEDYAN_HATA yapıyordu.
        if n < 1 or n > 24:
            continue

        at_var = True
        if hedef:
            at_var = any(hedef in metin_norm_at(a) for a in eslesen_atlar) or hedef in metin_norm_at(tablo_text)
        if not at_var:
            continue

        hp_ref = float(pd.Series(hp_liste).median())
        kilo_ref = float(pd.Series(kilo_liste).median())

        adaylar.append({
            "kaynak": "SOUP_TR_TD_ILK_SATIR_FIX",
            "tablo_no": tablo_no,
            "n": n,
            "puan": 20000 + n * 100,
            "at_var": at_var,
            "hp_ref": hp_ref,
            "kilo_ref": round(kilo_ref, 1),
            "hp_liste": hp_liste,
            "kilo_liste": kilo_liste,
            "atlar": eslesen_atlar,
        })

    return adaylar


def html_icinden_hp_siklet_medyani(html, detay_url="", at_adi="", row=None):
    """Yarış sonucu HTML'inden HP/Sıklet medyanını alır.

    KRİTİK FIX:
    pandas read_html bazı yarışlarda ilk veri satırını düşürüyor.
    Bu yüzden önce BeautifulSoup ile tablo satırlarını elle okuyoruz.
    Böylece ilk atın HP/Kilosu, örn. PİRNUR 67 / 60,5, listeye girer.
    """
    global son_medyan_hp_liste, son_medyan_kilo_liste, son_medyan_eslesme_liste
    global son_medyan_hp_liste_ham, son_medyan_kilo_liste_ham
    son_medyan_hp_liste = ""
    son_medyan_kilo_liste = ""
    son_medyan_eslesme_liste = ""
    son_medyan_hp_liste_ham = []
    son_medyan_kilo_liste_ham = []

    if not yaris_sonuc_sayfasi_mi(detay_url, html):
        raise Exception(f"Bu HTML yarış sonuç sayfası değil; medyan alınmaz: {detay_url}")

    if "ATKOSUBILGILERI" in str(detay_url).upper():
        raise Exception(f"At geçmiş sayfasından medyan almak yasak: {detay_url}")

    html_blok = html_kosu_blogu_kes(html, row=row, detay_url=detay_url, at_adi=at_adi)

    adaylar = soup_tablodan_hp_kilo_adaylari(html_blok, at_adi=at_adi)

    # Soup hiç tablo çıkaramazsa eski pandas yolu son çare. Ama pandas yolu artık öncelik değil.
    if not adaylar:
        redler = []
        try:
            tables = pd.read_html(StringIO(html_blok))
        except Exception as e:
            debug_kaydet(at_adi, detay_url, row=row, kaynak="SOUP_AND_PANDAS_FAIL", hata=str(e))
            raise Exception(f"Sonuç sayfasında tablo okunamadı: {at_adi} | {detay_url} | {e}")

        for tablo_no, df in enumerate(tables, start=1):
            tmp = hp_kilo_kolonlarini_bul(df)
            if tmp is None:
                continue
            n = len(tmp)
            # KRİTİK: n<5 filtresi kaldırıldı. 2-3-4 atlı koşular da gerçek koşudur.
            if n < 1:
                redler.append(f"pandas tablo {tablo_no}: HP/Sıklet var ama n<1 ({n})")
                continue
            if n > 24:
                redler.append(f"pandas tablo {tablo_no}: fazla uzun tablo n={n}")
                continue
            if not tablo_at_adi_iceriyor_mu(df, at_adi):
                redler.append(f"pandas tablo {tablo_no}: hedef at adı yok")
                continue
            hp_liste = [float(x) for x in tmp["HP_NUM"].tolist()]
            kilo_liste = [float(x) for x in tmp["KILO_NUM"].tolist()]
            adaylar.append({
                "kaynak": "PANDAS_FALLBACK",
                "tablo_no": tablo_no,
                "n": n,
                "puan": 10000 + n * 100,
                "at_var": True,
                "hp_ref": float(pd.Series(hp_liste).median()),
                "kilo_ref": round(float(pd.Series(kilo_liste).median()), 1),
                "hp_liste": hp_liste,
                "kilo_liste": kilo_liste,
            })

        if not adaylar:
            detay = " | ".join(redler[:12])
            debug_kaydet(at_adi, detay_url, row=row, kaynak="NO_HP_KILO_TABLE", hata=detay)
            raise Exception(f"İlgili koşu bloğunda HP + Sıklet tablosu bulunamadı: {at_adi} | {detay_url} | {detay}")

    secilen = sorted(adaylar, key=lambda x: (x.get("at_var", False), x.get("n", 0), x.get("puan", 0)), reverse=True)[0]

    son_medyan_hp_liste = "-".join([liste_deger_formatla(x) for x in secilen["hp_liste"]])
    son_medyan_kilo_liste = "-".join([liste_deger_formatla(x) for x in secilen["kilo_liste"]])
    son_medyan_eslesme_liste = "-".join([
        f"{liste_deger_formatla(hp)}/{liste_deger_formatla(kg)}"
        for hp, kg in zip(secilen["hp_liste"], secilen["kilo_liste"])
    ])
    son_medyan_hp_liste_ham = list(secilen["hp_liste"])
    son_medyan_kilo_liste_ham = list(secilen["kilo_liste"])

    print(
        f"MEDYAN OK: {at_adi} | kaynak={secilen['kaynak']} | tablo={secilen['tablo_no']} | "
        f"n={secilen['n']} | HP_LISTE=[{son_medyan_hp_liste}] | HP={secilen['hp_ref']} | "
        f"KILO_LISTE=[{son_medyan_kilo_liste}] | KILO={secilen['kilo_ref']} | "
        f"ESLESME=[{son_medyan_eslesme_liste}] | URL={detay_url}"
    )

    return secilen["hp_ref"], secilen["kilo_ref"]

def kosu_ici_hp_medyani_tiklayarak_hesapla(at_url, row, at_adi=""):
    """Medyan/HP-Kilo için 3 kez dener; olmazsa dıştaki except'e bırakır.

    Dış akış eski mantıkta kalır:
    - satır Veri sayfasına yazılır,
    - medyan/link eksikse Basarisiz_Atlar'a not düşülür,
    - sonraki ata geçilir.
    """
    global son_medyan_hp_liste, son_medyan_kilo_liste, son_medyan_eslesme_liste
    global son_medyan_hp_liste_ham, son_medyan_kilo_liste_ham

    # === KALICI KÜTÜPHANE İSABETİ ===
    # Koşu kimliği (Tarih|Şehir|Msf|Pist|Kcins) daha önce çekilmişse, TJK'ya
    # (Selenium tıklama) HİÇ gitme. Kayıtlı medyan/kilo + listeleri döndür.
    kimlik = ""
    if _kut_conn is not None:
        try:
            kimlik = kosu_cache_anahtari(row)
        except Exception:
            kimlik = ""
        if kimlik:
            kayit = _KUT.medyan_getir(_kut_conn, kimlik)
            if kayit is not None and kayit.get("hp_medyan") is not None:
                detay_cache_sayac["toplam_ihtiyac"] += 1
                detay_cache_sayac["kimlik_cache_isabet"] += 1
                hp_l = kayit["hp_liste"]
                kilo_l = kayit["kilo_liste"]
                # Kontrol kolonları (Medyan HP/Kilo/Eşleşme Listesi) birebir aynı kalsın.
                son_medyan_hp_liste_ham = list(hp_l)
                son_medyan_kilo_liste_ham = list(kilo_l)
                son_medyan_hp_liste = "-".join(liste_deger_formatla(x) for x in hp_l)
                son_medyan_kilo_liste = "-".join(liste_deger_formatla(x) for x in kilo_l)
                son_medyan_eslesme_liste = "-".join(
                    f"{liste_deger_formatla(hp)}/{liste_deger_formatla(kg)}"
                    for hp, kg in zip(hp_l, kilo_l))
                return kayit.get("detay_url", ""), kayit["hp_medyan"], kayit["kilo_medyan"]

    son_hata = None

    for deneme in range(1, MAX_MEDYAN_DENEME + 1):
        detay_cache_sayac["toplam_ihtiyac"] += 1
        detay_cache_sayac["gercek_detay_acilis"] += 1

        try:
            detay_url, html = gecmis_kosu_detay_html_tarihe_tikla(at_url, row, at_adi=at_adi)
            hp_ref, kilo_ref = html_icinden_hp_siklet_medyani(html, detay_url=detay_url, at_adi=at_adi, row=row)

            if hp_ref == "" or kilo_ref == "":
                raise Exception("HP/Kilo medyan boş döndü")

            # === KÜTÜPHANEYE YAZ === (sonraki çalıştırmalar bir daha çekmesin)
            if (_kut_conn is not None and kimlik and hp_ref is not None
                    and son_medyan_hp_liste_ham):
                _KUT.medyan_yaz(
                    _kut_conn, kimlik,
                    hp_medyan=hp_ref, kilo_medyan=kilo_ref,
                    hp_liste=son_medyan_hp_liste_ham,
                    kilo_liste=son_medyan_kilo_liste_ham,
                    detay_url=detay_url, kaynak="derece")

            return detay_url, hp_ref, kilo_ref

        except Exception as e:
            son_hata = e
            print(f"MEDYAN/HP-KILO ALINAMADI - TEKRAR DENENECEK ({deneme}/{MAX_MEDYAN_DENEME}): {at_adi} | {e}")
            try:
                debug_kaydet(at_adi, "", row=row, kaynak="MEDYAN_3_DENEME", hata=str(e))
            except Exception:
                pass
            time.sleep(3)

    raise Exception(f"Medyan HP/Kilo {MAX_MEDYAN_DENEME} denemede alınamadı. Satır Veri'ye yazılacak, medyan boş kalacak. Son hata: {son_hata}")


def driver_baslat():
    options = webdriver.ChromeOptions()

    if HEADLESS_CHROME:
        # Mac'te pencere açmadan arkada çalışır. Daha az görsel yük = biraz daha hızlı.
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
    else:
        options.add_argument("--start-maximized")

    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.page_load_strategy = "eager"

    d = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    d.set_page_load_timeout(60)
    return d


def guvenli_get(driver, url, max_deneme=5, bekleme=5):
    son_hata = None

    while True:
        try:
            driver.get(url)
            time.sleep(SELENIUM_BEKLEME_SN)
            return driver
        except (NoSuchWindowException, WebDriverException) as e:
            son_hata = e
            print(f"CHROME HATASI - DRIVER YENİDEN AÇILIYOR - HATA: {e}")

            try:
                driver.quit()
            except:
                pass

            time.sleep(bekleme)
            driver = driver_baslat()


def guvenli_linkleri_al(driver, max_deneme=5, bekleme=5):
    son_hata = None

    for deneme in range(1, max_deneme + 1):
        try:
            return driver.find_elements(By.TAG_NAME, "a"), driver
        except (NoSuchWindowException, WebDriverException) as e:
            son_hata = e
            print(f"LINK ALMA HATASI - DRIVER YENİDEN AÇILIYOR ({deneme}/{max_deneme}) - HATA: {e}")

            try:
                current = driver.current_url
            except:
                current = ANA_URL

            try:
                driver.quit()
            except:
                pass

            time.sleep(bekleme)
            driver = driver_baslat()
            driver = guvenli_get(driver, current, max_deneme=2, bekleme=bekleme)

    raise Exception(f"Linkler alınamadı. Son hata: {son_hata}")


def excel_temizle(df):
    """Excel'e nan/None yazma; boş bırak."""
    if df is None or len(df) == 0:
        return df
    return df.replace([pd.NA, float("inf"), float("-inf")], "").fillna("").replace({"nan": "", "NaN": "", "None": ""})


def ara_kaydet(tum_sonuclar, basarisiz_atlar):
    sonuc_df = excel_temizle(pd.DataFrame(tum_sonuclar))
    basarisiz_df = excel_temizle(pd.DataFrame(basarisiz_atlar))
    cache_rapor_df = excel_temizle(pd.DataFrame([detay_cache_sayac]))
    debug_df = excel_temizle(pd.DataFrame(raw_debug_tablolar))

    with pd.ExcelWriter(EXCEL_NAME, engine="openpyxl") as writer:
        sonuc_df.to_excel(writer, index=False, sheet_name="Veri")
        basarisiz_df.to_excel(writer, index=False, sheet_name="Basarisiz_Atlar")
        cache_rapor_df.to_excel(writer, index=False, sheet_name="Cache_Rapor")
        debug_df.to_excel(writer, index=False, sheet_name="Debug_Tablolar")


son_medyan_hp_liste = ""
son_medyan_kilo_liste = ""
son_medyan_eslesme_liste = ""
son_medyan_hp_liste_ham = []
son_medyan_kilo_liste_ham = []

# DEBUG: Yanlış medyan yazmak yerine ham tablo izlerini Excel'e döker.
raw_debug_tablolar = []

def debug_kaydet(at_adi, detay_url, row=None, kaynak="", hata=""):
    """Başarısız/şüpheli koşularda sayfada görünen ham tablo metinlerini Excel'e kaydeder."""
    global raw_debug_tablolar, driver
    tarih = satir_kolon_getir(row, "Tarih") if row is not None else ""
    sehir = satir_kolon_getir(row, "Şehir") if row is not None else ""
    msf = satir_kolon_getir(row, "Msf") if row is not None else ""
    kcins = satir_kolon_getir(row, "Kcins") if row is not None else ""

    try:
        tables = driver.find_elements(By.TAG_NAME, "table")
    except Exception as e:
        raw_debug_tablolar.append({
            "At Adı": at_adi, "Tarih": tarih, "Şehir": sehir, "Mesafe": msf, "Koşu Cinsi": kcins,
            "Kaynak": kaynak, "Detay URL": detay_url, "Tablo No": "", "Hata": hata + " | Selenium table okunamadı: " + str(e),
            "Tablo Metni": ""
        })
        return

    hedef = metin_norm_at(at_adi_sade(at_adi))
    bulunan = 0
    for i, t in enumerate(tables, start=1):
        try:
            txt = " ".join((t.text or "").split())
        except Exception:
            continue
        if not txt:
            continue
        txt_norm = metin_norm_at(txt)
        # Öncelik hedef at adı geçen veya HP/Sıklet izi olan tablolar.
        if hedef and hedef not in txt_norm and not ("HP" in txt.upper() and ("SIKLET" in txt.upper() or "SIKLET" in kolon_adi_norm(txt))):
            continue
        bulunan += 1
        raw_debug_tablolar.append({
            "At Adı": at_adi, "Tarih": tarih, "Şehir": sehir, "Mesafe": msf, "Koşu Cinsi": kcins,
            "Kaynak": kaynak, "Detay URL": detay_url, "Tablo No": i, "Hata": hata,
            "Tablo Metni": txt[:12000]
        })

    if bulunan == 0:
        try:
            body_txt = " ".join((driver.find_element(By.TAG_NAME, "body").text or "").split())
        except Exception:
            body_txt = ""
        raw_debug_tablolar.append({
            "At Adı": at_adi, "Tarih": tarih, "Şehir": sehir, "Mesafe": msf, "Koşu Cinsi": kcins,
            "Kaynak": kaynak, "Detay URL": detay_url, "Tablo No": "BODY", "Hata": hata + " | Uygun table yok",
            "Tablo Metni": body_txt[:12000]
        })



def sayfadaki_program_tarihi(driver):
    """Günlük yarış programı başlığındaki tarihi yakalar. Örn: 12/06/2026"""
    try:
        kaynak = driver.page_source or ""
    except Exception:
        kaynak = ""

    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+\w+\s*-\s*Yarış Programı", kaynak)
    if m:
        return m.group(1)

    try:
        body = driver.find_element(By.TAG_NAME, "body").text or ""
    except Exception:
        body = ""

    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+\w+\s*-\s*Yarış Programı", body)
    if m:
        return m.group(1)

    return ""


def ileri_programina_gec(driver, gun_sayisi=2):
    """TJK Günlük Yarış Programı sayfasında ileri gün okuna istenen gün sayısı kadar tıklar.

    Kural:
    - Ana sayfa bugünün programını açar.
    - gun_sayisi=2 ise yarının da ertesi gününe geçer.
    - Her tıklamada tarih değişimini bekler.
    - Tarih değişmezse yanlış gün çekmemek için durur.
    """
    print(f"\n{gun_sayisi} GÜN SONRAKİ PROGRAMA GEÇİLİYOR...")

    if gun_sayisi < 1:
        return driver

    for adim in range(1, gun_sayisi + 1):
        tarih_once = sayfadaki_program_tarihi(driver)
        print(f"İleri gün adımı {adim}/{gun_sayisi} | Program tarihi önce: {tarih_once if tarih_once else 'OKUNAMADI'}")

        ileri_tiklandi = False

        # 1) Normal Selenium yolu: görünen metni > / › / » olan link veya buton
        elemanlar = []
        try:
            elemanlar.extend(driver.find_elements(By.TAG_NAME, "a"))
        except Exception:
            pass
        try:
            elemanlar.extend(driver.find_elements(By.TAG_NAME, "button"))
        except Exception:
            pass

        for el in elemanlar:
            try:
                txt = (el.text or "").strip()
                title = (el.get_attribute("title") or "").strip().lower()
                aria = (el.get_attribute("aria-label") or "").strip().lower()
                cls = (el.get_attribute("class") or "").strip().lower()

                if txt in [">", "›", "»"] or "sonraki" in title or "ileri" in title or "next" in title or "sonraki" in aria or "ileri" in aria or "next" in aria or "next" in cls:
                    try:
                        el.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", el)
                    ileri_tiklandi = True
                    break
            except Exception:
                continue

        # 2) Fallback: JavaScript ile > metinli link/buton ara
        if not ileri_tiklandi:
            try:
                ileri_tiklandi = driver.execute_script("""
                    const els = Array.from(document.querySelectorAll('a,button'));
                    const e = els.find(x => {
                        const t = (x.innerText || x.textContent || '').trim();
                        const title = (x.getAttribute('title') || '').toLowerCase();
                        const aria = (x.getAttribute('aria-label') || '').toLowerCase();
                        return t === '>' || t === '›' || t === '»' ||
                               title.includes('sonraki') || title.includes('ileri') || title.includes('next') ||
                               aria.includes('sonraki') || aria.includes('ileri') || aria.includes('next');
                    });
                    if (e) { e.click(); return true; }
                    return false;
                """)
            except Exception:
                ileri_tiklandi = False

        if not ileri_tiklandi:
            raise Exception(f"İleri gün oku bulunamadı. Adım: {adim}/{gun_sayisi}")

        # Her tıklamada tarih değişene kadar bekle
        tarih_degisti = False
        for bekle_no in range(1, 31):
            time.sleep(1)
            tarih_sonra = sayfadaki_program_tarihi(driver)

            if bekle_no in [3, 6, 10, 20, 30]:
                print(f"İleri gün bekleme {bekle_no}/30 | Adım {adim}/{gun_sayisi} | Tarih: {tarih_sonra if tarih_sonra else 'OKUNAMADI'}")

            if tarih_sonra and tarih_once and tarih_sonra != tarih_once:
                print(f"İleri gün geçildi: {tarih_once} -> {tarih_sonra}")
                tarih_degisti = True
                break

            if tarih_sonra and not tarih_once:
                print(f"Program tarihi okundu: {tarih_sonra}")
                tarih_degisti = True
                break

        if not tarih_degisti:
            raise Exception(f"İleri gün okuna tıklandı ama program tarihi değişmedi. Adım: {adim}/{gun_sayisi}. Yanlış gün çekmemek için durduruldu.")

    return driver


def _ileri_tikla(driver):
    """Program sayfasında 'ileri/sonraki gün' okuna bir kez tıklar. Başarılıysa True."""
    try:
        els = driver.find_elements(By.TAG_NAME, "a") + driver.find_elements(By.TAG_NAME, "button")
    except Exception:
        els = []
    for el in els:
        try:
            txt = (el.text or "").strip()
            title = (el.get_attribute("title") or "").strip().lower()
            aria = (el.get_attribute("aria-label") or "").strip().lower()
            cls = (el.get_attribute("class") or "").strip().lower()
            if (txt in [">", "›", "»"] or "sonraki" in title or "ileri" in title or "next" in title
                    or "sonraki" in aria or "ileri" in aria or "next" in aria or "next" in cls):
                try:
                    el.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", el)
                return True
        except Exception:
            continue
    try:
        return bool(driver.execute_script("""
            const els = Array.from(document.querySelectorAll('a,button'));
            const e = els.find(x => {
                const t = (x.innerText || x.textContent || '').trim();
                const title = (x.getAttribute('title') || '').toLowerCase();
                const aria = (x.getAttribute('aria-label') || '').toLowerCase();
                return t === '>' || t === '›' || t === '»' ||
                       title.includes('sonraki') || title.includes('ileri') || title.includes('next') ||
                       aria.includes('sonraki') || aria.includes('ileri') || aria.includes('next');
            });
            if (e) { e.click(); return true; }
            return false;
        """))
    except Exception:
        return False


def hedef_gune_git(driver, hedef_dt):
    """Programı MUTLAK hedef tarihe (datetime) götürür. Site hangi günü gösterirse
    göstersin ileri tıklayarak hedefe ulaşır; zaten hedefteyse tıklamaz.
    Pipeline/galop ile aynı mantık -> gün ortasında site kaysa bile uyumlu kalır."""
    hedef = hedef_dt.strftime("%d/%m/%Y")
    print(f"\nHEDEF GÜN (mutlak): {hedef}")
    for _ in range(0, 8):
        simdi = sayfadaki_program_tarihi(driver)
        print(f"  sayfa: {simdi if simdi else 'OKUNAMADI'} | hedef: {hedef}")
        if simdi == hedef:
            print(f"  hedefe ulaşıldı: {hedef}")
            return driver
        if not _ileri_tikla(driver):
            raise Exception(f"İleri gün oku bulunamadı; hedef {hedef} sayfasına gidilemedi.")
        for _w in range(1, 31):
            time.sleep(1)
            s = sayfadaki_program_tarihi(driver)
            if s and s != simdi:
                break
    raise Exception(f"Hedef güne ({hedef}) ulaşılamadı; site farklı tarih gösteriyor.")



print("\n==============================")
print("TUM_ILLER_TUM_ATLAR_EKSIKSIZ_ILK4_ONCEKI_HP.py AKTIF")
print("ERTESI GUN + 3 MEDYAN DENEME MODU: Medyan gelmezse satir yine Veri sayfasina yazilir.")
print("==============================\n")

# =========================================
# DRIVER
# =========================================

driver = driver_baslat()


# =========================================
# ANA SAYFA
# =========================================

driver = guvenli_get(driver, ANA_URL, max_deneme=5, bekleme=5)
# ERTESI GUN MODU: Ana sayfa bugünü açar; bu yüzden sağ ok/ileri gün ile 2 gün ileri gitmek zorunlu.
driver = hedef_gune_git(driver, program_tarihi)
print(f"MUTLAK GÜN MODU: hedef tarih {program_tarihi.strftime('%d/%m/%Y')} (bugün+{PROGRAM_GUN_OFFSET}) programı çekilecek.")

print("\nAKTİF URL:")
print(driver.current_url)

print("\nSAYFA BAŞLIĞI:")
print(driver.title)


# =========================================
# LINKLER
# =========================================

links, driver = guvenli_linkleri_al(driver)

print("\nTOPLAM LINK:", len(links))


# =========================================
# ŞEHİR LİNKLERİ
# =========================================

sehir_linkleri = {}

for link in links:
    try:
        href = link.get_attribute("href")
        text = link.text.strip()

        if not href:
            continue

        if "GunlukYarisProgrami?SehirId=" not in href:
            continue

        for sehir in TURK_SEHIRLER:
            if sehir.lower() in text.lower():
                if sehir not in sehir_linkleri:
                    # BUGUN MODU: Sayfadaki şehir href'i aynen korunur.
                    sehir_linkleri[sehir] = href

    except:
        pass


print("\n=================================")
print("BULUNAN ŞEHİRLER")
print("=================================\n")

for s in sehir_linkleri:
    print(s)

print("\nTOPLAM ŞEHİR:", len(sehir_linkleri))

if TEST_MODU:
    # Şehir testini elle sabitle: İstanbul'a takılıp kalmasın.
    if TEST_SEHIR_ADI:
        sehir_linkleri = {k: v for k, v in sehir_linkleri.items() if k == TEST_SEHIR_ADI}
    else:
        sehir_linkleri = dict(list(sehir_linkleri.items())[:TEST_SEHIR_LIMIT])

    print("\nTEST MODU AKTİF")
    print("Test şehirleri:", list(sehir_linkleri.keys()))
    print("Test at başlangıç:", TEST_AT_BASLANGIC)
    print("Test at limiti / şehir:", TEST_AT_LIMIT)
    print("Test toplam at deneme limiti:", TEST_TOPLAM_AT_DENEME_LIMIT)


if len(sehir_linkleri) == 0:
    print("\nŞEHİR BULUNAMADI")
    print("\nTest bitti; pencere kapatılıyor...")
    driver.quit()
    exit()


# =========================================
# SONUÇLAR
# =========================================

tum_sonuclar = []
basarisiz_atlar = []


# =========================================
# HER ŞEHİR
# =========================================

for sehir_adi, sehir_url in sehir_linkleri.items():

    print(f"\n========== {sehir_adi} ==========\n")

    try:
        driver = guvenli_get(driver, sehir_url, max_deneme=5, bekleme=5)
        links, driver = guvenli_linkleri_al(driver, max_deneme=5, bekleme=5)
    except Exception as e:
        print("ŞEHİR HATASI:", sehir_adi)
        print(e)

        basarisiz_atlar.append({
            "Şehir": sehir_adi,
            "At ID": "",
            "At Adı": "",
            "URL": sehir_url,
            "Hata": "Şehir sayfası açılamadı: " + str(e)
        })

        ara_kaydet(tum_sonuclar, basarisiz_atlar)
        continue

    at_linkleri = {}

    for link in links:
        try:
            href = link.get_attribute("href")

            if not href:
                continue

            if "QueryParameter_AtId=" not in href:
                continue

            parsed = urlparse(href)
            params = parse_qs(parsed.query)

            if "QueryParameter_AtId" not in params:
                continue

            at_id = params["QueryParameter_AtId"][0]
            at_adi = link.text.strip()

            if not at_adi:
                continue

            hp = ""

            try:
                satir_text = link.find_element(By.XPATH, "./ancestor::tr").text
                m = re.search(r"\bHP\s*[:\-]?\s*(\d+)", satir_text, flags=re.IGNORECASE)
                if m:
                    hp = m.group(1)
            except:
                pass

            if at_id not in at_linkleri:
                at_linkleri[at_id] = {
                    "ad": at_adi,
                    "url": href,
                    "hp": hp
                }

        except:
            pass

    print("Bulunan at:", len(at_linkleri))

    if TEST_MODU:
        at_linkleri = dict(list(at_linkleri.items())[TEST_AT_BASLANGIC:TEST_AT_BASLANGIC + TEST_AT_LIMIT])
        print("TEST MODU - Başlangıç index:", TEST_AT_BASLANGIC)
        print("TEST MODU - İşlenecek at:", len(at_linkleri))


    # =====================================
    # HER AT
    # =====================================

    for at_id, bilgi in at_linkleri.items():

        global_test_stop = False
        if TEST_MODU:
            test_toplam_at_deneme_sayaci += 1
            print(f"TEST DENEME SAYACI: {test_toplam_at_deneme_sayaci}/{TEST_TOPLAM_AT_DENEME_LIMIT}")
            if test_toplam_at_deneme_sayaci > TEST_TOPLAM_AT_DENEME_LIMIT:
                print("TEST TOPLAM AT DENEME LIMITI DOLDU")
                global_test_stop = True
                break

        if TEST_MODU and len(tum_sonuclar) >= TEST_BASARILI_SATIR_LIMIT:
            print(f"TEST BASARILI SATIR LIMITI DOLDU, AT DONGUSU DURDU: {len(tum_sonuclar)}")
            break

        at_adi = bilgi["ad"]
        hp_program = bilgi.get("hp", "")

        try:
            if hp_program != "":
                print("AT:", at_adi, "HP:", hp_program)
            else:
                print("AT:", at_adi)

            at_url = (
                "https://www.tjk.org/TR/YarisSever/Query/ConnectedPage/"
                f"AtKosuBilgileri?1=1&QueryParameter_AtId={at_id}&Era=today"
            )

            tables, kullanilan_url, at_sayfa_html = read_html_at_retry(
                at_id=at_id,
                at_adi=at_adi,
                orijinal_href=bilgi.get("url", at_url),
                max_deneme=3,
                bekleme=3
            )

            if len(tables) < 2:
                raise Exception("At geçmiş koşu tablosu bulunamadı")

            link_kayitlari = kosu_tarih_linkleri_cikar(at_sayfa_html, kullanilan_url)

            df = tables[1]

            # SAYFA KESİKSE TAMAMLA: 'Toplam 62 sonuçtan 50 tanesi gösteriliyor' +
            # 'Daha Fazla Sonuç Göster' formu varsa devam sayfalarını da çek ->
            # en eski koşular gelsin ki İLK 3 HP GERÇEK değerleriyle hesaplansın.
            df = at_gecmisi_tamamla(df, at_sayfa_html, at_adi)

            ilk_3_hp_serisi, ilk_3_hp_tarihleri, ilk_3_hp_ham, ilk_3_hp_kcinsleri = atin_ilk_3_hp_bilgisi(df)

            # KRİTİK: İlk 4 filtrelemesi artık gevşek/sağlam.
            # "1", "1.", "1 ", "1/12" gibi varyasyonlar kaçmayacak.
            if "S" not in df.columns:
                raise Exception("At geçmiş tablosunda S kolonu yok; ilk-4 kontrolü yapılamadı")

            df3 = df[df["S"].apply(sira_ilk4_mi)].copy()

            if len(df3) == 0:
                print(f"ILK4 YOK: {at_adi} | son 6 ay filtresinden önce ilk-4 satırı bulunmadı")
                continue

            df3["TarihParsed"] = pd.to_datetime(
                df3["Tarih"],
                dayfirst=True,
                errors="coerce"
            )

            df3 = df3[
                (df3["TarihParsed"] >= alti_ay_once) &
                (df3["TarihParsed"] < pd.Timestamp(program_tarihi.date()))
            ]

            if len(df3) == 0:
                print(f"SON6AY ILK4 YOK: {at_adi}")
                continue

            for _, row in df3.iterrows():

                # KRİTİK: Bu noktaya gelen her satır Excel'e yazılacak.
                # Medyan/link/HP tablosu hatası satırı öldürmez; sadece Hata kolonuna yazılır.
                hp_final = onceki_kosu_hp_getir(df, row.name)
                hp_kaynak_notu = onceki_kosu_hp_debug_getir(df, row.name)

                kayit = satir_excel_kaydi_olustur(
                    row=row,
                    at_adi=at_adi,
                    hp_final=hp_final,
                    hp_kaynak_notu=hp_kaynak_notu,
                    ilk_3_hp_serisi=ilk_3_hp_serisi,
                    ilk_3_hp_kcinsleri=ilk_3_hp_kcinsleri,
                    ilk_3_hp_ham=ilk_3_hp_ham,
                )
                kayit["Kaynak URL"] = kullanilan_url

                try:
                    kosu_detay_url, kosu_ref_hp, kosu_ref_kilo = kosu_ici_hp_medyani_tiklayarak_hesapla(
                        kullanilan_url,
                        row,
                        at_adi=at_adi
                    )

                    kayit["Koşu HP Medyan"] = kosu_ref_hp
                    kayit["Koşu HP Dolu Atlar Kilo Medyan"] = kosu_ref_kilo
                    kayit["Medyan HP Listesi"] = son_medyan_hp_liste
                    kayit["Medyan Kilo Listesi"] = son_medyan_kilo_liste
                    kayit["HP-Kilo Eşleşme Listesi"] = son_medyan_eslesme_liste
                    kayit["Koşu Detay URL"] = kosu_detay_url
                    kayit["Satır Durumu"] = "OK"
                    kayit["Hata"] = ""

                except Exception as satir_hatasi:
                    hata_tarih = satir_kolon_getir(row, "Tarih")
                    hata_sehir = satir_kolon_getir(row, "Şehir")
                    hata_msf = satir_kolon_getir(row, "Msf")
                    print(f"SATIR MEDYAN HATASI AMA KAYIT YAZILACAK: {at_adi} | {hata_tarih} | {hata_sehir} | {hata_msf} | {satir_hatasi}")

                    kayit["Satır Durumu"] = "MEDYAN_HATA_KAYIT_VAR"
                    kayit["Hata"] = str(satir_hatasi)

                    basarisiz_atlar.append({
                        "Şehir": sehir_adi,
                        "At ID": at_id,
                        "At Adı": at_adi,
                        "Tarih": hata_tarih,
                        "Koşu Şehir": hata_sehir,
                        "Mesafe": hata_msf,
                        "URL": bilgi.get("url", ""),
                        "Hata": str(satir_hatasi),
                        "Not": "Satır Veri sayfasına yazıldı; sadece medyan/link eksik olabilir."
                    })

                tum_sonuclar.append(kayit)

                if TEST_MODU and len(tum_sonuclar) >= TEST_BASARILI_SATIR_LIMIT:
                    print(f"TEST BASARILI SATIR LIMITI DOLDU: {len(tum_sonuclar)}")
                    break

        except Exception as e:

            print("AT HATASI:", at_adi)
            print(e)

            basarisiz_atlar.append({
                "Şehir": sehir_adi,
                "At ID": at_id,
                "At Adı": at_adi,
                "URL": bilgi.get("url", ""),
                "Hata": str(e)
            })

        # Her 100 sonuçta ara kayıt
        if len(tum_sonuclar) > 0 and len(tum_sonuclar) % 100 == 0:
            ara_kaydet(tum_sonuclar, basarisiz_atlar)

    # Her şehir sonunda kesin kayıt
    ara_kaydet(tum_sonuclar, basarisiz_atlar)

    if TEST_MODU and len(tum_sonuclar) >= TEST_BASARILI_SATIR_LIMIT:
        print(f"\nTEST MODU - Başarılı satır limiti doldu: {len(tum_sonuclar)}")
        break

    if TEST_MODU and test_toplam_at_deneme_sayaci >= TEST_TOPLAM_AT_DENEME_LIMIT:
        print(f"\nTEST MODU - Toplam at deneme limiti doldu: {test_toplam_at_deneme_sayaci}")
        break


# =========================================
# EXCEL FINAL
# =========================================

ara_kaydet(tum_sonuclar, basarisiz_atlar)


# =========================================
# SON
# =========================================

print("\n=================================")
print("TAMAMLANDI")
print("=================================")

print("\nToplam satır:", len(tum_sonuclar))
print("\nExcel:")
print(EXCEL_NAME)

print("\n=================================")
print("CACHE / HIZ RAPORU")
print("=================================")
print("Toplam detay ihtiyacı:", detay_cache_sayac["toplam_ihtiyac"])
print("Kimlik cache isabeti:", detay_cache_sayac["kimlik_cache_isabet"])
print("URL cache isabeti:", detay_cache_sayac["url_cache_isabet"])
print("Gerçek detay açılışı:", detay_cache_sayac["gercek_detay_acilis"])
print("Kimlik cache pas geçildi:", detay_cache_sayac["kimlik_cache_pas_gecildi"])
print("Şüpheli HP/kilo/kolon pas geçildi:", detay_cache_sayac["kilo_sutun_suspect_pas"])

if len(tum_sonuclar) == 0:
    print("\nUYARI: Excel boş çıktı.")

if len(basarisiz_atlar) > 0:
    print("\n=================================")
    print("UYARI: EKSİK AT VAR")
    print("=================================")
    print("\nEksik at sayısı:", len(basarisiz_atlar))
    print("Excel içinde 'Basarisiz_Atlar' sayfasına yazıldı.")
    print("Bu dosya TAM kabul edilmemeli. Başarısız atlar tekrar çalıştırılmalı.")
else:
    print("\nEksik at yok. Tüm atlar başarıyla işlendi.")

if not TEST_MODU:
    print("\nTest bitti; pencere kapatılıyor...")

try:
    driver.quit()
except:
    pass


# ============================================================
# ERTESI GUN MODU BİTTİ. ŞİMDİ ÜRETİLEN EXCEL ÜSTÜNE FAZ30 STİL EKLENİYOR.
# ============================================================

print("\nYARIN EXCEL TAMAMLANDI; FAZ30 STIL BASLIYOR...")

# -*- coding: utf-8 -*-
"""
TJK VIDEO TAKIP - FAZ 30 - ISTANBUL 1 KOSU HIZLI TEST

Bu sürüm MP4 indirme yapmaz.
TJK CDN indirme kopmasıyla uğraşmaz.

Akış:
1) Günlük programdan şehir bulur.
2) İlk şehirden ilk atı seçer.
3) At geçmişindeki ilk 4 video ikonuna tıklar.
4) Video sayfasında At No okur.
5) Satırdaki Derece değerini gerçek yarış süresi kabul eder.
6) Derece * oran hesabıyla doğrudan yarış süresinin %20/%45 noktalarına gider.
7) HTML5 video player seek edilir.
8) Video alanının ekran görüntüsü alınır.
9) Birden fazla crop bölgesi çıkarılır.
10) OCR denenir.
11) Rapor yazar.
12) Her frame için Excel satırı üretir.

Çıktı:
video_browser_faz30_tek45_stabil_13/
├── video_frames/
├── crops/
├── rapor.txt
└── sonuc.xlsx

Çalıştır:
cd ~/Desktop && python3 video_browser_faz30_tek45_stabil_13.py

Gerekli paketler:
pip3 install selenium webdriver-manager opencv-python pillow pytesseract

Excel için ekstra paket gerekmez. .xlsx dosyası Python stdlib ile yazılır.

Tesseract:
brew install tesseract
"""

import re
import time
import pandas as pd
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from zipfile import ZipFile, ZIP_DEFLATED

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


ANA_URL = "https://www.tjk.org/TR/yarissever/Info/Page/GunlukYarisProgrami"

TURK_SEHIRLER = [
    "İstanbul",
    "Bursa",
    "İzmir",
    "Kocaeli",
    "Ankara",
    "Adana",
    "Şanlıurfa",
    "Diyarbakır",
    "Elazığ",
    "Antalya",
]

HEADLESS_CHROME = False
SELENIUM_BEKLEME_SN = 2

CIKTI_KLASORU = Path("ertesi_gun_modu_faz30_stil_rapor")
FRAME_KLASORU = CIKTI_KLASORU / "video_frames"
CROP_KLASORU = CIKTI_KLASORU / "crops"

# ERTESI_GUN_MODU çıktısı. Script Desktop'ta çalıştırılacaksa aynı klasöre koy.
INPUT_EXCEL = "TUM_ILLER_TUM_ATLAR_EKSIKSIZ_ILK4_SON6AY_ERTESI_GUN.xlsx"
OUTPUT_EXCEL = "TUM_ILLER_TUM_ATLAR_EKSIKSIZ_ILK4_SON6AY_ERTESI_GUN_STIL_EKLI.xlsx"

# ============================================================
# STİL (VİDEO/%45) ÇEKİMİ AÇIK/KAPALI
#   False = video AÇILMAZ; çıktı aynı yapıda ama Stil sütunu BOŞ. ÇOK HIZLI.
#   True  = her at için video açılıp %45'te seyir sırası çıkarılır (YAVAŞ).
# Stil'i yeni yer'deki BY sütunu kullanır; kapalıyken BY boş kalır.
# ============================================================
STIL_CEK = True

# %45 stilini ortak kütüphaneden oku/yaz (derece & 800 paylaşır). Video en çok vakti
# yiyen kısım; bir kez çekilen %45 sonraki çalıştırmalarda tekrar açılmaz.
STIL_KUTUPHANE_KULLAN = True
try:
    import tjk_kutuphane as _STILKUT
except Exception as _skut_e:
    _STILKUT = None
    print(f"[STİL ÖNBELLEK] tjk_kutuphane yüklenemedi ({_skut_e}) -> video her seferinde açılır")


def faz30_stil_kimlik(row):
    """Derece çıktı excel satırından %45 önbellek anahtarı üretir.
    Çıktı excel'inde çim cushion sayısı (örn. '3.3') YOK; bu yüzden anahtar
    kum/sentetik'te son800/kütüphane ile birebir, çimde ise KENDİ İÇİNDE tutarlıdır
    (derece tekrar çalıştırmalarında aynı satır -> aynı anahtar -> video atlanır).
    Anahtar boşsa video normal yolla çekilir."""
    def _n(x):
        t = str(x).strip().upper()
        return re.sub(r"\s+", " ", t)
    def _g(*adlar):
        for a in adlar:
            for c in row.index:
                if str(c).strip().upper() == a.strip().upper():
                    return row[c]
        return ""
    tarih = _n(_g("Tarih")); sehir = _n(_g("Şehir"))
    mesafe = _n(_g("Mesafe", "Msf"))
    pist = _n(_g("Pist")); durum = _n(_g("Pist Durumu"))
    cins = _n(_g("Koşu Cinsi", "Kcins"))
    pist_comb = (pist[:1] + ":" + durum) if (pist and durum) else (pist[:1] if pist else "")
    z = [tarih, sehir, mesafe, pist_comb, cins]
    if any(x == "" or x == "NAN" for x in z):
        return ""
    return "|".join(z)

# Sadece İstanbul satırları işlenecek. 0 = filtre sonrası hepsi.
SADECE_SEHIR = ""
LIMIT_SATIR = 0
BASLANGIC_SATIR = 0
ARA_KAYIT_HER = 25

# HIZLI TEST: İstanbul içinden ilk koşu gibi davranmak için Excel sırasındaki ilk N benzersiz at işlenir.
# Eğer ilk koşuda at sayısı farklıysa bu sayıyı değiştir.
ILK_KOSU_AT_LIMIT = 0

MAX_VIDEO_SATIRI = 4

ORANLAR = [
    ("45", 0.45),
]

# Starttan önce video açılıyor. İlk testlerde 30 sn iyi çalıştı.
# Gerekirse 20 / 22 / 24 / 26 diye ayarlanabilir.
VIDEO_BASLANGIC_OFFSET_SN = 0.0

# Seek sonrası video görüntüsünün oturması için.
SEEK_SONRASI_BEKLE_SN = 5
SEEK_SONRASI_OYNAT_SN = 2
SEEK_SONRASI_DURDUR_BEKLE_SN = 1

OCR_DENE = True

# Disk dolmasın: frame/crop/debug PNG dosyaları kalıcı kaydedilmez.
# Video screenshot sadece geçici alınır, sıra okunduktan sonra silinir.
GORSEL_KAYDET = False

# Video elementinin tamamı kaydedilir.
# Sonra video görüntüsünden birden fazla crop kesilir.
CROP_DEBUG_BOLGELERI = [
    ("ALT_GENIS", 0.03, 0.70, 0.98, 0.96),
    ("ALT_ORTA", 0.03, 0.62, 0.98, 0.82),
    ("ALT_UST", 0.03, 0.52, 0.98, 0.72),
    ("SOL_ALT", 0.00, 0.60, 0.55, 0.95),
    ("ORTA_ALT", 0.20, 0.58, 0.85, 0.88),
    ("SAG_ALT", 0.45, 0.58, 1.00, 0.92),
]


def klasorleri_hazirla():
    CIKTI_KLASORU.mkdir(parents=True, exist_ok=True)
    FRAME_KLASORU.mkdir(parents=True, exist_ok=True)
    if GORSEL_KAYDET:
        CROP_KLASORU.mkdir(parents=True, exist_ok=True)


def driver_baslat():
    options = webdriver.ChromeOptions()

    if HEADLESS_CHROME:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
    else:
        options.add_argument("--start-maximized")

    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--remote-debugging-port=0")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-sync")
    options.add_argument("--disable-default-apps")
    options.page_load_strategy = "eager"

    d = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    d.set_page_load_timeout(60)
    return d


def driver_canli_mi(driver):
    try:
        _ = driver.current_url
        return True
    except Exception:
        return False


def guvenli_get(driver, url, bekleme=SELENIUM_BEKLEME_SN):
    son_hata = None

    for deneme in range(1, 6):
        try:
            if driver is None or not driver_canli_mi(driver):
                print("CHROME DRIVER ÖLÜ - YENİDEN AÇILIYOR...")
                try:
                    driver.quit()
                except Exception:
                    pass
                driver = driver_baslat()
                time.sleep(2)

            driver.get(url)
            time.sleep(bekleme)
            return driver

        except Exception as e:
            son_hata = e
            print(f"CHROME/GET HATASI ({deneme}/5) - DRIVER YENİDEN AÇILIYOR - HATA: {e}")

            try:
                driver.quit()
            except Exception:
                pass

            time.sleep(3)
            driver = driver_baslat()

    raise Exception(f"Sayfa açılamadı: {url} | Son hata: {son_hata}")


def derece_to_seconds(value):
    t = str(value or "").strip()
    if not t:
        return 0.0

    t = t.replace(",", ".")
    t = re.sub(r"[^0-9.]", "", t)

    if not t:
        return 0.0

    parts = t.split(".")

    try:
        if len(parts) >= 3:
            dakika = int(parts[0])
            saniye = int(parts[1])
            salise = int(parts[2])
            return dakika * 60 + saniye + (salise / (10 ** len(parts[2])))

        if len(parts) == 2:
            return float(t)

        return float(t)
    except Exception:
        return 0.0


def satirdan_derece_cek(satir_text):
    """Satır içinden dereceyi sağlam yakalar.

    Eski sürüm 131.10 gibi 3 haneli saniyeleri kaçırıyordu; bu doğrudan
    HEDEF_SATIR_YOK üretir. TJK tarafında derece bazen 2.11.10, bazen 131.10,
    bazen 2:11.10, bazen de 13110 gibi düz gelebilir.
    """
    txt = str(satir_text or "")
    txt = txt.replace(",", ".")

    # 2.11.10 / 2:11.10
    m = re.search(r"\b([0-9]{1,2})[:.]([0-5]?\d)\.(\d{1,2})\b", txt)
    if m:
        return f"{m.group(1)}.{int(m.group(2)):02d}.{m.group(3)}"

    # 94.96 / 131.10 / 143.25 gibi saniye.salise formatı.
    adaylar = re.findall(r"\b(\d{2,3}\.\d{1,2})\b", txt)
    if adaylar:
        try:
            return sorted(adaylar, key=lambda x: float(x), reverse=True)[0]
        except Exception:
            return adaylar[0]

    # 9496 / 13110 gibi çıplak derece. 4-5 haneli sayı yarış derecesi adayıdır.
    adaylar = re.findall(r"\b(\d{4,5})\b", re.sub(r"[^0-9 ]", " ", txt))
    if adaylar:
        try:
            n = sorted([int(x) for x in adaylar], reverse=True)[0]
            return f"{n / 100.0:.2f}"
        except Exception:
            pass

    return ""


def at_no_oku(driver):
    metinler = []

    try:
        metinler.append(driver.find_element(By.TAG_NAME, "body").text or "")
    except Exception:
        pass

    try:
        metinler.append(driver.page_source or "")
    except Exception:
        pass

    tum = "\n".join(metinler)

    kaliplar = [
        r"At\s*No\s*[:：]\s*(\d+)",
        r"At\s*No\s+(\d+)",
        r"AtNo\s*[:：]?\s*(\d+)",
    ]

    for p in kaliplar:
        m = re.search(p, tum, flags=re.IGNORECASE)
        if m:
            return m.group(1)

    return ""


def mp4_url_oku(driver):
    js = """
    const v = document.querySelector('video');
    if (!v) return '';
    return v.currentSrc || v.src || '';
    """
    try:
        src = driver.execute_script(js)
        if src and ".mp4" in src:
            return src
    except Exception:
        pass

    try:
        html = driver.page_source or ""
    except Exception:
        html = ""

    m = re.search(r"https://video-cdn\.tjk\.org/videoftp/[^\"'\s<>]+\.mp4", html)
    if m:
        return m.group(0)

    return ""


def video_sayfasi_mi(driver):
    """
    Açılan sayfanın gerçekten video sayfası olup olmadığını kontrol eder.
    AtKosuBilgileri gibi yanlış sayfalar burada elenir.
    """
    try:
        url = driver.current_url or ""
    except Exception:
        url = ""

    if "YarisVideo" in url or "YarisVideoAt" in url:
        return True

    try:
        has_video = driver.execute_script("return !!document.querySelector('video');")
        if has_video:
            return True
    except Exception:
        pass

    try:
        src = mp4_url_oku(driver)
        if src and ".mp4" in src:
            return True
    except Exception:
        pass

    return False


def video_mevcut_mu(driver):
    """
    Sayfada HTML5 video gerçekten var mı?
    """
    try:
        return bool(driver.execute_script("return !!document.querySelector('video');"))
    except Exception:
        return False



def sehir_linklerini_bul(driver):
    driver = guvenli_get(driver, ANA_URL, bekleme=5)
    print("Günlük program açıldı.")
    time.sleep(3)

    links = driver.find_elements(By.TAG_NAME, "a")
    sehir_linkleri = {}

    for link in links:
        try:
            href = link.get_attribute("href")
            text = link.text.strip()

            if not href:
                continue

            if "GunlukYarisProgrami?SehirId=" not in href:
                continue

            for sehir in TURK_SEHIRLER:
                if sehir.lower() in text.lower():
                    if sehir not in sehir_linkleri:
                        sehir_linkleri[sehir] = href.split("&Era=")[0]
        except Exception:
            pass

    print("Şehir sayısı:", len(sehir_linkleri))
    for s, u in sehir_linkleri.items():
        print(f"- {s}: {u}")

    return sehir_linkleri


def at_linklerini_bul(driver, sehir_url):
    driver = guvenli_get(driver, sehir_url, bekleme=5)
    print("Şehir sayfası açıldı:")
    print(driver.current_url)
    time.sleep(3)

    links = driver.find_elements(By.TAG_NAME, "a")
    at_linkleri = {}

    for link in links:
        try:
            href = link.get_attribute("href")
            text = link.text.strip()

            if not href:
                continue

            if "QueryParameter_AtId=" not in href:
                continue

            parsed = urlparse(href)
            params = parse_qs(parsed.query)

            if "QueryParameter_AtId" not in params:
                continue

            at_id = params["QueryParameter_AtId"][0]
            at_adi = text.strip() or f"AT_{at_id}"

            if at_id not in at_linkleri:
                at_linkleri[at_id] = {"At ID": at_id, "At Adı": at_adi, "URL": href}
        except Exception:
            pass

    print("At link sayısı:", len(at_linkleri))
    for i, a in enumerate(list(at_linkleri.values())[:10], start=1):
        print(f"[{i}] {a['At Adı']} | {a['At ID']} | {a['URL']}")

    return list(at_linkleri.values())


def tarih_satirlarini_bul(driver):
    """At geçmişindeki gerçek satır adaylarını döndürür.

    Sadece görünen text içinde tarih aramak tırt: bazı TJK satırlarında tarih/link
    outerHTML içinde duruyor veya video/sonuç butonu ayrı hücrede görünüyor.
    Bu yüzden tarih + video/sonuç/link sinyali olan satırlar da aday yapılır.
    """
    satirlar = driver.find_elements(By.TAG_NAME, "tr")
    tarih_satirlari = []
    seen = set()

    for tr in satirlar:
        try:
            txt = " ".join((tr.text or "").split())
        except Exception:
            txt = ""
        try:
            html = tr.get_attribute("outerHTML") or ""
        except Exception:
            html = ""

        birlesik = f"{txt} {html}"
        b_up = birlesik.upper()

        tarih_var = bool(re.search(r"\b\d{1,2}[./]\d{1,2}[./]\d{4}\b", birlesik))
        sonuc_video_var = (
            "GUNLUKYARISSONUCLARI" in b_up
            or "YARISSONUCLARI" in b_up
            or "YARISVIDEO" in b_up
            or "VIDEO" in b_up
            or "İZLE" in b_up
            or "IZLE" in b_up
            or "PLAY" in b_up
        )

        if tarih_var or sonuc_video_var:
            try:
                key = tr.id
            except Exception:
                key = id(tr)
            if key not in seen:
                tarih_satirlari.append(tr)
                seen.add(key)

    return tarih_satirlari


def satir_text(tr):
    try:
        return " ".join((tr.text or "").split())
    except Exception:
        return ""


def satirdaki_video_butonunu_bul(tr):
    adaylar = []

    try:
        els = tr.find_elements(By.XPATH, ".//*")
    except Exception:
        els = []

    for el in els:
        try:
            tag = el.tag_name.lower()
            text = " ".join((el.text or "").split())
            href = (el.get_attribute("href") or "").strip()
            src = (el.get_attribute("src") or "").strip()
            title = (el.get_attribute("title") or "").strip()
            alt = (el.get_attribute("alt") or "").strip()
            cls = (el.get_attribute("class") or "").strip()
            onclick = (el.get_attribute("onclick") or "").strip()
            outer = (el.get_attribute("outerHTML") or "").strip()

            birlesik = " ".join([tag, text, href, src, title, alt, cls, onclick, outer]).lower()

            sinyal = (
                "video" in birlesik
                or "play" in birlesik
                or "izle" in birlesik
                or "kamera" in birlesik
                or "fa-play" in birlesik
                or "camera" in birlesik
            )

            if sinyal:
                loc = el.location_once_scrolled_into_view
                size = el.size
                x = loc.get("x", 0)
                w = size.get("width", 0)
                h = size.get("height", 0)

                if w >= 1 and h >= 1:
                    adaylar.append((x, el))
        except Exception:
            pass

    if adaylar:
        adaylar = sorted(adaylar, key=lambda x: x[0], reverse=True)
        return adaylar[0][1], "VIDEO_SINYALLI"

    tiklanabilirler = []

    for xpath in [".//a", ".//button", ".//img", ".//span", ".//i"]:
        try:
            for el in tr.find_elements(By.XPATH, xpath):
                try:
                    loc = el.location_once_scrolled_into_view
                    size = el.size
                    x = loc.get("x", 0)
                    w = size.get("width", 0)
                    h = size.get("height", 0)

                    if w >= 1 and h >= 1:
                        tiklanabilirler.append((x, el))
                except Exception:
                    pass
        except Exception:
            pass

    # ESKİ FALLBACK TIRT:
    # "EN_SAG_TIKLANABILIR" bazen video yerine AtKosuBilgileri linkine tıklıyor.
    # Bu da video_info=None üretip Excel'e sahte seyir sırası yazdırıyordu.
    # Artık sinyal yoksa satır atlanır.
    return None, "VIDEO_SINYALI_YOK_ATLANDI"


def satirdaki_video_url_bul(tr, base_url=""):
    """Satırdaki video linkini doğrudan URL olarak yakalar."""
    try:
        html = tr.get_attribute("outerHTML") or ""
    except Exception:
        html = ""

    adaylar = []
    for m in re.finditer(r'(?:href|src)=["\']([^"\']+)["\']', html, flags=re.IGNORECASE):
        u = m.group(1).strip()
        if not u or u.startswith("javascript"):
            continue
        up = u.upper()
        if "YARISVIDEO" in up or "VIDEO" in up or ".MP4" in up:
            adaylar.append(urljoin(base_url or "https://www.tjk.org/", u))

    for m in re.finditer(r'["\']([^"\']*(?:YarisVideo|YARISVIDEO|video|VIDEO)[^"\']*)["\']', html):
        u = m.group(1).strip()
        if u and not u.lower().startswith("javascript"):
            adaylar.append(urljoin(base_url or "https://www.tjk.org/", u))

    temiz = []
    for u in adaylar:
        if u not in temiz:
            temiz.append(u)
    return temiz[0] if temiz else ""

def elemente_tikla(driver, el):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", el)
        time.sleep(0.5)
    except Exception:
        pass

    try:
        el.click()
        return True, "NORMAL_CLICK"
    except Exception as e1:
        try:
            driver.execute_script("arguments[0].click();", el)
            return True, f"JS_CLICK | normal hata={e1}"
        except Exception as e2:
            return False, f"CLICK_HATA normal={e1} js={e2}"


def yeni_pencereye_gec(driver, onceki_handles):
    time.sleep(2)
    handles = driver.window_handles
    yeni = [h for h in handles if h not in onceki_handles]

    if yeni:
        driver.switch_to.window(yeni[-1])
        time.sleep(3)
        return True

    return False


def video_kayitlarini_topla(driver, at):
    driver = guvenli_get(driver, at["URL"], bekleme=5)
    print("At geçmiş sayfası açıldı:")
    print(driver.current_url)
    time.sleep(5)

    tarih_satirlari = tarih_satirlarini_bul(driver)
    print("Tarih satırı sayısı:", len(tarih_satirlari))

    ana_handle = driver.current_window_handle
    ana_url = driver.current_url

    video_kayitlari = []

    for idx, tr in enumerate(tarih_satirlari[:MAX_VIDEO_SATIRI], start=1):
        txt = satir_text(tr)
        print("")
        print(f"VIDEO SATIR {idx}: {txt}")

        derece = satirdan_derece_cek(txt)
        derece_sn = derece_to_seconds(derece)

        el, kaynak = satirdaki_video_butonunu_bul(tr)
        print("Buton kaynak:", kaynak)

        if el is None:
            print("Video butonu bulunamadı.")
            continue

        onceki_handles = set(driver.window_handles)

        ok, click_notu = elemente_tikla(driver, el)
        print("Click:", ok, click_notu)

        if not ok:
            continue

        yeni_pencere = yeni_pencereye_gec(driver, onceki_handles)
        time.sleep(5)

        video_url = driver.current_url
        at_no = at_no_oku(driver)
        mp4_url = mp4_url_oku(driver)

        print("Video URL:", video_url)
        print("At No:", at_no)
        print("Derece:", derece, "=>", derece_sn)
        print("MP4 URL:", mp4_url)

        video_kayitlari.append({
            "idx": idx,
            "satir": txt,
            "derece": derece,
            "derece_sn": derece_sn,
            "video_url": video_url,
            "at_no": at_no,
            "mp4_url": mp4_url,
            "yeni_pencere": yeni_pencere,
        })

        if yeni_pencere:
            try:
                driver.close()
            except Exception:
                pass
            try:
                driver.switch_to.window(ana_handle)
                time.sleep(1)
            except Exception:
                pass
        else:
            driver = guvenli_get(driver, ana_url, bekleme=3)
            time.sleep(2)
            tarih_satirlari = tarih_satirlarini_bul(driver)

    return video_kayitlari


def video_play(driver):
    js = """
    const v = document.querySelector('video');
    if (!v) return false;
    v.muted = true;
    const p = v.play();
    return true;
    """
    try:
        return bool(driver.execute_script(js))
    except Exception:
        return False


def video_pause(driver):
    js = """
    const v = document.querySelector('video');
    if (!v) return false;
    v.pause();
    return true;
    """
    try:
        return bool(driver.execute_script(js))
    except Exception:
        return False


def video_seek(driver, saniye):
    js = """
    const v = document.querySelector('video');
    if (!v) return false;
    v.pause();
    v.currentTime = arguments[0];
    return true;
    """
    try:
        return bool(driver.execute_script(js, float(saniye)))
    except Exception:
        return False


def video_seek_oturana_kadar_bekle(driver, hedef_saniye, max_bekle=18, tolerans=0.75):
    """
    Kritik düzeltme:
    seek=True sadece komutun verildiğini söyler.
    Görüntünün gerçekten yeni saniyeye oturması için currentTime + readyState kontrol edilir.
    """
    hedef_saniye = float(hedef_saniye)

    for i in range(max_bekle):
        info = video_info(driver)
        print(f"SEEK OTURMA KONTROL {i+1}/{max_bekle}: hedef={hedef_saniye:.2f} info={info}")

        if info:
            try:
                cur = float(info.get("currentTime", 0) or 0)
                ready = int(info.get("readyState", 0) or 0)

                if abs(cur - hedef_saniye) <= tolerans and ready >= 2:
                    time.sleep(1.0)
                    return True, info

                # Bazen currentTime hedefe gider ama readyState 1'de kalır.
                # Bu durumda kısa play/pause görüntüyü zorla yeniler.
                if abs(cur - hedef_saniye) <= tolerans and ready < 2:
                    video_play(driver)
                    time.sleep(0.8)
                    video_pause(driver)

            except Exception:
                pass

        time.sleep(1)

    return False, video_info(driver)


def video_info(driver):
    js = """
    const v = document.querySelector('video');
    if (!v) return null;
    return {
        currentTime: v.currentTime || 0,
        duration: v.duration || 0,
        readyState: v.readyState || 0,
        paused: v.paused,
        src: v.currentSrc || v.src || ''
    };
    """
    try:
        return driver.execute_script(js)
    except Exception:
        return None


def video_hazir_bekle(driver, max_bekle=20):
    for i in range(max_bekle):
        info = video_info(driver)
        print(f"VIDEO INFO {i+1}/{max_bekle}: {info}")

        if info:
            try:
                dur = float(info.get("duration", 0) or 0)
                ready = int(info.get("readyState", 0) or 0)
                if dur > 0 and ready >= 2:
                    return True
            except Exception:
                pass

        video_play(driver)
        time.sleep(1)

    return False


def video_element_screenshot(driver, out_path):
    """
    Sadece video player alanını alır.
    Önce #rmpPlayer, olmazsa video elementi.
    """
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    el = None

    try:
        el = driver.find_element(By.ID, "rmpPlayer")
    except Exception:
        pass

    if el is None:
        try:
            el = driver.find_element(By.TAG_NAME, "video")
        except Exception:
            pass

    if el is None:
        driver.save_screenshot(str(out_path))
        return out_path, "FULL_SCREEN_FALLBACK"

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", el)
        time.sleep(0.5)
    except Exception:
        pass

    el.screenshot(str(out_path))
    return out_path, "VIDEO_ELEMENT"


def crop_debug_bolgeleri_uret(frame_path, idx, etiket, at_no):
    if not GORSEL_KAYDET:
        return []
    import cv2

    img = cv2.imread(str(frame_path))
    if img is None:
        raise Exception(f"Frame okunamadı crop için: {frame_path}")

    h, w = img.shape[:2]
    ciktilar = []

    for ad, rx1, ry1, rx2, ry2 in CROP_DEBUG_BOLGELERI:
        x1 = max(0, min(w - 1, int(w * rx1)))
        x2 = max(1, min(w, int(w * rx2)))
        y1 = max(0, min(h - 1, int(h * ry1)))
        y2 = max(1, min(h, int(h * ry2)))

        crop = img[y1:y2, x1:x2]

        base = CROP_KLASORU / f"video_{idx}_{etiket}_atno_{at_no}_{ad}"
        normal_path = Path(str(base) + ".png")
        prep_path = Path(str(base) + "_PREP.png")

        normal_path.parent.mkdir(parents=True, exist_ok=True)
        cv2.imwrite(str(normal_path), crop)

        try:
            gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
            gray = cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
            gray = cv2.GaussianBlur(gray, (3, 3), 0)
            _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            cv2.imwrite(str(prep_path), th)
        except Exception:
            prep_path = normal_path

        ciktilar.append((ad, normal_path, prep_path))

    return ciktilar


def ocr_dene_gelismis(image_path):
    if not OCR_DENE:
        return ""

    try:
        import pytesseract
        from PIL import Image
    except Exception as e:
        return f"OCR_PAKET_YOK: {e}"

    try:
        img = Image.open(image_path)

        configs = [
            "--psm 6 -c tessedit_char_whitelist=0123456789",
            "--psm 7 -c tessedit_char_whitelist=0123456789",
            "--psm 11 -c tessedit_char_whitelist=0123456789",
            "--psm 13 -c tessedit_char_whitelist=0123456789",
        ]

        sonuclar = []

        for cfg in configs:
            try:
                text = pytesseract.image_to_string(img, config=cfg)
                text = re.sub(r"\s+", " ", text).strip()
                if text:
                    sonuclar.append(text)
            except Exception as e:
                sonuclar.append(f"HATA:{e}")

        temiz = []

        for s in sonuclar:
            if s not in temiz:
                temiz.append(s)

        return " || ".join(temiz)

    except Exception as e:
        return f"OCR_HATA: {e}"



def satir_bilgileri_ayikla(satir_text):
    txt = " ".join(str(satir_text or "").split())
    tarih = ""
    sehir = ""
    mesafe = ""
    m = re.search(r"\b(\d{1,2}\.\d{1,2}\.\d{4})\s+([^\s]+)\s+(\d{3,4})\b", txt)
    if m:
        tarih = m.group(1)
        sehir = m.group(2)
        mesafe = m.group(3)
    else:
        m_tarih = re.search(r"\b(\d{1,2}\.\d{1,2}\.\d{4})\b", txt)
        if m_tarih:
            tarih = m_tarih.group(1)
    derece = satirdan_derece_cek(txt)
    return {"tarih": tarih, "sehir": sehir, "mesafe": mesafe, "derece": derece, "satir": txt}


def excel_xml_escape(value):
    s = "" if value is None else str(value)
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def excel_col_name(n):
    name = ""
    while n:
        n, rem = divmod(n - 1, 26)
        name = chr(65 + rem) + name
    return name


def xlsx_hucre(ref, value, style_id=None):
    style = f' s="{style_id}"' if style_id is not None else ""
    if value is None:
        return f'<c r="{ref}"{style}/>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"{style}><v>{value}</v></c>'
    return f'<c r="{ref}" t="inlineStr"{style}><is><t>{excel_xml_escape(value)}</t></is></c>'


def xlsx_yaz(rows, out_path):
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    headers = [
        "Video No", "Tarih", "Şehir", "Mesafe", "At No", "Derece", "Derece Sn",
        "Oran", "Hedef Saniye", "Seek Oturdu", "Seyir Sırası", "Sıra Kaynağı",
        "Frame Dosyası", "Video URL", "MP4 URL", "Satır",
    ]
    sheet_rows = []
    header_cells = [xlsx_hucre(f"{excel_col_name(i)}1", h, style_id=1) for i, h in enumerate(headers, start=1)]
    sheet_rows.append('<row r="1">' + ''.join(header_cells) + '</row>')
    for r_idx, row in enumerate(rows, start=2):
        cells = []
        for c_idx, h in enumerate(headers, start=1):
            cells.append(xlsx_hucre(f"{excel_col_name(c_idx)}{r_idx}", row.get(h, ""), style_id=0))
        sheet_rows.append(f'<row r="{r_idx}">' + ''.join(cells) + '</row>')
    max_row = max(1, len(rows) + 1)
    max_col = len(headers)
    dimension = f"A1:{excel_col_name(max_col)}{max_row}"
    widths = {1:10,2:13,3:13,4:10,5:9,6:12,7:11,8:8,9:13,10:12,11:12,12:40,13:36,14:55,15:55,16:90}
    cols_xml = "<cols>" + "".join(f'<col min="{i}" max="{i}" width="{widths.get(i,14)}" customWidth="1"/>' for i in range(1, max_col + 1)) + "</cols>"
    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dimension}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  {cols_xml}
  <sheetData>{''.join(sheet_rows)}</sheetData>
  <autoFilter ref="{dimension}"/>
</worksheet>'''
    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD9E2F3"/></left><right style="thin"><color rgb="FFD9E2F3"/></right><top style="thin"><color rgb="FFD9E2F3"/></top><bottom style="thin"><color rgb="FFD9E2F3"/></bottom><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''
    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sonuc" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'''
    with ZipFile(out_path, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        z.writestr("xl/styles.xml", styles_xml)
    return out_path


def video_kayitlarini_topla_ve_aninda_isle(driver, at):
    """
    Hızlandırılmış akış:
    Eski sürüm önce video URL'lerini topluyordu, sonra aynı videoları tekrar açıp foto çekiyordu.
    Bu sürüm video ikonuna tıkladığı anda aynı açık video sayfasında %20/%45 foto çeker.
    Böylece her video gereksiz yere ikinci kez açılmaz.
    """
    driver = guvenli_get(driver, at["URL"], bekleme=5)
    print("At geçmiş sayfası açıldı:")
    print(driver.current_url)
    time.sleep(5)

    tarih_satirlari = tarih_satirlarini_bul(driver)
    print("Tarih satırı sayısı:", len(tarih_satirlari))

    ana_handle = driver.current_window_handle
    ana_url = driver.current_url

    raporlar = []
    excel_rows = []
    islenen_video_sayisi = 0

    for idx in range(1, min(MAX_VIDEO_SATIRI, len(tarih_satirlari)) + 1):
        # Sayfa geri yüklenirse stale element yememek için satırları her tur tazele.
        tarih_satirlari = tarih_satirlarini_bul(driver)
        if idx > len(tarih_satirlari):
            break

        tr = tarih_satirlari[idx - 1]
        txt = satir_text(tr)
        print("")
        print(f"VIDEO SATIR {idx}: {txt}")

        derece = satirdan_derece_cek(txt)
        derece_sn = derece_to_seconds(derece)

        el, kaynak = satirdaki_video_butonunu_bul(tr)
        print("Buton kaynak:", kaynak)

        if el is None:
            print("Video butonu bulunamadı.")
            continue

        onceki_handles = set(driver.window_handles)

        ok, click_notu = elemente_tikla(driver, el)
        print("Click:", ok, click_notu)

        if not ok:
            continue

        yeni_pencere = yeni_pencereye_gec(driver, onceki_handles)
        time.sleep(5)

        video_url = driver.current_url

        # Kritik kontrol: yanlışlıkla AtKosuBilgileri gibi video olmayan sayfaya gidildiyse satırı atla.
        if not video_sayfasi_mi(driver):
            print("UYARI: Açılan sayfa video değil, satır atlandı:", video_url)

            if yeni_pencere:
                try:
                    driver.close()
                except Exception:
                    pass
                try:
                    driver.switch_to.window(ana_handle)
                    time.sleep(1)
                except Exception:
                    pass
            else:
                driver = guvenli_get(driver, ana_url, bekleme=3)
                time.sleep(2)

            continue

        at_no = at_no_oku(driver)
        mp4_url = mp4_url_oku(driver)

        print("Video URL:", video_url)
        print("At No:", at_no)
        print("Derece:", derece, "=>", derece_sn)
        print("MP4 URL:", mp4_url)

        if not at_no:
            print("UYARI: At No okunamadı, satır atlandı.")
            if yeni_pencere:
                try:
                    driver.close()
                except Exception:
                    pass
                try:
                    driver.switch_to.window(ana_handle)
                    time.sleep(1)
                except Exception:
                    pass
            else:
                driver = guvenli_get(driver, ana_url, bekleme=3)
                time.sleep(2)
            continue

        kayit = {
            "idx": idx,
            "satir": txt,
            "derece": derece,
            "derece_sn": derece_sn,
            "video_url": video_url,
            "at_no": at_no,
            "mp4_url": mp4_url,
            "yeni_pencere": yeni_pencere,
        }

        try:
            rapor, satir_excel_rows = browser_video_isle(driver, kayit, acik_sayfayi_kullan=True)
            raporlar.append(rapor)
            excel_rows.extend(satir_excel_rows)
            islenen_video_sayisi += 1
            print(rapor)
        except Exception as e:
            hata = f"VIDEO {idx} GENEL HATA: {e}"
            raporlar.append(hata)
            print(hata)

        if yeni_pencere:
            try:
                driver.close()
            except Exception:
                pass
            try:
                driver.switch_to.window(ana_handle)
                time.sleep(1)
            except Exception:
                pass
        else:
            driver = guvenli_get(driver, ana_url, bekleme=3)
            time.sleep(2)

    print("")
    print("TOPLAM İŞLENEN VIDEO:", islenen_video_sayisi)
    return raporlar, excel_rows



def _hue_dist(a, b):
    d = abs(float(a) - float(b))
    return min(d, 180.0 - d)


def _slot_renkten_no(roi):
    """
    Şerit kutusunu renginden at numarasına çevirir.
    Bu sürümün ana mantığı budur; OCR'ye abanmaz.

    Desteklenen standart renkler:
    1 kırmızı, 2 mavi, 3 turuncu, 4 yeşil, 5 siyah,
    6 mor, 7 camgöbeği, 8 kahve, 9 haki/sarı,
    10 lila, 11 koyu yeşil, 12 pembe, 14 zeytin.
    """
    try:
        import cv2
        import numpy as np
    except Exception:
        return "", "CV_YOK"

    if roi is None or roi.size == 0:
        return "", "ROI_BOS"

    hsv = cv2.cvtColor(roi, cv2.COLOR_BGR2HSV)
    h, w = hsv.shape[:2]

    # Merkez kumaş rengini al; beyaz rakam ve kenarı azalt.
    core = hsv[int(h * 0.20):int(h * 0.80), int(w * 0.14):int(w * 0.86)]
    if core.size == 0:
        core = hsv

    H = core[:, :, 0].astype("float32")
    S = core[:, :, 1].astype("float32")
    V = core[:, :, 2].astype("float32")

    dark_ratio = float((V < 78).mean())
    gray_ratio = float(((S < 75) & (V > 90) & (V < 215)).mean())

    if dark_ratio > 0.38:
        return "5", f"DARK={dark_ratio:.2f}"

    colored = (S > 45) & (V > 35)

    if colored.sum() < 25:
        # Açık gri/beyaz slotlar çoğu zaman 13 numara.
        bright_gray_ratio = float(((S < 80) & (V >= 150)).mean())
        if bright_gray_ratio > 0.35:
            return "13", f"BRIGHT_GRAY_13={bright_gray_ratio:.2f}"
        if gray_ratio > 0.45:
            return "10", f"GRAY={gray_ratio:.2f}"
        return "", f"RENK_AZ gray={gray_ratio:.2f}"

    hm = float(np.median(H[colored]))
    sm = float(np.median(S[colored]))
    vm = float(np.median(V[colored]))

    prototypes = {
        "1":  (177, 205, 200),
        "2":  (116, 235, 160),
        "3":  (15,  220, 235),
        "4":  (59,  165, 155),
        "6":  (138, 210, 150),
        "7":  (101, 215, 235),
        "8":  (7,   145, 82),
        "9":  (22,  145, 192),
        "10": (122, 115, 205),
        "11": (62,  122, 88),
        "12": (168, 195, 184),
        "13": (0,   35,  200),
        "14": (32,  165, 130),
        "17": (160, 155, 185),
    }

    best_no = ""
    best_score = 9999.0

    for no, (ph, ps, pv) in prototypes.items():
        score = _hue_dist(hm, ph) + abs(sm - ps) / 9.0 + abs(vm - pv) / 9.0
        if score < best_score:
            best_score = score
            best_no = no

    if best_score <= 35:
        return best_no, f"HSV={hm:.1f},{sm:.1f},{vm:.1f}_SCORE={best_score:.1f}"

    return "", f"ESIK_DISI_HSV={hm:.1f},{sm:.1f},{vm:.1f}_SCORE={best_score:.1f}"


def _slot_ocr_destek(roi, debug_dir, prefix):
    """
    Renk kararsızsa destek OCR.
    Ana yöntem değil.
    """
    try:
        import cv2
        import numpy as np
        import pytesseract
        from PIL import Image
    except Exception:
        return ""

    if roi is None or roi.size == 0:
        return ""

    hsv = cv2.cvtColor(roi, cv2.COLOR_BGR2HSV)

    masks = [
        cv2.inRange(hsv, (0, 0, 105), (179, 125, 255)),
        cv2.inRange(hsv, (0, 0, 0), (179, 255, 95)),
    ]

    found = []

    for mi, m in enumerate(masks, start=1):
        m = cv2.dilate(m, np.ones((2, 2), np.uint8), iterations=1)
        big = cv2.resize(m, None, fx=5, fy=5, interpolation=cv2.INTER_CUBIC)
        big = cv2.copyMakeBorder(big, 15, 15, 15, 15, cv2.BORDER_CONSTANT, value=0)

        if GORSEL_KAYDET:
            try:
                cv2.imwrite(str(debug_dir / f"{prefix}_ocrmask{mi}.png"), big)
            except Exception:
                pass

        for psm in [10, 8, 7]:
            try:
                txt = pytesseract.image_to_string(
                    Image.fromarray(big),
                    config=f"--psm {psm} -c tessedit_char_whitelist=0123456789"
                )
                tok = re.sub(r"\D", "", txt or "")
                if tok:
                    found.append(tok)
            except Exception:
                pass

    for tok in found:
        try:
            n = int(tok)
            if 1 <= n <= 20:
                return str(n)
        except Exception:
            pass

    for tok in found:
        for suf in ["10", "11", "12", "13", "14", "17"]:
            if tok.endswith(suf) or tok.startswith(suf):
                return suf
        if len(tok) >= 2 and len(set(tok)) == 1 and tok[0] != "0":
            return tok[0]

    return ""


def _sira_kutularini_bul(img, debug_dir, prefix):
    """
    Kutuları doğrudan tüm frame alt bandından bulur.
    Kritik düzeltme: cv2.RETR_LIST kullanır.
    Böylece büyük şerit/arka plan konturu altında kalan kutular da alınır.
    """
    import cv2
    import numpy as np

    h, w = img.shape[:2]

    # Kullanıcı görsellerinde güvenilir alan:
    # Şerit daima alt bölümde, kontrol barının üstünde.
    x1, y1 = int(w * 0.17), int(h * 0.70)
    x2, y2 = int(w * 0.95), int(h * 0.84)

    band = img[y1:y2, x1:x2].copy()
    if GORSEL_KAYDET:
        debug_dir.mkdir(parents=True, exist_ok=True)
    band_path = debug_dir / f"{prefix}_SIRA_BANDI.png"
    if GORSEL_KAYDET:
        cv2.imwrite(str(band_path), band)

    bh, bw = band.shape[:2]
    hsv = cv2.cvtColor(band, cv2.COLOR_BGR2HSV)

    # Renkli + siyah kutular. RETR_LIST ile iç konturlar da yakalanır.
    mask = (((hsv[:, :, 1] > 50) & (hsv[:, :, 2] > 30)) | (hsv[:, :, 2] < 90)).astype("uint8") * 255
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8), iterations=1)

    mask_path = debug_dir / f"{prefix}_SIRA_MASK.png"
    if GORSEL_KAYDET:
        cv2.imwrite(str(mask_path), mask)

    contours, _ = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

    boxes = []

    for c in contours:
        x, y, ww, hh = cv2.boundingRect(c)
        area = ww * hh
        aspect = ww / max(1, hh)

        if area < 500 or area > 6000:
            continue
        if ww < bw * 0.025 or ww > bw * 0.12:
            continue
        if hh < bh * 0.18 or hh > bh * 0.60:
            continue
        if aspect < 0.70 or aspect > 3.00:
            continue

        # Soldaki sarı mesafe kapsülü.
        if x < bw * 0.09 and ww > bw * 0.06:
            continue

        boxes.append((x, y, ww, hh))

    boxes = sorted(boxes, key=lambda b: b[0])

    # Duplicate temizliği.
    dedup = []
    for b in boxes:
        if not any(abs(b[0] - d[0]) < 6 and abs(b[2] - d[2]) < 12 for d in dedup):
            dedup.append(b)
    boxes = dedup

    if not boxes:
        return band, [], str(band_path)

    # Kutular 5 veya 10 olabilir. Kaçan slotları geometriyle tamamla.
    xs = [b[0] for b in boxes]
    ws = [b[2] for b in boxes]
    hs = [b[3] for b in boxes]
    ys = [b[1] for b in boxes]

    box_w = int(np.median(ws))
    box_h = int(np.median(hs))
    box_y = int(np.median(ys))

    diffs = [xs[i + 1] - xs[i] for i in range(len(xs) - 1)]
    good = [d for d in diffs if box_w * 0.90 <= d <= box_w * 1.75]
    step = int(np.median(good)) if good else int(box_w * 1.25)

    # İlk slotu geri doldur. Örneğin 10'lu şeritte ilk 11/12 bazen kaçabiliyor.
    start_x = boxes[0][0]
    while start_x - step > bw * 0.08:
        start_x -= step

    # Önce kutular arası boşlukları doldur.
    filled = []
    cur_x = start_x

    # 5 mi 10 mu?
    # Eğer algılanan kutu sayısı 6+ ise veya son kutu sağlara uzanıyorsa 10 kabul et.
    likely_10 = len(boxes) >= 6 or boxes[-1][0] > bw * 0.55

    max_slots = 10 if likely_10 else len(boxes)

    for k in range(max_slots):
        sx = start_x + k * step
        if sx + box_w * 0.65 > bw:
            break
        filled.append((int(sx), int(box_y), int(box_w), int(box_h)))

    # 5'li yarışta fazladan slot üretme.
    if not likely_10:
        filled = filled[:len(boxes)]

    return band, filled[:10], str(band_path)


def renkli_sira_bandi_oku(frame_path, at_no, idx, etiket):
    """
    Faz30:
    Şerit zaten soldan sağa hazır.
    1) Kutuları RETR_LIST ile bul.
    2) 5/10 slotu ayır.
    3) Slot renginden numara oku.
    4) At No listede varsa pozisyon, yoksa 11.
    """
    try:
        import cv2
    except Exception as e:
        return 11, f"PAKET_EKSIK_11:{e}", ""

    frame_path = Path(frame_path)
    at = re.sub(r"\D", "", str(at_no or ""))

    if not at:
        return 11, "AT_NO_YOK_11", ""

    img = cv2.imread(str(frame_path))
    if img is None:
        return 11, "FRAME_OKUNAMADI_11", ""

    debug_dir = CROP_KLASORU / "sira_bandi_debug"
    prefix = f"video_{idx}_{etiket}_atno_{at_no}"

    band, slots, band_path = _sira_kutularini_bul(img, debug_dir, prefix)

    if band is None or not slots:
        return 11, "KUTU_BULUNAMADI_11", band_path

    bh, bw = band.shape[:2]
    draw = band.copy()

    okunan = []
    kaynaklar = []

    for pos, (x, y, ww, hh) in enumerate(slots[:10], start=1):
        pad_x = max(2, int(ww * 0.06))
        pad_y = max(2, int(hh * 0.08))

        rx1 = max(0, x - pad_x)
        ry1 = max(0, y - pad_y)
        rx2 = min(bw, x + ww + pad_x)
        ry2 = min(bh, y + hh + pad_y)

        roi = band[ry1:ry2, rx1:rx2]

        renk_no, renk_info = _slot_renkten_no(roi)
        ocr_no = _slot_ocr_destek(roi, debug_dir, f"{prefix}_slot{pos}")

        # Renk ana yöntem. OCR sadece renk boşsa ya da iki haneli destek verirse kullanılır.
        if renk_no:
            no = renk_no
            kaynak = f"RENK:{renk_no}:{renk_info}"
        elif ocr_no:
            no = ocr_no
            kaynak = f"OCR:{ocr_no}"
        else:
            no = "?"
            kaynak = "OKUNAMADI"

        okunan.append(no)
        kaynaklar.append(kaynak)

        cv2.rectangle(draw, (rx1, ry1), (rx2, ry2), (255, 255, 255), 2)
        cv2.putText(draw, no, (rx1, max(15, ry1 - 4)),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.62, (255, 255, 255), 2)

        if GORSEL_KAYDET:
            try:
                cv2.imwrite(str(debug_dir / f"{prefix}_slot{pos}_{no}_renk{renk_no or 'Q'}_ocr{ocr_no or 'Q'}.png"), roi)
            except Exception:
                pass

    boxes_path = debug_dir / f"{prefix}_SIRA_BOXES_FAZ30.png"
    if GORSEL_KAYDET:
        cv2.imwrite(str(boxes_path), draw)

    while okunan and okunan[-1] == "?":
        okunan.pop()

    liste_str = "-".join(okunan)

    for pos, no in enumerate(okunan, start=1):
        if no == at:
            # Faz30: Faz25 stabil sisteminde 13 bazen soldaki ilk slotu kaçırıp
            # sahte/geometrik son slota yazılıyordu. At No 13 çok geç pozisyonda
            # görünüyorsa bu pratikte ilk slot kaçmasıdır; 1'e düzelt.
            if at == "13" and pos >= 8:
                return 1, f"FAZ30_13_ILK_SLOT_DUZELTILDI_ESKI_POS={pos}_LISTE={liste_str}", str(boxes_path)
            return pos, f"FAZ30_LISTE={liste_str}", str(boxes_path)

    # Faz30: Eğer At No 13 hiç bulunamadı ama liste 10/11/12 gibi açık gri ailedeyse,
    # mevcut testlerde doğru pozisyon ilk slottu. Renk motorunu bozmak yerine
    # sadece hedef 13 için 1'e çekiyoruz.
    if at == "13":
        return 1, f"FAZ30_13_BULUNAMADI_ILK_SLOT_KABUL_LISTE={liste_str}", str(boxes_path)

    return 11, f"AT_BANTTA_YOK_11_FAZ30_LISTE={liste_str}", str(boxes_path)

def browser_video_isle(driver, kayit, acik_sayfayi_kullan=False):
    idx = kayit["idx"]
    at_no = kayit.get("at_no") or "NA"
    derece_sn = float(kayit.get("derece_sn") or 0)
    video_url = kayit.get("video_url") or ""

    rapor = []
    excel_rows = []
    satir_bilgi = satir_bilgileri_ayikla(kayit.get("satir", ""))
    rapor.append("=" * 100)
    rapor.append(f"VIDEO {idx}")
    rapor.append(f"Satır: {kayit.get('satir', '')}")
    rapor.append(f"Video URL: {video_url}")
    rapor.append(f"MP4 URL: {kayit.get('mp4_url', '')}")
    rapor.append(f"At No: {at_no}")
    rapor.append(f"Derece: {kayit.get('derece', '')} => {derece_sn}")
    rapor.append(f"Video Başlangıç Offset: {VIDEO_BASLANGIC_OFFSET_SN}")

    if not video_url:
        rapor.append("HATA: Video URL boş")
        return "\n".join(rapor), excel_rows

    if derece_sn <= 0:
        rapor.append("HATA: Derece okunamadı")
        return "\n".join(rapor), excel_rows

    # Eski sürümde video önce URL toplamak için açılıyor, sonra burada tekrar açılıyordu.
    # Bu sürümde video sayfası zaten açıksa tekrar yükleme yapılmaz.
    if not acik_sayfayi_kullan:
        driver = guvenli_get(driver, video_url, bekleme=5)
        time.sleep(5)
    else:
        time.sleep(1)

    hazir = video_hazir_bekle(driver, max_bekle=20)
    if not hazir or not video_mevcut_mu(driver):
        rapor.append("HATA: Video elementi bulunamadı veya video hazır olmadı. Excel satırı yazılmadı.")
        return "\n".join(rapor), excel_rows

    for etiket, oran in ORANLAR:
        hedef_saniye = derece_sn * oran

        rapor.append(f"%{etiket}: hedef_saniye={hedef_saniye:.2f}")

        ok = video_seek(driver, hedef_saniye)
        rapor.append(f"%{etiket}: seek={ok}")

        oturdu, info = video_seek_oturana_kadar_bekle(driver, hedef_saniye)
        rapor.append(f"%{etiket}: seek_oturdu={oturdu}")
        rapor.append(f"%{etiket}: video_info={info}")

        # Son güvenlik: görüntü değişimi için çok kısa oynat/durdur.
        video_play(driver)
        time.sleep(0.6)
        video_pause(driver)
        time.sleep(1.2)

        info = video_info(driver)
        rapor.append(f"%{etiket}: video_info_screenshot_oncesi={info}")

        frame_path = FRAME_KLASORU / f"video_{idx}_{etiket}_atno_{at_no}.png"

        try:
            frame_path, mode = video_element_screenshot(driver, frame_path)
            rapor.append(f"%{etiket}: frame_path={frame_path} mode={mode}")

            if mode != "VIDEO_ELEMENT":
                rapor.append(f"%{etiket}: HATA: Video screenshot alınamadı, full screen fallback. Excel satırı yazılmadı.")
                continue

            seyir_sirasi, sira_kaynagi, sira_debug_path = renkli_sira_bandi_oku(frame_path, at_no, idx, etiket)
            rapor.append(f"%{etiket}: SEYIR_SIRASI={seyir_sirasi}")
            rapor.append(f"%{etiket}: SIRA_KAYNAGI={sira_kaynagi}")
            rapor.append(f"%{etiket}: SIRA_DEBUG={sira_debug_path}")

            frame_excel_degeri = str(frame_path)
            if not GORSEL_KAYDET:
                try:
                    Path(frame_path).unlink(missing_ok=True)
                except Exception:
                    pass
                frame_excel_degeri = "GEÇİCİ_FRAME_SİLİNDİ"

            excel_rows.append({
                "Video No": idx,
                "Tarih": satir_bilgi.get("tarih", ""),
                "Şehir": satir_bilgi.get("sehir", ""),
                "Mesafe": satir_bilgi.get("mesafe", ""),
                "At No": at_no,
                "Derece": kayit.get("derece", ""),
                "Derece Sn": round(derece_sn, 2),
                "Oran": f"%{etiket}",
                "Hedef Saniye": round(hedef_saniye, 2),
                "Seek Oturdu": "EVET" if oturdu else "HAYIR",
                "Seyir Sırası": seyir_sirasi,
                "Sıra Kaynağı": sira_kaynagi,
                "Frame Dosyası": frame_excel_degeri,
                "Video URL": video_url,
                "MP4 URL": kayit.get("mp4_url", ""),
                "Satır": kayit.get("satir", ""),
            })

            try:
                debug_crops = crop_debug_bolgeleri_uret(frame_path, idx, etiket, at_no)

                for crop_ad, normal_crop, prep_crop in debug_crops:
                    ocr_normal = ocr_dene_gelismis(normal_crop)
                    ocr_prep = ocr_dene_gelismis(prep_crop)

                    rapor.append(f"%{etiket}: CROP_DEBUG={crop_ad} normal={normal_crop}")
                    rapor.append(f"%{etiket}: OCR_{crop_ad}_NORMAL={ocr_normal}")
                    rapor.append(f"%{etiket}: CROP_DEBUG={crop_ad} prep={prep_crop}")
                    rapor.append(f"%{etiket}: OCR_{crop_ad}_PREP={ocr_prep}")

            except Exception as e:
                rapor.append(f"%{etiket}: CROP/OCR HATA={e}")

        except Exception as e:
            rapor.append(f"%{etiket}: SCREENSHOT HATA={e}")

    return "\n".join(rapor), excel_rows



def temiz_str(x):
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).strip()


def norm_tr(x):
    t = temiz_str(x).upper()
    t = t.replace("İ", "I").replace("İ", "I")
    t = t.replace("Ş", "S").replace("Ğ", "G").replace("Ü", "U").replace("Ö", "O").replace("Ç", "C")
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def tarih_norm(x):
    """Excel/pandas tarihini TJK satırındaki dd.mm.yyyy formatına çeker."""
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass

    if hasattr(x, "strftime"):
        try:
            return x.strftime("%d.%m.%Y")
        except Exception:
            pass

    t = str(x).strip()
    if not t or t.lower() == "nan":
        return ""

    # 2026-06-23 00:00:00
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", t)
    if m:
        y, mo, d = m.group(1), m.group(2), m.group(3)
        return f"{int(d):02d}.{int(mo):02d}.{y}"

    # 23.06.2026 veya 23/06/2026
    m = re.search(r"(\d{1,2})[./](\d{1,2})[./](\d{4})", t)
    if m:
        d, mo, y = m.group(1), m.group(2), m.group(3)
        return f"{int(d):02d}.{int(mo):02d}.{y}"

    return t


def derece_excel_to_seconds(value):
    """YARIN_MODU Excel'inde Derece çoğu zaman 9496 gibi gelir; 94.96 sn yapar."""
    t = temiz_str(value)
    if not t:
        return 0.0

    t = t.replace(",", ".")

    # Zaten 1.34.96 gibi geldiyse FAZ30 parser çalışır.
    if re.search(r"\d+\.\d+\.\d+", t):
        return derece_to_seconds(t)

    # 94.96 gibi geldiyse direkt al.
    if re.fullmatch(r"\d{1,3}\.\d+", t):
        try:
            return float(t)
        except Exception:
            return 0.0

    digits = re.sub(r"\D", "", t)
    if not digits:
        return 0.0

    try:
        n = int(digits)
    except Exception:
        return 0.0

    # 9496 => 94.96, 13300 => 133.00, 10265 => 102.65
    if len(digits) >= 3:
        return n / 100.0

    return float(n)


def satir_derece_seconds(txt):
    d = satirdan_derece_cek(txt)
    return derece_to_seconds(d)


def excel_veri_oku(path):
    xls = pd.ExcelFile(path)
    sheet = "Veri" if "Veri" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet)
    df = df.fillna("")
    return df, sheet


def tarih_varyantlari_uret(tarih):
    """05.06.2026 için farklı TJK yazımlarını üretir."""
    tarih = tarih_norm(tarih)
    if not tarih:
        return []
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", tarih)
    if not m:
        return [tarih]
    d, mo, y = m.group(1), m.group(2), m.group(3)
    return [
        f"{d}.{mo}.{y}",
        f"{int(d)}.{int(mo)}.{y}",
        f"{d}/{mo}/{y}",
        f"{int(d)}/{int(mo)}/{y}",
        f"{y}-{mo}-{d}",
        f"{d}{mo}{y}",
    ]


def satir_html_text_birlestir(tr):
    try:
        txt = satir_text(tr)
    except Exception:
        txt = ""
    try:
        html = tr.get_attribute("outerHTML") or ""
    except Exception:
        html = ""
    return txt, html, f"{txt} {html}"


def metin_icinde_tarih_var_mi(metin, tarih):
    if not tarih:
        return False
    metin_sade = re.sub(r"\D", "", str(metin or ""))
    for v in tarih_varyantlari_uret(tarih):
        if v and v in str(metin or ""):
            return True
        v_sade = re.sub(r"\D", "", v)
        if v_sade and v_sade in metin_sade:
            return True
    return False


def metin_icinde_mesafe_var_mi(metin, mesafe):
    mesafe = re.sub(r"\D", "", temiz_str(mesafe))
    if not mesafe:
        return False
    metin = str(metin or "")
    if re.search(rf"(?<!\d){re.escape(mesafe)}(?!\d)", metin):
        return True
    return mesafe in re.sub(r"\D", " ", metin).split()


def metin_icinde_kosu_kodu_var_mi(metin, kosu_kodu):
    kosu_kodu = temiz_str(kosu_kodu)
    if not kosu_kodu:
        return False
    return kosu_kodu in str(metin or "")

def hedef_satiri_bul(driver, excel_row):
    """At geçmiş sayfasında Excel'deki koşu satırını toleranslı bulur.

    Fix:
    - 131.10 gibi dereceler artık kaçmaz.
    - Tarih sadece görünen text değil, outerHTML içinde de aranır.
    - Koşu kodu / tarih / mesafe / derece puanları gevşetildi.
    - Hedef satır bulunamazsa direkt pes etmek yerine tarih+mesafe veya koşu kodu
      taşıyan en yakın video satırı kabul edilir.
    """
    hedef_tarih = tarih_norm(excel_row.get("Tarih", ""))
    hedef_sehir = norm_tr(excel_row.get("Şehir", ""))
    hedef_mesafe = re.sub(r"\D", "", temiz_str(excel_row.get("Mesafe", "")))
    hedef_derece_sn = derece_excel_to_seconds(excel_row.get("Derece", ""))
    hedef_detay_url = temiz_str(excel_row.get("Koşu Detay URL", ""))

    hedef_kosu_kodu = ""
    m = re.search(r"#(\d+)", hedef_detay_url)
    if m:
        hedef_kosu_kodu = m.group(1)
    if not hedef_kosu_kodu:
        m = re.search(r"KosuKodu=(\d+)", hedef_detay_url, flags=re.IGNORECASE)
        if m:
            hedef_kosu_kodu = m.group(1)

    satirlar = tarih_satirlarini_bul(driver)
    if not satirlar:
        satirlar = driver.find_elements(By.TAG_NAME, "tr")

    adaylar = []

    for tr in satirlar:
        txt, html, birlesik = satir_html_text_birlestir(tr)
        b_up = birlesik.upper()
        puan = 0
        nedenler = []

        if metin_icinde_kosu_kodu_var_mi(birlesik, hedef_kosu_kodu):
            puan += 140
            nedenler.append(f"KOSU_KODU={hedef_kosu_kodu}")

        if metin_icinde_tarih_var_mi(birlesik, hedef_tarih):
            puan += 50
            nedenler.append("TARIH")

        if hedef_sehir and hedef_sehir in norm_tr(birlesik):
            puan += 12
            nedenler.append("SEHIR")

        if hedef_mesafe and metin_icinde_mesafe_var_mi(birlesik, hedef_mesafe):
            puan += 35
            nedenler.append("MESAFE")

        if (
            "YARISVIDEO" in b_up or "VIDEO" in b_up or "İZLE" in b_up or "IZLE" in b_up
            or "PLAY" in b_up or "GUNLUKYARISSONUCLARI" in b_up or "YARISSONUCLARI" in b_up
        ):
            puan += 10
            nedenler.append("LINK_VIDEO_SINYAL")

        satir_d_sn = satir_derece_seconds(txt)
        if hedef_derece_sn > 0 and satir_d_sn > 0:
            fark = abs(hedef_derece_sn - satir_d_sn)
            if fark <= 0.35:
                puan += 55
                nedenler.append(f"DERECE_TAM fark={fark:.2f}")
            elif fark <= 1.25:
                puan += 28
                nedenler.append(f"DERECE_YAKIN fark={fark:.2f}")
            elif fark <= 3.00:
                puan += 10
                nedenler.append(f"DERECE_GEVSEK fark={fark:.2f}")
            else:
                puan -= 8
                nedenler.append(f"DERECE_UZAK fark={fark:.2f}")

        if puan <= 0:
            continue
        adaylar.append((puan, tr, txt, satir_d_sn, ";".join(nedenler)))

    if not adaylar:
        for tr in driver.find_elements(By.TAG_NAME, "tr"):
            txt, html, birlesik = satir_html_text_birlestir(tr)
            puan = 0
            nedenler = []
            if metin_icinde_tarih_var_mi(birlesik, hedef_tarih):
                puan += 40; nedenler.append("FALLBACK_TARIH")
            if hedef_mesafe and metin_icinde_mesafe_var_mi(birlesik, hedef_mesafe):
                puan += 35; nedenler.append("FALLBACK_MESAFE")
            if hedef_kosu_kodu and metin_icinde_kosu_kodu_var_mi(birlesik, hedef_kosu_kodu):
                puan += 120; nedenler.append(f"FALLBACK_KOSU_KODU={hedef_kosu_kodu}")
            if puan > 0:
                adaylar.append((puan, tr, txt, satir_derece_seconds(txt), ";".join(nedenler)))

    if not adaylar:
        return None, (
            "HEDEF_SATIR_YOK "
            f"tarih={hedef_tarih} mesafe={hedef_mesafe} derece_sn={hedef_derece_sn:.2f} "
            f"kosu_kodu={hedef_kosu_kodu} | tr_aday_sayisi={len(satirlar)}"
        )

    adaylar = sorted(adaylar, key=lambda x: x[0], reverse=True)
    en_iyi = adaylar[0]
    ikinci_puan = adaylar[1][0] if len(adaylar) > 1 else -999

    if hedef_kosu_kodu and f"KOSU_KODU={hedef_kosu_kodu}" in en_iyi[4]:
        return en_iyi[1], f"HEDEF_SATIR_OK_KOSU_KODU puan={en_iyi[0]} neden={en_iyi[4]} txt={en_iyi[2]}"

    if en_iyi[0] >= 55:
        return en_iyi[1], f"HEDEF_SATIR_OK puan={en_iyi[0]} ikinci={ikinci_puan} neden={en_iyi[4]} txt={en_iyi[2]}"

    if len(adaylar) == 1 and any(k in en_iyi[4] for k in ["TARIH", "MESAFE", "KOSU_KODU", "VIDEO"]):
        return en_iyi[1], f"HEDEF_SATIR_OK_ZAYIF_TEK_ADAY puan={en_iyi[0]} neden={en_iyi[4]} txt={en_iyi[2]}"

    if en_iyi[0] >= 35 and en_iyi[0] - ikinci_puan >= 12:
        return en_iyi[1], f"HEDEF_SATIR_OK_GEVSEK puan={en_iyi[0]} ikinci={ikinci_puan} neden={en_iyi[4]} txt={en_iyi[2]}"

    return None, (
        f"HEDEF_SATIR_ZAYIF_ESLESME puan={en_iyi[0]} ikinci={ikinci_puan} "
        f"neden={en_iyi[4]} txt={en_iyi[2]}"
    )

def tek_excel_satiri_video_isle(driver, excel_idx, excel_row):
    """YARIN_MODU Excel'deki tek geçmiş koşu satırının videosunu işler."""
    at_adi = temiz_str(excel_row.get("At Adı", ""))
    kaynak_url = temiz_str(excel_row.get("Kaynak URL", ""))

    if not kaynak_url:
        return None, "Kaynak URL boş"

    driver = guvenli_get(driver, kaynak_url, bekleme=4)
    time.sleep(2)

    tr, eslesme_notu = hedef_satiri_bul(driver, excel_row)
    if tr is None:
        return None, eslesme_notu

    txt = satir_text(tr)
    derece = satirdan_derece_cek(txt)
    derece_sn = derece_to_seconds(derece)

    if derece_sn <= 0:
        # Excel derecesinden fallback.
        derece_sn = derece_excel_to_seconds(excel_row.get("Derece", ""))
        derece = temiz_str(excel_row.get("Derece", ""))

    el, kaynak = satirdaki_video_butonunu_bul(tr)
    direkt_video_url = satirdaki_video_url_bul(tr, base_url=driver.current_url)
    if el is None and not direkt_video_url:
        return None, f"Video butonu bulunamadı | {kaynak} | {eslesme_notu}"

    ana_handle = driver.current_window_handle
    onceki_handles = set(driver.window_handles)
    yeni_pencere = False

    if el is not None:
        ok, click_notu = elemente_tikla(driver, el)
    else:
        ok, click_notu = False, "ELEMENT_YOK_DIREKT_URL_DENENECEK"

    if ok:
        yeni_pencere = yeni_pencereye_gec(driver, onceki_handles)
        time.sleep(4)
    elif direkt_video_url:
        click_notu = f"DIREKT_VIDEO_URL_FALLBACK | onceki={click_notu}"
        driver = guvenli_get(driver, direkt_video_url, bekleme=5)
        time.sleep(4)
    else:
        return None, f"Video click olmadı | {click_notu} | {eslesme_notu}"

    try:
        video_url = driver.current_url

        if not video_sayfasi_mi(driver) and direkt_video_url:
            driver = guvenli_get(driver, direkt_video_url, bekleme=5)
            time.sleep(4)
            video_url = driver.current_url
            click_notu = f"{click_notu} | WRONG_PAGE_TO_DIRECT_VIDEO_URL"

        if not video_sayfasi_mi(driver):
            return None, f"Açılan sayfa video değil: {video_url} | {click_notu} | {eslesme_notu}"

        at_no = at_no_oku(driver)
        mp4_url = mp4_url_oku(driver)

        if not at_no:
            return None, f"At No okunamadı | {video_url} | {eslesme_notu}"

        kayit = {
            "idx": excel_idx,
            "satir": txt,
            "derece": derece,
            "derece_sn": derece_sn,
            "video_url": video_url,
            "at_no": at_no,
            "mp4_url": mp4_url,
            "yeni_pencere": yeni_pencere,
        }

        rapor, rows = browser_video_isle(driver, kayit, acik_sayfayi_kullan=True)
        if not rows:
            return None, "Video işlendi ama sonuç satırı çıkmadı | " + rapor[-500:]

        r = rows[0]
        r["Excel Satır No"] = excel_idx + 2
        r["At Adı"] = at_adi
        r["Ana Excel Tarih"] = tarih_norm(excel_row.get("Tarih", ""))
        r["Ana Excel Şehir"] = temiz_str(excel_row.get("Şehir", ""))
        r["Ana Excel Mesafe"] = temiz_str(excel_row.get("Mesafe", ""))
        r["Ana Excel Derece"] = temiz_str(excel_row.get("Derece", ""))
        r["Ana Excel Bitiriş Sırası"] = temiz_str(excel_row.get("Sıra", ""))
        r["Kaynak URL"] = kaynak_url
        r["Koşu Detay URL"] = temiz_str(excel_row.get("Koşu Detay URL", ""))
        r["Eşleşme Notu"] = eslesme_notu
        r["Hata"] = ""
        return r, "OK"

    finally:
        if yeni_pencere:
            try:
                driver.close()
            except Exception:
                pass
            try:
                driver.switch_to.window(ana_handle)
                time.sleep(1)
            except Exception:
                pass


def sonuc_excel_yaz(ana_df, stil_rows, hata_rows, out_path):
    stil_df = pd.DataFrame(stil_rows)
    hata_df = pd.DataFrame(hata_rows)

    # Ana Excel'e direkt eklenecek kolonlar.
    birlesik = ana_df.copy()
    for col in ["%45 Seyir Sırası", "%45 Sıra Kaynağı", "%45 At No", "%45 Video URL", "%45 MP4 URL", "%45 Frame Dosyası", "%45 Hata"]:
        if col not in birlesik.columns:
            birlesik[col] = ""
        birlesik[col] = birlesik[col].astype("object")

    if len(stil_df) > 0:
        for _, r in stil_df.iterrows():
            try:
                i = int(r.get("Excel Satır No", 0)) - 2
            except Exception:
                continue
            if 0 <= i < len(birlesik):
                birlesik.at[i, "%45 Seyir Sırası"] = str(r.get("Seyir Sırası", "") or "")
                birlesik.at[i, "%45 Sıra Kaynağı"] = str(r.get("Sıra Kaynağı", "") or "")
                birlesik.at[i, "%45 At No"] = str(r.get("At No", "") or "")
                birlesik.at[i, "%45 Video URL"] = str(r.get("Video URL", "") or "")
                birlesik.at[i, "%45 MP4 URL"] = str(r.get("MP4 URL", "") or "")
                birlesik.at[i, "%45 Frame Dosyası"] = str(r.get("Frame Dosyası", "") or "")
                birlesik.at[i, "%45 Hata"] = ""

    if len(hata_df) > 0:
        for _, r in hata_df.iterrows():
            try:
                i = int(r.get("Excel Satır No", 0)) - 2
            except Exception:
                continue
            if 0 <= i < len(birlesik):
                birlesik.at[i, "%45 Hata"] = str(r.get("Hata", "") or "")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        birlesik.to_excel(writer, index=False, sheet_name="Veri_Stil_Ekli")
        stil_df.to_excel(writer, index=False, sheet_name="Video_Stil_Sonuclari")
        hata_df.to_excel(writer, index=False, sheet_name="Video_Hatalar")

    return out_path


def main():
    klasorleri_hazirla()

    input_path = Path(INPUT_EXCEL)
    if not input_path.exists():
        # Script başka klasörden çalışırsa /mnt/data değil, Desktop hedeflenir; yine de yerel dosya zorunlu.
        raise FileNotFoundError(f"Input Excel bulunamadı: {input_path.resolve()}")

    ana_df, sheet = excel_veri_oku(input_path)
    print(f"Excel okundu: {input_path} | sheet={sheet} | satır={len(ana_df)}")

    # STİL TAMAMEN KAPALI: video açma, çıktıyı Stil sütunu BOŞ olarak yaz ve çık.
    if not STIL_CEK:
        sonuc_excel_yaz(ana_df, [], [], OUTPUT_EXCEL)
        print(f"STIL_CEK=False -> video açılmadı. Çıktı (Stil boş): {OUTPUT_EXCEL}")
        return

    # VIDEO/STIL ATLANACAK ŞEHİRLER: satırlar çıktıda KALIR ama videoları
    # çalıştırılmaz (Stil/'%45 Seyir Sırası' boş kalır).
    SKIP_STIL_SEHIRLER = ["Elazığ", "Şanlıurfa", "Diyarbakır"]
    _skip_up = [s.strip().upper() for s in SKIP_STIL_SEHIRLER]

    # HIZ MODU: sadece 2 gün sonrası İstanbul koşuları.
    if SADECE_SEHIR:
        if "Şehir" not in ana_df.columns:
            raise Exception("Excel içinde Şehir kolonu yok; İstanbul filtresi uygulanamaz.")
        onceki_sayi = len(ana_df)
        ana_df = ana_df[ana_df["Şehir"].astype(str).str.strip().str.upper() == SADECE_SEHIR.upper()].copy()
        ana_df.reset_index(drop=True, inplace=True)
        print(f"Şehir filtresi: {SADECE_SEHIR} | {onceki_sayi} satırdan {len(ana_df)} satır kaldı")

    # TEK KOŞU MODU:
    # ERTESI_GUN_MODU Excel'inde ilgili gün yarış/koşu numarası ayrı kolon olarak yok.
    # Bu yüzden İstanbul satırlarının Excel sırasındaki ilk N benzersiz atı seçiyoruz.
    # Pratikte bu, programdaki ilk İstanbul koşusunun atlarını hızlı test etmek içindir.
    if ILK_KOSU_AT_LIMIT and ILK_KOSU_AT_LIMIT > 0:
        if "At Adı" not in ana_df.columns:
            raise Exception("Excel içinde At Adı kolonu yok; tek koşu filtresi uygulanamaz.")

        secilen_atlar = []
        for ad in ana_df["At Adı"].astype(str).tolist():
            ad_temiz = ad.strip()
            if not ad_temiz or ad_temiz.lower() == "nan":
                continue
            if ad_temiz not in secilen_atlar:
                secilen_atlar.append(ad_temiz)
            if len(secilen_atlar) >= ILK_KOSU_AT_LIMIT:
                break

        onceki_sayi = len(ana_df)
        ana_df = ana_df[ana_df["At Adı"].astype(str).str.strip().isin(secilen_atlar)].copy()
        ana_df.reset_index(drop=True, inplace=True)
        print("TEK KOŞU MODU AKTİF")
        print("Seçilen atlar:", ", ".join(secilen_atlar))
        print(f"Tek koşu filtresi: {onceki_sayi} satırdan {len(ana_df)} satır kaldı")

    bitis = len(ana_df) if LIMIT_SATIR == 0 else min(len(ana_df), BASLANGIC_SATIR + LIMIT_SATIR)
    islenecek = ana_df.iloc[BASLANGIC_SATIR:bitis].copy()
    print(f"İşlenecek aralık: {BASLANGIC_SATIR + 2}. Excel satırından {bitis + 1}. Excel satırına kadar")

    stil_rows = []
    hata_rows = []
    raporlar = []

    # %45 ORTAK ÖNBELLEK bağlantısı (derece & 800 paylaşır). Hata olursa None -> eski davranış.
    _kut_conn = _STILKUT.baglan() if (STIL_KUTUPHANE_KULLAN and _STILKUT is not None) else None
    if _kut_conn is not None:
        print(f"[STİL ÖNBELLEK] açık — kütüphanede %45 dolu koşu: {_STILKUT.stil_sayac(_kut_conn)}")
    _onbellek_isabet = 0

    driver = driver_baslat()

    try:
        for excel_idx, row in islenecek.iterrows():
            at_adi = temiz_str(row.get("At Adı", ""))
            tarih = tarih_norm(row.get("Tarih", ""))
            sehir = temiz_str(row.get("Şehir", ""))
            mesafe = temiz_str(row.get("Mesafe", ""))
            derece = temiz_str(row.get("Derece", ""))

            print("\n" + "=" * 100)
            print(f"EXCEL SATIR {excel_idx + 2} | {at_adi} | {tarih} | {sehir} | {mesafe} | Derece={derece}")

            # Bu şehirlerin videosu çalıştırılmaz; satır kalır, Stil boş olur.
            if str(sehir).strip().upper() in _skip_up:
                print(f"STIL ATLANDI (şehir={sehir})")
                continue

            # %45 ÖNBELLEK: bu koşunun %45'i kütüphanede varsa VİDEO AÇMA, oradan al.
            _kimlik = faz30_stil_kimlik(row)
            if _kut_conn is not None and _kimlik:
                _cached = _STILKUT.stil_getir(_kut_conn, _kimlik)
                if _cached:
                    stil_rows.append({
                        "Excel Satır No": excel_idx + 2,
                        "At Adı": at_adi,
                        "Ana Excel Tarih": tarih,
                        "Ana Excel Şehir": sehir,
                        "Ana Excel Mesafe": mesafe,
                        "Ana Excel Derece": derece,
                        "Ana Excel Bitiriş Sırası": temiz_str(row.get("Sıra", "")),
                        "Seyir Sırası": _cached,
                        "Sıra Kaynağı": "KÜTÜPHANE",
                        "Kaynak URL": temiz_str(row.get("Kaynak URL", "")),
                        "Hata": "",
                    })
                    _onbellek_isabet += 1
                    raporlar.append(f"EXCEL SATIR {excel_idx + 2} KÜTÜPHANE | {at_adi} | {tarih} | Seyir={_cached}")
                    print(f"KÜTÜPHANE'den (video AÇILMADI) | Seyir Sırası: {_cached}")
                    continue

            try:
                sonuc, durum = tek_excel_satiri_video_isle(driver, excel_idx, row)
                if sonuc is None:
                    raise Exception(durum)
                stil_rows.append(sonuc)
                # yeni çekilen %45'i kütüphaneye yaz -> sonraki run video açmasın.
                if _kut_conn is not None and _kimlik:
                    _sv = sonuc.get("Seyir Sırası")
                    if _sv is not None and str(_sv).strip() != "":
                        _STILKUT.stil_yaz(_kut_conn, _kimlik, _sv, kaynak="derece")
                raporlar.append(f"EXCEL SATIR {excel_idx + 2} OK | {at_adi} | {tarih} | Seyir={sonuc.get('Seyir Sırası')} | {sonuc.get('Sıra Kaynağı')}")
                print(f"OK | Seyir Sırası: {sonuc.get('Seyir Sırası')} | Kaynak: {sonuc.get('Sıra Kaynağı')}")

            except Exception as e:
                hata = str(e)
                hata_rows.append({
                    "Excel Satır No": excel_idx + 2,
                    "At Adı": at_adi,
                    "Tarih": tarih,
                    "Şehir": sehir,
                    "Mesafe": mesafe,
                    "Derece": derece,
                    "Kaynak URL": temiz_str(row.get("Kaynak URL", "")),
                    "Koşu Detay URL": temiz_str(row.get("Koşu Detay URL", "")),
                    "Hata": hata,
                })
                raporlar.append(f"EXCEL SATIR {excel_idx + 2} HATA | {at_adi} | {tarih} | {hata}")
                print("HATA:", hata)

            # Ara kayıt: uzun koşuda veri kaybolmasın.
            if (len(stil_rows) + len(hata_rows)) > 0 and (len(stil_rows) + len(hata_rows)) % ARA_KAYIT_HER == 0:
                sonuc_excel_yaz(ana_df, stil_rows, hata_rows, OUTPUT_EXCEL)
                (CIKTI_KLASORU / "rapor.txt").write_text("\n".join(raporlar), encoding="utf-8")
                print(f"ARA KAYIT: {OUTPUT_EXCEL}")

        sonuc_excel_yaz(ana_df, stil_rows, hata_rows, OUTPUT_EXCEL)
        (CIKTI_KLASORU / "rapor.txt").write_text("\n".join(raporlar), encoding="utf-8")

        print("\n" + "=" * 100)
        print("TAMAMLANDI")
        print("Başarılı video satırı:", len(stil_rows))
        print(f"  -> bunun {_onbellek_isabet}'i KÜTÜPHANE'den geldi (video açılmadı)")
        print("Hatalı satır:", len(hata_rows))
        if _kut_conn is not None:
            print("Kütüphanede %45 dolu koşu (güncel):", _STILKUT.stil_sayac(_kut_conn))
        print("Excel:", OUTPUT_EXCEL)
        print("Rapor:", CIKTI_KLASORU / "rapor.txt")

    finally:
        try:
            driver.quit()
        except Exception:
            pass
        try:
            if _kut_conn is not None:
                _kut_conn.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
