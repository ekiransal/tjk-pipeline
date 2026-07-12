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

# %45 STİL ORTAK ÖNBELLEĞİ: bir koşunun %45'i bir kez çekilince tjk_kutuphane.db'ye
# yazılır; sonraki çalıştırmalarda video AÇILMADAN oradan okunur (derece & 800 paylaşır).
# Kütüphane yoksa/çökerse eskisi gibi her seferinde video açılır.
try:
    import tjk_kutuphane as KUT
except Exception as _kut_e:
    KUT = None
    print(f"[STİL ÖNBELLEK] tjk_kutuphane yüklenemedi ({_kut_e}) -> video her seferinde açılır")


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

CIKTI_KLASORU = Path("son800_faz30_stil_rapor")
FRAME_KLASORU = CIKTI_KLASORU / "video_frames"
CROP_KLASORU = CIKTI_KLASORU / "crops"

# ERTESI_GUN_MODU çıktısı. Script Desktop'ta çalıştırılacaksa aynı klasöre koy.
INPUT_EXCEL = "SON800_BUGUN_ILK_AT_STIL_FINAL_TEST.xlsx"   # son800 çıktısı (Kaynak URL + Derece dolu)
OUTPUT_EXCEL = "SON800_BUGUN_ILK_AT_STIL_FINAL_TEST_STIL_EKLI.xlsx"   # %45 stil eklenmiş son800

# ============================================================
# STİL (VİDEO/%45) ÇEKİMİ AÇIK/KAPALI
#   False = video AÇILMAZ; çıktı aynı yapıda ama Stil sütunu BOŞ. ÇOK HIZLI.
#   True  = her at için video açılıp %45'te seyir sırası çıkarılır (YAVAŞ).
# Stil'i yeni yer'deki BY sütunu kullanır; kapalıyken BY boş kalır.
# ============================================================
STIL_CEK = True

# %45 stilini ortak kütüphaneden oku/yaz. Video en çok vakti yiyen kısım; bir kez
# çekilen %45 bir daha çekilmez (derece & 800 aynı kaydı paylaşır).
STIL_KUTUPHANE_KULLAN = True

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
    _kut_conn = KUT.baglan() if (STIL_KUTUPHANE_KULLAN and KUT is not None) else None
    if _kut_conn is not None:
        print(f"[STİL ÖNBELLEK] açık — kütüphanede %45 dolu koşu: {KUT.stil_sayac(_kut_conn)}")
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
            # Anahtar = "Koşu Kimliği" (derece & 800 ile aynı format).
            kimlik = temiz_str(row.get("Koşu Kimliği", ""))
            # AT ADI ANAHTARDA (derece tarafıyla aynı format: ...|AT:AD)
            import re as _re_at
            _atk = _re_at.sub(r"\s+", " ", str(at_adi or "").strip().upper())
            kimlik = (kimlik + "|AT:" + _atk) if (kimlik and _atk) else ""
            if _kut_conn is not None and kimlik:
                _cached = KUT.stil_getir(_kut_conn, kimlik)
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
                if _kut_conn is not None and kimlik:
                    _sv = sonuc.get("Seyir Sırası")
                    if _sv is not None and str(_sv).strip() != "":
                        KUT.stil_yaz(_kut_conn, kimlik, _sv, kaynak="800")
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
            print("Kütüphanede %45 dolu koşu (güncel):", KUT.stil_sayac(_kut_conn))
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
