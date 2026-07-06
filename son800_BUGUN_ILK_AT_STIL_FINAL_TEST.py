# -*- coding: utf-8 -*-
"""
BUGUN + ILK AT KESIN DUR TEST
- TJK günlük programda BUGÜNÜ açar. 2 gün sonraya gitmez.
- İlk şehirdeki ilk atı alır.
- O atın son 1 yıldaki SADECE 1. olduğu ilk geçmiş koşuya bakar.
- Sonuç sayfasına girer, Son 800 değerini ve foto/finiş sırasını çekmeye çalışır.
- Başarılı da olsa hatalı da olsa BAŞKA ATA GEÇMEZ.
- Excel oluşturup kapanır.
"""

import re
import sys
import time
from io import StringIO
from datetime import datetime, timedelta
from urllib.parse import urlencode

import pandas as pd
import requests
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import WebDriverException, NoSuchWindowException
from webdriver_manager.chrome import ChromeDriverManager

# =========================
# AYARLAR
# =========================
ANA_URL = "https://www.tjk.org/TR/yarissever/Info/Page/GunlukYarisProgrami"
EXCEL_NAME = "SON800_BUGUN_ILK_AT_STIL_FINAL_TEST.xlsx"
HEADLESS_CHROME = True
SELENIUM_BEKLEME_SN = 2

TURK_SEHIRLER = [
    "İstanbul", "Bursa", "İzmir", "Kocaeli", "Ankara",
    "Adana", "Şanlıurfa", "Diyarbakır", "Elazığ", "Antalya"
]

# MUTLAK GÜN: hedef = gerçek bugün + PROGRAM_GUN_OFFSET (derece/galop/pipeline ile aynı).
# Değeri TEK YERDEN ayarla: gun_ayar.py içindeki OFFSET. (0=bugün,1=yarın,2=...)
try:
    from gun_ayar import OFFSET as PROGRAM_GUN_OFFSET
except Exception:
    PROGRAM_GUN_OFFSET = 0
bugun = datetime.today()
try:
    from gun_ayar import hedef_gun as _hedef_gun
    program_tarihi = _hedef_gun()
except Exception:
    program_tarihi = bugun + timedelta(days=PROGRAM_GUN_OFFSET)
bir_yil_once = program_tarihi - timedelta(days=375)

# --- SAHA MEDYAN AYRIŞTIRICI + KALICI KÜTÜPHANE (opsiyonel) ---
# son800 zaten her referans koşunun sonuç sayfasını açıyor; o sayfadan HP/kilo
# medyanını da alıp kütüphaneye yazıyoruz. 0-183 gün derece'den hazır (kütüphane
# isabeti), 183-375 arası bu sayfadan bedava. Modül yoksa özellik sessiz kapanır.
SAHA_MEDYAN_KULLAN = True
try:
    import tjk_saha_medyan as SM
    import tjk_kutuphane as KUT
    _kut_conn = KUT.baglan() if SAHA_MEDYAN_KULLAN else None
    if _kut_conn is not None:
        print(f"[KÜTÜPHANE] açık — kayıtlı koşu sayısı: {KUT.sayac(_kut_conn)}")
except Exception as _sm_imp_e:
    print(f"[SAHA MEDYAN] modül yüklenemedi (görmezden geliniyor): {_sm_imp_e}")
    SM = None
    KUT = None
    _kut_conn = None

SESSION = requests.Session()
driver = None

# =========================
# GENEL TEMİZLİK
# =========================
def kolon_adi_norm(c):
    if isinstance(c, tuple):
        c = " ".join([str(x) for x in c if str(x).strip() and str(x).strip().lower() != "nan"])
    t = str(c).strip().upper()
    t = t.replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ü", "U").replace("Ö", "O").replace("Ç", "C")
    t = re.sub(r"\s+", " ", t)
    return t


def metin_norm_at(x):
    t = str(x or "").strip().upper()
    t = t.replace("İ", "I").replace("İ", "I")
    t = t.replace("Ş", "S").replace("Ğ", "G").replace("Ü", "U").replace("Ö", "O").replace("Ç", "C")
    t = re.sub(r"[^A-Z0-9]+", "", t)
    return t


def satir_kolon_getir(row, kolon_adi):
    if row is None:
        return ""
    if kolon_adi in row.index:
        return row[kolon_adi]
    hedef = kolon_adi_norm(kolon_adi)
    for c in row.index:
        if kolon_adi_norm(c) == hedef:
            return row[c]
    return ""


def sira_birinci_mi(value):
    if pd.isna(value):
        return False
    m = re.search(r"(\d+)", str(value).strip())
    return bool(m and int(m.group(1)) == 1)


def pist_durumu_temizle(x):
    x = str(x).strip()
    x = re.sub(r"\s+\d+([.,]\d+)?$", "", x)
    return x.strip()


def siklet_temizle(x):
    x = str(x).strip()
    x = re.sub(r"[^0-9,.]", "", x)
    if x.isdigit() and len(x) == 3:
        return x[:2] + "," + x[2]
    return x.replace(".", ",")


def mesafe_temizle(x):
    m = re.search(r"(\d{3,4})", str(x))
    return m.group(1) if m else str(x).strip()


def irk_getir(row):
    grup = str(satir_kolon_getir(row, "Grup")).upper()
    if grup.endswith("A"):
        return "Arap"
    if grup.endswith("İ") or grup.endswith("I"):
        return "İngiliz"
    return ""


def pist_ayir(row):
    pist = str(satir_kolon_getir(row, "Pist")).strip()
    pist_turu = pist.split(":")[0].strip()
    pist_durumu = ""
    if ":" in pist:
        pist_durumu = pist_durumu_temizle(pist.split(":", 1)[1])
    if pist_turu == "S":
        pist_turu = "Sentetik"
    elif pist_turu == "K":
        pist_turu = "Kum"
    elif pist_turu == "Ç":
        pist_turu = "Çim"
    return pist_turu, pist_durumu

# =========================
# SON 800 DÖNÜŞÜM
# =========================
def son800_derece_temizle(value):
    value = str(value or "").strip()
    if value == "" or value.lower() == "nan":
        return ""
    value = value.replace(":", ".")
    value = value.replace(".", "")
    value = re.sub(r"[^0-9]", "", value)
    if not value:
        return ""
    number = int(value)
    if 6000 <= number <= 19999:
        number -= 4000
    elif 20000 <= number <= 29999:
        number -= 8000
    elif 30000 <= number <= 39999:
        number -= 12000
    return number


def son800_aralik_parse(text):
    txt = str(text or "")
    txt = txt.replace("Son 800", "")
    txt = txt.replace("SON 800", "")
    txt = txt.replace(":", " ")
    m = re.search(r"(\d+[.:]\d+[.:]\d+)\s*-\s*(\d+[.:]\d+[.:]\d+)", txt)
    if not m:
        return "", "", ""
    bas = son800_derece_temizle(m.group(1))
    son = son800_derece_temizle(m.group(2))
    return bas, son, f"{m.group(1)}-{m.group(2)}"


def son800_html_icinden_cek(html):
    text = BeautifulSoup(str(html or ""), "html.parser").get_text(" ", strip=True)
    # örn: Son 800 :0.52.72-0.53.28
    m = re.search(r"Son\s*800\s*[:：]?\s*(\d+[.:]\d+[.:]\d+\s*-\s*\d+[.:]\d+[.:]\d+)", text, flags=re.I)
    if m:
        return son800_aralik_parse(m.group(1))
    return son800_aralik_parse(text)

# =========================
# FOTO/FİNİŞ SIRASI
# =========================
def foto_sirasi_html_icinden_cek(html, at_adi):
    """Sonuç tablosunda hedef atın S/No gibi bitiriş sırasını yakalamaya çalışır.
    Bulamazsa boş döner. Burada stil adı yok, sadece sayı.
    """
    hedef = metin_norm_at(at_adi)
    soup = BeautifulSoup(str(html or ""), "html.parser")
    for table in soup.find_all("table"):
        rows = []
        for tr in table.find_all("tr"):
            cells = tr.find_all(["th", "td"])
            vals = [" ".join(c.get_text(" ", strip=True).split()) for c in cells]
            if vals:
                rows.append(vals)
        if not rows:
            continue
        table_text = " ".join(" ".join(r) for r in rows)
        if hedef and hedef not in metin_norm_at(table_text):
            continue

        # Header varsa S / Derece / At İsmi kolonlarını kullan.
        header_idx = None
        s_idx = None
        at_idx = None
        for i, vals in enumerate(rows[:8]):
            norms = [re.sub(r"[^A-Z0-9]", "", kolon_adi_norm(v)) for v in vals]
            for j, n in enumerate(norms):
                if n in ["S", "SIRA", "DERECE"] and s_idx is None:
                    s_idx = j
                if n in ["ATISMI", "ATADI", "AT"] or "ATISMI" in n:
                    at_idx = j
            if s_idx is not None and at_idx is not None:
                header_idx = i
                break

        if header_idx is not None:
            for vals in rows[header_idx + 1:]:
                if len(vals) <= max(s_idx, at_idx):
                    continue
                if hedef in metin_norm_at(vals[at_idx]):
                    m = re.search(r"(\d+)", vals[s_idx])
                    if m:
                        return int(m.group(1))

        # Header yakalanmadıysa hedef at satırındaki ilk sayıyı sıra kabul etmeye çalış.
        for vals in rows:
            line = " ".join(vals)
            if hedef in metin_norm_at(line):
                # Genelde ilk hücre S olur.
                if vals:
                    m = re.search(r"(\d+)", vals[0])
                    if m:
                        return int(m.group(1))
    return ""

# =========================
# HTTP / SELENIUM
# =========================
def fetch_html_url(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/125 Safari/537.36",
        "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
        "Referer": ANA_URL,
    }
    r = SESSION.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    return r.text


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
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.page_load_strategy = "eager"
    d = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    d.set_page_load_timeout(60)
    return d


def guvenli_get(d, url, bekleme=SELENIUM_BEKLEME_SN):
    global driver
    for deneme in range(1, 4):
        try:
            if d is None:
                d = driver_baslat()
                driver = d
            d.get(url)
            time.sleep(bekleme)
            return d
        except (NoSuchWindowException, WebDriverException, Exception):
            try:
                d.quit()
            except Exception:
                pass
            d = driver_baslat()
            driver = d
            time.sleep(2)
    raise Exception(f"Sayfa açılamadı: {url}")

# =========================
# PROGRAMDAN İLK ATI AL
# =========================
def _sayfa_program_tarihi(driver):
    try:
        kaynak = driver.page_source or ""
    except Exception:
        kaynak = ""
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+\w+\s*-\s*Yarış Programı", kaynak)
    return m.group(1) if m else ""


def ileri_gun_git(driver, gun_sayisi):
    if gun_sayisi < 1:
        return driver
    print(f"\n{gun_sayisi} GÜN SONRAKİ PROGRAMA GEÇİLİYOR...")
    for adim in range(1, gun_sayisi + 1):
        tarih_once = _sayfa_program_tarihi(driver)
        tiklandi = False
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
                    tiklandi = True
                    break
            except Exception:
                continue
        if not tiklandi:
            try:
                tiklandi = driver.execute_script("""
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
                tiklandi = False
        if not tiklandi:
            raise Exception(f"İleri gün oku bulunamadı (adım {adim}/{gun_sayisi}).")
        for _ in range(30):
            time.sleep(1)
            tarih_sonra = _sayfa_program_tarihi(driver)
            if tarih_sonra and tarih_sonra != tarih_once:
                print(f"  geçildi: {tarih_once} -> {tarih_sonra}")
                break
            if tarih_sonra and not tarih_once:
                break
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
    """Programı MUTLAK hedef tarihe (datetime) götürür; pipeline/galop/derece ile aynı."""
    hedef = hedef_dt.strftime("%d/%m/%Y")
    print(f"\nHEDEF GÜN (mutlak): {hedef}")
    for _ in range(0, 8):
        simdi = _sayfa_program_tarihi(driver)
        print(f"  sayfa: {simdi if simdi else 'OKUNAMADI'} | hedef: {hedef}")
        if simdi == hedef:
            print(f"  hedefe ulaşıldı: {hedef}")
            return driver
        if not _ileri_tikla(driver):
            raise Exception(f"İleri gün oku bulunamadı; hedef {hedef} sayfasına gidilemedi.")
        for _w in range(1, 31):
            time.sleep(1)
            s = _sayfa_program_tarihi(driver)
            if s and s != simdi:
                break
    raise Exception(f"Hedef güne ({hedef}) ulaşılamadı; site farklı tarih gösteriyor.")


def sehir_linklerini_bul():
    global driver
    driver = guvenli_get(driver, ANA_URL, bekleme=5)
    driver = hedef_gune_git(driver, program_tarihi)
    links = driver.find_elements(By.TAG_NAME, "a")
    sehir_linkleri = []
    for link in links:
        try:
            href = link.get_attribute("href") or ""
            text = link.text.strip()
            if "GunlukYarisProgrami?SehirId=" not in href:
                continue
            for sehir in TURK_SEHIRLER:
                if sehir.lower() in text.lower():
                    temiz_url = href.split("&Era=")[0]  # BUGÜN. Era/offset ekleme yok.
                    if (sehir, temiz_url) not in sehir_linkleri:
                        sehir_linkleri.append((sehir, temiz_url))
        except Exception:
            pass
    return sehir_linkleri


def ilk_program_ati_bul():
    global driver
    sehirler = sehir_linklerini_bul()
    if not sehirler:
        raise Exception("Bugünkü programda şehir linki bulunamadı")

    # Sadece ilk şehir. Başka şehre geçme yok.
    sehir, url = sehirler[0]
    print(f"BUGÜN İLK ŞEHİR: {sehir} | {url}")
    driver = guvenli_get(driver, url, bekleme=4)

    links = driver.find_elements(By.TAG_NAME, "a")
    for a in links:
        href = a.get_attribute("href") or ""
        text = " ".join((a.text or "").split())
        if "AtKosuBilgileri" not in href:
            continue
        m = re.search(r"QueryParameter_AtId=(\d+)", href)
        if not m:
            continue
        at_id = m.group(1)
        at_adi = text.strip()
        if not at_adi:
            at_adi = f"AT_{at_id}"
        print(f"İLK AT KİLİTLENDİ: {at_adi} | {at_id}")
        return {"sehir": sehir, "at_adi": at_adi, "at_id": at_id, "href": href}

    raise Exception("İlk şehirde at linki bulunamadı")


def tum_program_atlari_bul():
    """TÜM Türkiye şehirlerindeki TÜM atları bulur (at_id'ye göre tekilleştirir)."""
    global driver
    sehirler = sehir_linklerini_bul()
    if not sehirler:
        raise Exception("Bugünkü programda şehir linki bulunamadı")
    atlar = []
    gorulen = set()
    for sehir, url in sehirler:
        print(f"\n===== ŞEHİR: {sehir} =====")
        try:
            driver = guvenli_get(driver, url, bekleme=4)
            links = driver.find_elements(By.TAG_NAME, "a")
        except Exception as e:
            print(f"  {sehir} açılamadı: {e}")
            continue
        say = 0
        for a in links:
            try:
                href = a.get_attribute("href") or ""
                text = " ".join((a.text or "").split())
            except Exception:
                continue
            if "AtKosuBilgileri" not in href:
                continue
            m = re.search(r"QueryParameter_AtId=(\d+)", href)
            if not m:
                continue
            at_id = m.group(1)
            if at_id in gorulen:
                continue
            gorulen.add(at_id)
            atlar.append({
                "sehir": sehir,
                "at_adi": text.strip() or f"AT_{at_id}",
                "at_id": at_id,
                "href": href,
            })
            say += 1
        print(f"  {sehir}: {say} at")
    print(f"\nTOPLAM AT: {len(atlar)}")
    return atlar

# =========================
# AT GEÇMİŞİ
# =========================
def at_url_adaylari(at_id, orijinal_href=""):
    adaylar = []
    if orijinal_href:
        adaylar.append(orijinal_href)
    base1 = "https://www.tjk.org/TR/YarisSever/Query/ConnectedPage/AtKosuBilgileri"
    base2 = "https://www.tjk.org/TR/yarissever/Query/ConnectedPage/AtKosuBilgileri"
    for base in [base1, base2]:
        for params in [
            {"1": "1", "QueryParameter_AtId": at_id, "Era": "today"},
            {"QueryParameter_AtId": at_id, "Era": "today"},
            {"1": "1", "QueryParameter_AtId": at_id},
            {"QueryParameter_AtId": at_id},
        ]:
            adaylar.append(base + "?" + urlencode(params))
    temiz = []
    for u in adaylar:
        if u and u not in temiz:
            temiz.append(u)
    return temiz


def at_gecmis_oku(at_id, at_adi, href):
    son_hata = None
    for url in at_url_adaylari(at_id, href):
        try:
            html = fetch_html_url(url)
            tables = pd.read_html(StringIO(html))
            print(f"AT GEÇMİŞİ OKUNDU: {at_adi} | {url}")
            return tables, url
        except Exception as e:
            son_hata = e
            print(f"AT GEÇMİŞİ OKUNAMADI, SONRAKİ URL: {url} | {e}")
    raise Exception(f"At geçmişi okunamadı: {at_adi} | {son_hata}")


def tarih_kolonu_bul(df):
    for c in df.columns:
        cn = kolon_adi_norm(c)
        if cn == "TARIH" or cn.endswith(" TARIH"):
            return c
    return None


def gecmis_tablo_sec(tables):
    for df in tables:
        cols = [kolon_adi_norm(c) for c in df.columns]
        if any(c == "TARIH" or c.endswith(" TARIH") for c in cols) and any(c == "S" for c in cols):
            return df.copy()
    # fallback: en çok kolonlu tablo
    if tables:
        return sorted(tables, key=lambda x: len(x.columns), reverse=True)[0].copy()
    raise Exception("Geçmiş tablo bulunamadı")


def ilk_birincilik_satiri_bul(df):
    tarih_col = tarih_kolonu_bul(df)
    if tarih_col is None:
        raise Exception("Tarih kolonu bulunamadı")
    tmp = df.copy()
    tmp["__TARIH__"] = pd.to_datetime(tmp[tarih_col], dayfirst=True, errors="coerce")
    tmp = tmp.dropna(subset=["__TARIH__"])
    tmp = tmp[(tmp["__TARIH__"] >= bir_yil_once) & (tmp["__TARIH__"] <= program_tarihi + timedelta(days=1))]
    tmp = tmp[tmp.apply(lambda r: sira_birinci_mi(satir_kolon_getir(r, "S")), axis=1)]
    if len(tmp) == 0:
        return None
    # TJK geçmişinde en yeni genelde üstte. Orijinal sıra korunur, ilk uygun satır alınır.
    return tmp.iloc[0]


def tum_birincilik_satirlari_bul(df):
    """Atın son 1 yılda 1. olduğu TÜM geçmiş koşu satırlarını döndürür (sıra korunur)."""
    tarih_col = tarih_kolonu_bul(df)
    if tarih_col is None:
        return []
    tmp = df.copy()
    tmp["__TARIH__"] = pd.to_datetime(tmp[tarih_col], dayfirst=True, errors="coerce")
    tmp = tmp.dropna(subset=["__TARIH__"])
    tmp = tmp[(tmp["__TARIH__"] >= bir_yil_once) & (tmp["__TARIH__"] <= program_tarihi + timedelta(days=1))]
    tmp = tmp[tmp.apply(lambda r: sira_birinci_mi(satir_kolon_getir(r, "S")), axis=1)]
    return [tmp.iloc[i] for i in range(len(tmp))]

# =========================
# SONUÇ SAYFASI LİNKİ
# =========================
def sonuc_linki_mi(href):
    h = str(href or "").upper()
    return "GUNLUKYARISSONUCLARI" in h or "YARISSONUCLARI" in h


def satirdan_sonuc_linki_bul(tr, tarih):
    adaylar = []
    for a in tr.find_elements(By.TAG_NAME, "a"):
        try:
            href = (a.get_attribute("href") or "").strip()
            txt = " ".join((a.text or "").split())
            if href and sonuc_linki_mi(href):
                puan = 10
                if tarih and tarih in txt:
                    puan += 3
                adaylar.append((puan, href))
        except Exception:
            pass
    if not adaylar:
        return ""
    return sorted(adaylar, key=lambda x: x[0], reverse=True)[0][1]


def gecmis_kosu_detay_html_tarihe_tikla(at_url, row, at_adi):
    global driver
    tarih = str(satir_kolon_getir(row, "Tarih")).strip()
    if not tarih:
        raise Exception("Geçmiş satırda tarih boş")

    driver = guvenli_get(driver, at_url, bekleme=3)
    satirlar = driver.find_elements(By.TAG_NAME, "tr")
    adaylar = []
    for tr in satirlar:
        try:
            txt = " ".join(tr.text.split())
            if tarih not in txt:
                continue
            href = satirdan_sonuc_linki_bul(tr, tarih)
            if href:
                adaylar.append((txt, href))
        except Exception:
            pass

    if not adaylar:
        raise Exception(f"Sonuç linki bulunamadı: {at_adi} | {tarih}")

    href = adaylar[0][1]
    print(f"SONUÇ HREF: {href}")
    driver = guvenli_get(driver, href, bekleme=5)
    html = driver.page_source or ""
    if len(html) < 1000:
        raise Exception("Sonuç sayfası HTML boş/eksik")
    return href, html

# =========================
# EXCEL SATIRI
# =========================
def kayit_olustur(at_bilgi, row, son800_bas="", son800_son="", son800_raw="", foto_sirasi="", durum="OK", hata="",
                  kosu_kimlik="", eski_hp="", saha_hp="", saha_kilo="", saha_hp_liste="", saha_kilo_liste="",
                  kaynak_url="", derece=""):
    at_adi = at_bilgi.get("at_adi", "")
    pist, pist_durumu = pist_ayir(row) if row is not None else ("", "")
    tarih_raw = satir_kolon_getir(row, "Tarih") if row is not None else ""
    yil = ""
    try:
        yil = pd.to_datetime(tarih_raw, dayfirst=True, errors="coerce").year
        if pd.isna(yil):
            yil = program_tarihi.year
    except Exception:
        yil = program_tarihi.year

    return {
        "AtKey": f"{at_adi}1",
        "Sıra": 1,
        "Yıl": yil,
        "Şehir": satir_kolon_getir(row, "Şehir") if row is not None else at_bilgi.get("sehir", ""),
        "Pist": pist,
        "Pist Durumu": pist_durumu,
        "Mesafe": mesafe_temizle(satir_kolon_getir(row, "Msf")) if row is not None else "",
        "Kilo": siklet_temizle(satir_kolon_getir(row, "Sıklet")) if row is not None else "",
        "At Adı": at_adi,
        "Irk": irk_getir(row) if row is not None else "",
        "Son800Baş": son800_bas,
        "Son800Son": son800_son,
        "Koşu Cinsi": satir_kolon_getir(row, "Kcins") if row is not None else "",
        "Stil": foto_sirasi,
        "Tarih": tarih_raw,
        "Son800 Ham": son800_raw,
        # --- 800 DOMİNANS için ek kolonlar ---
        "Koşu Kimliği": kosu_kimlik,
        "Kaynak URL": kaynak_url,
        "Derece": derece,
        "Eski HP": eski_hp,
        "Saha HP Medyan": saha_hp,
        "Saha Kilo Medyan": saha_kilo,
        "Medyan HP Listesi": saha_hp_liste,
        "Medyan Kilo Listesi": saha_kilo_liste,
        "Durum": durum,
        "Hata": hata,
    }


def saha_medyan_ve_eski_hp(html, row, at_adi, gecmis_df, detay_url):
    """Referans koşunun saha HP/kilo medyanı + listeleri (kütüphaneden ya da
    sayfadan) ve atın eski HP'si. BEST-EFFORT: hata olursa boş döner, son800'ü
    asla bozmaz."""
    sonuc = {"kimlik": "", "eski_hp": "", "saha_hp": "", "saha_kilo": "",
             "saha_hp_liste": "", "saha_kilo_liste": ""}

    # Eski HP: geçmiş tablosunda referans satırın bir altındaki (önceki) koşu.
    if SM is not None:
        try:
            sonuc["eski_hp"] = SM.eski_hp_bul(gecmis_df, row.name)
        except Exception:
            sonuc["eski_hp"] = ""

    if SM is None:
        return sonuc

    try:
        kimlik = SM.kosu_kimlik(row)
        sonuc["kimlik"] = kimlik
        hp_l = kilo_l = None
        hp_ref = kilo_ref = ""

        # 1) Kütüphane isabeti (0-183 gün derece'den hazır) -> sayfayı ayrıştırma.
        kayit_k = None
        if _kut_conn is not None and kimlik:
            kayit_k = KUT.medyan_getir(_kut_conn, kimlik)
        if kayit_k is not None and kayit_k.get("hp_medyan") is not None:
            hp_ref = kayit_k["hp_medyan"]
            kilo_ref = kayit_k["kilo_medyan"]
            hp_l = kayit_k["hp_liste"]
            kilo_l = kayit_k["kilo_liste"]
        else:
            # 2) Zaten açık olan sonuç sayfasından bedava ayrıştır + kütüphaneye yaz.
            hp_ref, kilo_ref, hp_l, kilo_l = SM.saha_medyan_cek(
                html, at_adi=at_adi, detay_url=detay_url, row=row)
            if _kut_conn is not None and kimlik and hp_l:
                KUT.medyan_yaz(_kut_conn, kimlik, hp_medyan=hp_ref, kilo_medyan=kilo_ref,
                               hp_liste=hp_l, kilo_liste=kilo_l, detay_url=detay_url, kaynak="800")

        sonuc["saha_hp"] = hp_ref
        sonuc["saha_kilo"] = kilo_ref
        if hp_l is not None:
            sonuc["saha_hp_liste"] = "-".join(SM.liste_deger_formatla(x) for x in hp_l)
        if kilo_l is not None:
            sonuc["saha_kilo_liste"] = "-".join(SM.liste_deger_formatla(x) for x in kilo_l)
    except Exception as e:
        print(f"  SAHA MEDYAN alınamadı (son800 etkilenmez): {at_adi} | {e}")

    return sonuc


def excel_yaz(kayitlar):
    df = pd.DataFrame(kayitlar)
    df = df.replace([pd.NA, float("inf"), float("-inf")], "").fillna("")
    df.to_excel(EXCEL_NAME, index=False)
    print(f"\nEXCEL OLUŞTU: {EXCEL_NAME}")

# =========================
# MAIN
# =========================
def main():
    """TAM MOD: tüm şehirler -> tüm atlar -> her atın son 1 yıldaki TÜM 1.lik
    koşuları -> her biri için Son 800 çek. Çıktı: EXCEL_NAME (Veri sayfası)."""
    global driver
    kayitlar = []
    basarisiz = []
    try:
        driver = driver_baslat()
        atlar = tum_program_atlari_bul()

        for i, at_bilgi in enumerate(atlar, 1):
            at_adi = at_bilgi["at_adi"]
            print(f"\n[{i}/{len(atlar)}] AT: {at_adi} ({at_bilgi['sehir']})")
            try:
                tables, at_url = at_gecmis_oku(at_bilgi["at_id"], at_adi, at_bilgi["href"])
                gecmis_df = gecmis_tablo_sec(tables)
                rows = tum_birincilik_satirlari_bul(gecmis_df)
                if not rows:
                    print("  son 1 yılda 1.lik satırı yok, atlanıyor")
                    continue
                for row in rows:
                    try:
                        detay_url, html = gecmis_kosu_detay_html_tarihe_tikla(at_url, row, at_adi)
                        son800_bas, son800_son, son800_raw = son800_html_icinden_cek(html)
                        foto_sirasi = ""   # STİL/foto sırası ÇEKİLMİYOR (vazgeçildi)
                        if son800_bas == "" or son800_son == "":
                            raise Exception("Son 800 değeri bulunamadı")
                        # 800 DOMİNANS için: saha HP/kilo medyanı + eski HP (best-effort)
                        ek = saha_medyan_ve_eski_hp(html, row, at_adi, gecmis_df, detay_url)
                        kayitlar.append(kayit_olustur(at_bilgi, row, son800_bas, son800_son,
                                                      son800_raw, foto_sirasi, durum="OK", hata="",
                                                      kosu_kimlik=ek["kimlik"], eski_hp=ek["eski_hp"],
                                                      saha_hp=ek["saha_hp"], saha_kilo=ek["saha_kilo"],
                                                      saha_hp_liste=ek["saha_hp_liste"],
                                                      saha_kilo_liste=ek["saha_kilo_liste"],
                                                      kaynak_url=detay_url,
                                                      derece=satir_kolon_getir(row, "Derece")))
                        print(f"  OK: {satir_kolon_getir(row,'Tarih')} | Son800={son800_bas}-{son800_son} | SahaHP={ek['saha_hp']} | EskiHP={ek['eski_hp']}")
                    except Exception as e:
                        kayitlar.append(kayit_olustur(at_bilgi, row, durum="HATA", hata=str(e)))
                        basarisiz.append({"At Adı": at_adi, "Tarih": str(satir_kolon_getir(row, "Tarih")), "Hata": str(e)})
                        print(f"  SATIR HATASI: {e}")
            except Exception as e:
                basarisiz.append({"At Adı": at_adi, "Tarih": "", "Hata": str(e)})
                print(f"  AT HATASI: {e}")

            # Her 50 kayıtta ara kayıt (uzun çalışmada veri kaybı olmasın)
            if kayitlar and len(kayitlar) % 50 == 0:
                excel_yaz(kayitlar)

        excel_yaz(kayitlar)
        print(f"\nTAMAMLANDI. Toplam kayıt: {len(kayitlar)} | başarısız: {len(basarisiz)}")
        if basarisiz:
            print(f"Başarısız örnek: {basarisiz[:3]}")

    except Exception as e:
        print("GENEL HATA:", e)
        try:
            excel_yaz(kayitlar)
        except Exception:
            pass
    finally:
        try:
            if driver is not None:
                driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
