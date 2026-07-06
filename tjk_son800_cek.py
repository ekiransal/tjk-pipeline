#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJK SON800 İSTATİSTİK ÇEKİCİ  (kaçak/sprinter için 800-fark kaynağı)
===================================================================

Kaynak: https://www.tjk.org/TR/YarisSever/Query/Page/Son800Ist?QueryParameter_YIL=YYYY
Bu sayfada YIL filtresi URL'den ÇALIŞIR (Koşu Sorgulama'nın aksine).

Kolonlar (11):
  Yıl | Şehir Adı | Pist | Pist Durumu | Mesafe | Kilo | At İsmi | Irk |
  Son 800 | Son 800'e İlk Giren | Koşu Cinsi

fark = (Son 800'e İlk Giren) − (Son 800)   [derece_temizle tamsayı biçiminde]
  - fark 0..4  -> kaçak
  - fark > (ırk,zemin) 75p eşiği -> sprinter

Sağlamlık: satır sayımı ve toplu okuma JS ile (senin son800_2026.py'nin çalışan
mantığı + benim koşu scraperındaki JS okuma/None-güvenli okuma).

Kullanım:
  python3 tjk_son800_cek.py 2026 [son800_2026.xlsx]
"""

import sys
import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

BASE = "https://www.tjk.org/TR/YarisSever/Query/Page/Son800Ist"
MAX_TIK = 800

KOLONLAR = ["Yıl", "Şehir Adı", "Pist", "Pist Durumu", "Mesafe", "Kilo",
            "At İsmi", "Irk", "Son 800", "Son 800'e İlk Giren", "Koşu Cinsi"]


def derece_temizle(value):
    """'0.49.62' -> 4962 ; >6000 ise -4000 (senin kuralın). Geçersiz -> None."""
    v = str(value).strip().replace(".", "")
    if v == "" or not v.lstrip("-").isdigit():
        return None
    n = int(v)
    if n > 6000:
        n -= 4000
    return n


def _satir_sayisi(driver):
    return len(driver.find_elements(By.CSS_SELECTOR, "table tbody tr"))


def _daha_fazla_butonu(driver):
    for item in (driver.find_elements(By.TAG_NAME, "button")
                 + driver.find_elements(By.TAG_NAME, "a")):
        try:
            if "Daha Fazla Sonuç Göster" in (item.text or "").strip():
                return item
        except Exception:
            pass
    return None


_JS_TABLO = r"""
const tables = Array.from(document.querySelectorAll('table'));
if (!tables.length) return [];
tables.sort((a, b) => b.querySelectorAll('tr').length - a.querySelectorAll('tr').length);
const out = [];
const trs = tables[0].querySelectorAll('tr');
for (let i = 0; i < trs.length; i++) {
    const tds = trs[i].querySelectorAll('td');
    if (!tds.length) continue;                       // başlık satırı
    const vals = Array.from(tds).map(td => (td.innerText || '').trim());
    if (vals.every(v => v === '')) continue;
    out.push(vals);
}
return out;
"""


def _tablo_oku(driver):
    try:
        rows = driver.execute_script(_JS_TABLO)
    except Exception as e:
        print("  Tablo okuma uyarısı:", e)
        rows = []
    return [[str(x) for x in r] for r in (rows or [])]


def cek(yil="2026", cikti=None, headless=True):
    yil = str(yil).strip()
    cikti = cikti or f"son800_{yil}.xlsx"
    url = f"{BASE}?QueryParameter_YIL={yil}"

    opts = webdriver.ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    try:
        driver.get(url)
        time.sleep(10)
        print(f"{yil} Son800 sayfası açıldı, sayfalanıyor...")

        tik = 0
        while tik < MAX_TIK:
            before = _satir_sayisi(driver)
            more = _daha_fazla_butonu(driver)
            if more is None:
                print(f"Daha fazla sonuç kalmadı. Satır={before}")
                break
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", more)
            driver.execute_script("arguments[0].click();", more)
            tik += 1
            bas = time.time()
            buyudu = False
            while True:
                n2 = _satir_sayisi(driver)
                if n2 > before:
                    if tik % 5 == 0:
                        print(f"tık {tik}: satır {before}->{n2}")
                    buyudu = True
                    break
                if time.time() - bas > 90:
                    break
                time.sleep(0.4)
            if not buyudu:
                print(f"İlerlemedi (tık {tik}). Mevcutla devam. Satır={before}")
                break

        ham = _tablo_oku(driver)
        print("Ham satır:", len(ham))
        kayit = []
        for v in ham:
            def g(i):
                return v[i] if i < len(v) else ""
            kayit.append({
                "Yıl": g(0), "Şehir Adı": g(1), "Pist": g(2), "Pist Durumu": g(3),
                "Mesafe": g(4), "Kilo": g(5), "At İsmi": g(6), "Irk": g(7),
                "Son 800": g(8), "Son 800'e İlk Giren": g(9), "Koşu Cinsi": g(10),
            })
        df = pd.DataFrame(kayit, columns=KOLONLAR)
        # Sadece istenen yıl
        df = df[df["Yıl"].astype(str).str.strip().isin([yil, yil + ".0"])]
        # Son800 dolu
        df = df[df["Son 800"].astype(str).str.strip().ne("")]
        df = df[df["Son 800'e İlk Giren"].astype(str).str.strip().ne("")]
        # Sayısal + fark
        s = df["Son 800"].map(derece_temizle)
        ilk = df["Son 800'e İlk Giren"].map(derece_temizle)
        df["Son 800"] = s
        df["Son 800'e İlk Giren"] = ilk
        df["fark"] = ilk - s          # (İlk Giren) − (Son 800)
        df = df.dropna(subset=["Son 800", "Son 800'e İlk Giren", "fark"])
        df = df.drop_duplicates()
        df.to_excel(cikti, index=False)
        neg = int((df["fark"] < 0).sum())
        print(f"\nExcel oluştu: {cikti} | satır: {len(df)} | negatif fark: {neg} "
              f"(0 beklenir; çoksa formül yönü ters demektir)")
        return df
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    yil = sys.argv[1] if len(sys.argv) > 1 else "2026"
    out = sys.argv[2] if len(sys.argv) > 2 else f"son800_{yil}.xlsx"
    cek(yil, out, headless=False)
