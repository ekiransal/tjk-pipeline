#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJK KOŞU SORGULAMA ÇEKİCİ  (orjin kütüphanesi ana kaynağı)
==========================================================

Kaynak: https://www.tjk.org/TR/YarisSever/Query/Page/KosuSorgulama
Her satır bir koşunun BİRİNCİSİ (kazanan).

KESİN TESPİT (gözlemle doğrulandı):
  - URL gün filtresi (?QueryParameter_Tarih=) ÇALIŞMIYOR.
  - "Daha Fazla Sonuç Göster" ÇALIŞIYOR ve sayfa EN YENİ KOŞU EN ÜSTTE sıralı.
  => Doğru yöntem: en yeniden başla, "Daha Fazla" ile geriye in, BAŞLANGIÇ
     tarihini geçince DUR. Sonra [başlangıç, bitiş] aralığına süz.

Sağlamlık:
  - Satır sayımı ve en-eski-tarih TEK JS çağrısıyla, sadece TARİH içeren
    satırlardan okunur (çöp/footer satırı sayılmaz).
  - Güvenlik freni: en fazla MAX_TIK "Daha Fazla" basışı (sonsuz inişi önler).

Kullanım:
  python3 tjk_kosu_sorgulama.py 01.05.2026 02.07.2026 [cikti.xlsx]
"""

import sys
import time
import re
from datetime import date
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


URL = "https://www.tjk.org/TR/YarisSever/Query/Page/KosuSorgulama"
MAX_TIK = 600          # güvenlik freni (en fazla bu kadar "Daha Fazla")

CIKTI_KOLONLARI = [
    "tarih", "şehir", "koşu", "grup", "ırk", "cins", "apr", "mesafe",
    "zemin", "kilo", "baba", "ana", "kazanç", "at adı", "yaş", "derece", "h_puani",
]
# 0 Tarih|1 Şehir|2 Koşu|3 Grup|4 Koşu Cinsi|5 Apr|6 Mesafe|7 Pist|8 Sıklet|
# 9 Orijin(Baba-Anne)|10 İkramiye|11 Birinci|12 Yaş|13 Derece|14 H.Puanı


def irk_ayir(grup):
    tl = str(grup or "").lower()
    if "arap" in tl:
        return "Arap"
    if "ingiliz" in tl or "i̇ngiliz" in tl:
        return "İngiliz"
    return ""


def baba_ana_ayir(orijin):
    s = str(orijin or "").replace("–", "-").replace("—", "-").strip()
    if " - " in s:
        b, a = s.split(" - ", 1)
    elif "-" in s:
        b, a = s.split("-", 1)
    else:
        b, a = s, ""
    return b.strip(), a.strip()


def _tarih_parse(s):
    m = re.search(r"(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})", str(s or "").strip())
    if not m:
        return None
    g, ay, yil = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return date(yil, ay, g)
    except Exception:
        return None


# --- TEK JS çağrısı: tarih satır sayısı + en yeni/en eski tarih metni ---
_JS_BILGI = r"""
const re = /\d{1,2}[.\/]\d{1,2}[.\/]\d{4}/;
const rows = document.querySelectorAll('table tbody tr');
let ilk = '', son = '', n = 0;
for (const r of rows) {
    const c = r.querySelector('td');
    if (!c) continue;
    const m = (c.innerText || '').trim().match(re);
    if (!m) continue;               // sadece TARİH satırları
    if (n === 0) ilk = m[0];        // en üst = en yeni
    son = m[0];                     // en alt = en eski
    n++;
}
return {n: n, ilk: ilk, son: son};
"""


def _bilgi(driver):
    try:
        r = driver.execute_script(_JS_BILGI)
    except Exception:
        r = {"n": 0, "ilk": "", "son": ""}
    return r.get("n", 0), _tarih_parse(r.get("ilk", "")), _tarih_parse(r.get("son", ""))


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
const re = /\d{1,2}[.\/]\d{1,2}[.\/]\d{4}/;
const out = [];
for (const tr of tables[0].querySelectorAll('tr')) {
    const tds = tr.querySelectorAll('td');
    if (!tds.length) continue;
    const vals = Array.from(tds).map(td => (td.innerText || '').trim());
    if (!re.test(vals[0] || '')) continue;   // sadece TARİH satırları; çöp/footer elenir
    out.push(vals);
}
return out;
"""


def _tablo_satirlari(driver):
    """SON okuma: ana tablodaki geçerli satırlar — TEK JS çağrısı (hızlı, None-hatasız)."""
    try:
        rows = driver.execute_script(_JS_TABLO)
    except Exception as e:
        print(f"  Tablo okuma uyarısı: {e}")
        rows = []
    return [[str(x) for x in r] for r in (rows or [])]


def sayfala(driver, db, dbit, bekleme_max=90):
    driver.get(URL)
    time.sleep(10)

    tik = 0
    while True:
        n, en_yeni, en_eski = _bilgi(driver)

        # Başlangıç tarihini geçtik mi? -> yeter.
        if en_eski is not None and en_eski < db:
            print(f"Aralık kapsandı: en eski yüklü {en_eski} < başlangıç {db}. Satır={n}")
            break
        if tik >= MAX_TIK:
            print(f"Güvenlik freni: {MAX_TIK} tık. En eski {en_eski}. Satır={n}")
            break

        more = _daha_fazla_butonu(driver)
        if more is None:
            print(f"Daha fazla buton yok. En eski {en_eski}. Satır={n}")
            break

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", more)
        driver.execute_script("arguments[0].click();", more)
        tik += 1

        bas = time.time()
        buyudu = False
        while True:
            n2, _, en_eski2 = _bilgi(driver)
            if n2 > n:
                buyudu = True
                if tik % 5 == 0 or en_eski2 != en_eski:
                    print(f"tık {tik}: satır {n}->{n2}  en eski={en_eski2}")
                break
            if time.time() - bas > bekleme_max:
                break
            time.sleep(0.4)

        if not buyudu:
            print(f"İlerlemedi (tık {tik}). Mevcutla devam. En eski {en_eski}, satır {n}.")
            break

    return _tablo_satirlari(driver)


def ham_satiri_kutuphaneye(vals):
    def g(i):
        return vals[i] if i < len(vals) else ""
    baba, ana = baba_ana_ayir(g(9))
    m = re.search(r"(\d{3,4})", str(g(6)))
    return {
        "tarih": g(0), "şehir": g(1), "koşu": g(2), "grup": g(3),
        "ırk": irk_ayir(g(3)), "cins": g(4), "apr": g(5),
        "mesafe": int(m.group(1)) if m else "", "zemin": g(7), "kilo": g(8),
        "baba": baba, "ana": ana, "kazanç": g(10),
        "at adı": g(11), "yaş": g(12), "derece": g(13), "h_puani": g(14),
    }


def cek(tarih_bas, tarih_bit, cikti="kosu_sorgulama_yeni.xlsx", headless=True):
    db = _tarih_parse(tarih_bas)
    dbit = _tarih_parse(tarih_bit)
    if db is None or dbit is None:
        raise SystemExit(f"Tarih GG.AA.YYYY olmalı. Alınan: {tarih_bas} / {tarih_bit}")
    if db > dbit:
        db, dbit = dbit, db
    print(f"Hedef aralık: {db} - {dbit}")

    opts = webdriver.ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    try:
        ham = sayfala(driver, db, dbit)
        # Aralığa süz
        suz = [v for v in ham if (_tarih_parse(v[0]) is not None
                                  and db <= _tarih_parse(v[0]) <= dbit)]
        kayitlar = [ham_satiri_kutuphaneye(v) for v in suz]
        df = pd.DataFrame(kayitlar, columns=CIKTI_KOLONLARI)
        df = df[df["tarih"].astype(str).str.strip().ne("")]
        aralik_ham = len(df)
        # Koşu kimliğine göre mükerrer temizliği (sayfalama tekrarları / bir koşu = bir galip)
        df = df.drop_duplicates(subset=["tarih", "şehir", "koşu", "at adı"], keep="first")
        df = df.reset_index(drop=True)
        df.to_excel(cikti, index=False)
        print(f"\nExcel oluştu: {cikti}")
        print(f"  toplam yüklü: {len(ham)} | aralıkta (ham): {aralik_ham} | "
              f"mükerrer sonrası: {len(df)}")
        return df
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Kullanım: python3 tjk_kosu_sorgulama.py GG.AA.YYYY(baş) GG.AA.YYYY(bitiş) [cikti.xlsx]")
        raise SystemExit(1)
    tb = sys.argv[1]
    te = sys.argv[2]
    out = sys.argv[3] if len(sys.argv) >= 4 else "kosu_sorgulama_yeni.xlsx"
    cek(tb, te, out, headless=False)
