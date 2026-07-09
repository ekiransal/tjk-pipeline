#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GEÇ ÇIKIŞ TARAMASI
==================
Yarın koşacak TÜM atların GERÇEK SON 3 koşusuna bakar (derece/son800 filtrelerinden
bağımsız — atın takvimdeki son 3 koşusu) ve sonuç sayfasındaki "G. Çık." sütunundan
kaç boy geç çıktığını okur. Problemli atları raporlar + web sayfası için JSON üretir.

Veri kaynağı:
  - At listesi + at geçmiş sayfası URL'si: derece çıktısı (TUM_ILLER...xlsx 'Kaynak URL')
  - Son 3 koşu: atın AtKosuBilgileri sayfası (filtresiz, tarih sırasına göre)
  - G. Çık.: her koşunun GunlukYarisSonuclari sayfası (HP/kilo çektiğimiz tablo)

Önbellek: tjk_kutuphane.db 'gec_cikis' tablosu — bir at+tarih bir kez okunur,
sonraki çalıştırmalarda sayfa AÇILMAZ. Boy '' = geç çıkış yok (kontrol edildi).

Çıktı:
  - Terminal raporu (problemli atlar)
  - web/gec_cikis.json  {at: {"problem": bool, "kosular": [{"tarih","boy"},...]}}
"""
import os
import re
import json
import time
import sqlite3
import datetime
import pandas as pd
import requests
from io import StringIO

INPUT_EXCEL = "TUM_ILLER_TUM_ATLAR_EKSIKSIZ_ILK4_SON6AY_ERTESI_GUN.xlsx"
DB = "tjk_kutuphane.db"
OUT_JSON = os.path.join("web", "gec_cikis.json")
SON_N = 3          # atın gerçek son N koşusu
BEKLE_SN = 1.0     # istekler arası nezaket beklemesi
TIMEOUT = 30

ANA_URL = "https://www.tjk.org/TR/yarissever/Info/Page/GunlukYarisProgrami"
UA = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"),
    "Referer": ANA_URL,
}
SESSION = requests.Session()

TARIH_RE = re.compile(r"\b(\d{2})[./](\d{2})[./](\d{4})\b")


def fetch(url):
    r = SESSION.get(url, headers=UA, timeout=TIMEOUT)
    r.raise_for_status()
    return r.text


# --- SONUÇ SAYFALARI JS İLE YÜKLENİYOR: requests tablo göremez, Chrome gerekir ---
_DRV = None

def _driver():
    global _DRV
    if _DRV is None:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
        opts = webdriver.ChromeOptions()
        opts.add_argument("--headless=new")
        opts.add_argument("--window-size=1600,1000")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--blink-settings=imagesEnabled=false")
        _DRV = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
        _DRV.set_page_load_timeout(60)
        print("  [CHROME] sonuç sayfaları için görünmez tarayıcı açıldı")
    return _DRV


def fetch_sonuc(url):
    """Sonuç sayfasını Chrome ile açar (JS tabloları yüklensin)."""
    d = _driver()
    d.get(url)
    time.sleep(2.5)
    return d.page_source


def temiz_at(s):
    """'TANJU BEY(7) KG DB SK' -> 'TANJU BEY'. Parantez/rozet ekleri atılır."""
    s = str(s or "")
    s = re.sub(r"\(.*?\)", " ", s)               # (7), (IRE) ...
    s = re.sub(r"\b(KG|DB|SK|GKR|K|G)\b", " ", s)  # rozetler
    s = re.sub(r"[^A-Za-zÇĞİÖŞÜçğıöşü0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip().upper()
    return s


def _tarih_parse(s):
    m = TARIH_RE.search(str(s or ""))
    if not m:
        return None
    try:
        return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except Exception:
        return None


def db_baglan():
    conn = sqlite3.connect(DB, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA busy_timeout=30000;")
    conn.execute("""CREATE TABLE IF NOT EXISTS gec_cikis (
        at TEXT, tarih TEXT, boy TEXT, kaynak_url TEXT, guncelleme TEXT,
        PRIMARY KEY (at, tarih))""")
    # v2: derece alanı eklendi + eski tablo (hep '' yazan sürümün kayıtları) BYPASS
    conn.execute("""CREATE TABLE IF NOT EXISTS gec_cikis2 (
        at TEXT, tarih TEXT, boy TEXT, derece TEXT, kaynak_url TEXT, guncelleme TEXT,
        PRIMARY KEY (at, tarih))""")
    conn.commit()
    return conn


def cache_al(conn, at, tarih):
    r = conn.execute("SELECT boy, derece FROM gec_cikis2 WHERE at=? AND tarih=?",
                     (at, tarih)).fetchone()
    if r is None:
        return None
    if (r[1] or "") == "?":
        return None          # belirsiz kayıt: JS sorunu dönemi -> yeniden oku
    return (r[0] or "", r[1] or "")


def cache_yaz(conn, at, tarih, boy, url, derece=""):
    try:
        conn.execute(
            "INSERT OR REPLACE INTO gec_cikis2 (at,tarih,boy,derece,kaynak_url,guncelleme) "
            "VALUES (?,?,?,?,?,?)",
            (at, tarih, boy or "", derece or "", url or "",
             datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
    except Exception as e:
        print(f"  [CACHE] yazılamadı ({e}) — devam")


def son_kosulari_bul(at_html, hedef_tarih, n=SON_N):
    """At geçmiş sayfasından SON n koşuyu (tarih + sonuç sayfası linki) çıkarır.
    HEDEF tarihten (bugünkü/yarınki koşu) ÖNCEKİ koşular alınır — filtre YOK."""
    kosular = []  # (tarih, sonuc_url)
    # satır satır: her <tr> içinde tarih + GunlukYarisSonuclari linki ara
    for tr in re.split(r"<tr[\s>]", at_html)[1:]:
        tr = tr.split("</tr>")[0]
        t = _tarih_parse(tr)
        if t is None or (hedef_tarih and t >= hedef_tarih):
            continue
        m = re.search(r'href="([^"]*GunlukYarisSonuclari[^"]*)"', tr)
        url = m.group(1).replace("&amp;", "&") if m else ""
        if url and url.startswith("/"):
            url = "https://www.tjk.org" + url
        kosular.append((t, url))
    kosular.sort(key=lambda x: x[0], reverse=True)
    return kosular[:n]


def gcik_kolon_bul(df):
    for c in df.columns:
        cs = re.sub(r"[^A-ZÇIİĞÖŞÜ]", "", str(c).upper().replace("İ", "I"))
        if "CIK" in cs or "ÇIK" in cs or cs in ("GCIK", "GC"):
            return c
    return None


_TANI_KALAN = 2   # ilk 2 çözümsüz sayfada kolonları terminale dök (teşhis)

def _hucre_str(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "-") else s

def sonuc_sayfasindan_oku(html, at_adi):
    """Sonuç sayfasından atın (G.Çık, Derece) ikilisini okur.
    Dönüş: (gcik|''|None, derece|''|None)  None = at/kolon bulunamadı (belirsiz)."""
    global _TANI_KALAN
    try:
        tablolar = pd.read_html(StringIO(html))
    except Exception:
        return None, None
    at_n = temiz_at(at_adi)
    for df in tablolar:
        gk = gcik_kolon_bul(df)
        at_kol = drc_kol = None
        for c in df.columns:
            cs = re.sub(r"[^A-ZÇIİĞÖŞÜ ]", "", str(c).upper().replace("İ", "I"))
            if at_kol is None and ("ISMI" in cs or "ISIM" in cs or cs.strip() in ("AT", "AT ADI")):
                at_kol = c
            if drc_kol is None and "DERECE" in cs:
                drc_kol = c
        if at_kol is None:
            continue
        for _, row in df.iterrows():
            hucre = temiz_at(row.get(at_kol, ""))
            if hucre and (hucre == at_n or hucre.startswith(at_n) or at_n.startswith(hucre)):
                gcik = _hucre_str(row.get(gk, "")) if gk is not None else None
                drc = _hucre_str(row.get(drc_kol, "")) if drc_kol is not None else None
                if (gk is None or drc_kol is None) and _TANI_KALAN > 0:
                    _TANI_KALAN -= 1
                    print(f"  [TANI] kolonlar: {[str(x)[:18] for x in df.columns][:14]}"
                          f" | G.Çık kolonu: {gk} | Derece kolonu: {drc_kol}")
                return gcik, drc
    if _TANI_KALAN > 0:
        _TANI_KALAN -= 1
        try:
            print(f"  [TANI] at bulunamadı ({at_n}); sayfadaki tablo kolonları:")
            for df in tablolar[:3]:
                print("        ", [str(x)[:18] for x in df.columns][:14])
        except Exception:
            pass
    return None, None


def main():
    print("=" * 60)
    print(f"GEÇ ÇIKIŞ TARAMASI — her atın GERÇEK son {SON_N} koşusu")
    print("=" * 60)
    df = pd.read_excel(INPUT_EXCEL)
    # at -> at geçmiş sayfası URL (Era paramı temizlenir; tam liste görünsün)
    atlar = {}
    for _, r in df.iterrows():
        ad = str(r.get("At Adı", "") or "").strip()
        url = str(r.get("Kaynak URL", "") or "")
        if ad and "AtId" in url and ad not in atlar:
            atlar[ad] = url
    print(f"Koşacak at: {len(atlar)}")

    # hedef tarih (yarınki koşu) — gun_ayar.py'den; yoksa bugün+1
    hedef = None
    try:
        s = open("gun_ayar.py", encoding="utf-8").read()
        m = re.search(r'(?m)^HEDEF_TARIH\s*=\s*"(\d{2}\.\d{2}\.\d{4})"', s)
        if m:
            hedef = _tarih_parse(m.group(1))
    except Exception:
        pass
    if hedef is None:
        hedef = datetime.date.today() + datetime.timedelta(days=1)
    print(f"Hedef tarih: {hedef.strftime('%d.%m.%Y')} (bu tarihten önceki son {SON_N} koşu)")

    conn = db_baglan()
    sonuc = {}
    sayf_cache = {}   # sonuç sayfası URL -> html (aynı gün sayfası tekrar inmesin)
    n_fetch = n_cache = n_hata = 0

    for i, (at, at_url) in enumerate(sorted(atlar.items()), 1):
        try:
            kosular = None
            # 1) önce SON_N koşunun hepsi cache'te mi? (at sayfasını hiç açmadan)
            #    cache anahtarı at+tarih olduğu için önce tarihleri bilmemiz lazım ->
            #    at sayfası yine de gerekir; ama sonuç sayfaları cache'ten gelir.
            html = fetch(at_url)
            time.sleep(BEKLE_SN)
            kosular = son_kosulari_bul(html, hedef, SON_N)
            if not kosular:
                print(f"[{i}/{len(atlar)}] {at}: geçmiş koşu bulunamadı")
                continue
            kayitlar = []
            for t, surl in kosular:
                ts = t.strftime("%d.%m.%Y")
                cv = cache_al(conn, temiz_at(at), ts)
                if cv is None:
                    if not surl:
                        boy, drc = "", ""
                    else:
                        if surl not in sayf_cache:
                            sayf_cache[surl] = fetch_sonuc(surl)
                            n_fetch += 1
                        g, d = sonuc_sayfasindan_oku(sayf_cache[surl], at)
                        boy = "" if g is None else g
                        drc = "?" if d is None else d      # ? = belirsiz (derecesiz SAYILMAZ)
                    cache_yaz(conn, temiz_at(at), ts, boy, surl, drc)
                else:
                    boy, drc = cv
                    n_cache += 1
                dsiz = (drc == "")                          # derece hücresi BOŞ -> derecesiz
                kayitlar.append({"tarih": ts, "boy": boy, "dsiz": dsiz})
            problemli = [k for k in kayitlar if k["boy"] or k.get("dsiz")]
            sonuc[temiz_at(at)] = {"problem": bool(problemli), "kosular": kayitlar}
            durum = " | ".join(f"{k['tarih']}:{k['boy'] or ('DRSZ' if k.get('dsiz') else '-')}" for k in kayitlar)
            isaret = "⚠ GEÇ ÇIKIŞ/DERECESİZ" if problemli else "ok"
            print(f"[{i}/{len(atlar)}] {at}: {durum}  [{isaret}]")
        except Exception as e:
            n_hata += 1
            print(f"[{i}/{len(atlar)}] {at}: HATA {str(e)[:90]}")

    os.makedirs("web", exist_ok=True)
    json.dump(sonuc, open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False)
    problemli = {a: v for a, v in sonuc.items() if v["problem"]}
    print("\n" + "=" * 60)
    print(f"BİTTİ — taranan at: {len(sonuc)} | problemli: {len(problemli)} | "
          f"sayfa indirme: {n_fetch} | önbellekten: {n_cache} | hata: {n_hata}")
    for a, v in problemli.items():
        d = ", ".join(f"{k['tarih']} ({k['boy'] or 'derecesiz'})" for k in v["kosular"]
                      if k["boy"] or k.get("dsiz"))
        print(f"  ⚠ {a}: {d}")
    print(f"JSON: {OUT_JSON}")
    conn.close()
    try:
        if _DRV is not None:
            _DRV.quit()
    except Exception:
        pass


if __name__ == "__main__":
    main()
