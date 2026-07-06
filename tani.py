#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TANI (debug) — TJK program sayfasının gerçek yapısını dosyaya döker.
Çalıştır:  python3 tani.py
Üretir (aynı klasöre):
  debug_ana.html        -> ana günlük program sayfası
  debug_sehir.html      -> ilk Türkiye şehrinin program sayfası
  debug_ozet.txt        -> şehir linkleri + bulunan tabloların başlık/şekil özeti
Sonra bu üç dosyayı bana ilet (ya da klasörde dururlar, ben okurum).
"""
import re, time, sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

ANA_URL = "https://www.tjk.org/TR/yarissever/Info/Page/GunlukYarisProgrami"
SEHIRLER = ["İstanbul","Bursa","İzmir","Kocaeli","Ankara","Adana","Şanlıurfa","Diyarbakır","Elazığ","Antalya"]

opts = webdriver.ChromeOptions()
opts.add_argument("--headless=new"); opts.add_argument("--window-size=1920,1080")
opts.add_argument("--disable-gpu"); opts.add_argument("--no-sandbox"); opts.add_argument("--disable-dev-shm-usage")
opts.page_load_strategy = "eager"
d = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
d.set_page_load_timeout(60)

ozet = []
def log(s):
    print(s); ozet.append(str(s))

try:
    d.get(ANA_URL); time.sleep(3)
    open("debug_ana.html","w",encoding="utf-8").write(d.page_source)
    log(f"ANA URL: {d.current_url}")

    # Şehir linkleri
    sehir_link = {}
    for a in d.find_elements(By.TAG_NAME,"a"):
        try:
            href=a.get_attribute("href") or ""; txt=(a.text or "").strip()
            if "SehirId=" in href and ("GunlukYarisProgrami" in href or "Yaris" in href):
                for s in SEHIRLER:
                    if s.lower() in txt.lower() and s not in sehir_link:
                        sehir_link[s]=href
        except Exception: pass
    log(f"Bulunan şehir linkleri ({len(sehir_link)}): {list(sehir_link.items())[:5]}")

    # Ana sayfadaki tüm SehirId linklerini de göster (eşleşme başarısızsa diye)
    tum_sehirid=[]
    for a in d.find_elements(By.TAG_NAME,"a"):
        try:
            href=a.get_attribute("href") or ""; txt=(a.text or "").strip()
            if "SehirId=" in href: tum_sehirid.append((txt[:30], href))
        except Exception: pass
    log(f"Ham SehirId link sayısı: {len(tum_sehirid)} | ilk 12:")
    for t in tum_sehirid[:12]: log(f"   {t}")

    # İlk şehri aç
    if sehir_link:
        sehir, url = list(sehir_link.items())[0]
    elif tum_sehirid:
        sehir, url = "ILK_HAM", tum_sehirid[0][1]
    else:
        sehir, url = None, None

    if url:
        log(f"\nAçılan şehir: {sehir} -> {url}")
        d.get(url); time.sleep(3)
        try:
            d.execute_script("window.scrollTo(0,document.body.scrollHeight);"); time.sleep(1)
            d.execute_script("window.scrollTo(0,0);"); time.sleep(1)
        except Exception: pass
        html=d.page_source
        open("debug_sehir.html","w",encoding="utf-8").write(html)
        log(f"şehir HTML uzunluk: {len(html)}")

        # 'Koşu' başlıkları
        basliklar=re.findall(r"\d{1,2}\.\s*Ko[şs]u[^\n<]{0,80}", html)
        log(f"Koşu başlığı sayısı: {len(basliklar)} | ilk 3: {basliklar[:3]}")

        # Tablolar
        try:
            tablolar=pd.read_html(html)
            log(f"\npd.read_html tablo sayısı: {len(tablolar)}")
            for i,df in enumerate(tablolar):
                cols=[str(c) for c in df.columns]
                log(f"  TABLO {i}: shape={df.shape} | kolonlar={cols[:14]}")
        except Exception as e:
            log(f"pd.read_html HATA: {e}")
    else:
        log("Hiç şehir linki bulunamadı.")
finally:
    try: d.quit()
    except Exception: pass
    open("debug_ozet.txt","w",encoding="utf-8").write("\n".join(ozet))
    print("\nBitti. Klasörde: debug_ana.html, debug_sehir.html, debug_ozet.txt")
