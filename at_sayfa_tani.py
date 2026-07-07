#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AT SAYFASI TANI — "daha fazla göster" / eksik geçmiş kontrolü
=============================================================
Sorun: İlk 3 HP tablonun EN ALTINDAN (en eski koşular) alınıyor. Sayfa uzun
listede kesiliyorsa ilk 3 HP yanlış olur. Bu script requests'in sayfayı TAM mı
kesik mi aldığını ölçer ve sayfadaki "daha fazla" mekanizmasının izini arar.

Kullanım:
  python3 at_sayfa_tani.py                # derece excel'inden en çok koşulu atı seçer
  python3 at_sayfa_tani.py "AT ADI"       # belirli at
"""
import re
import sys
import pandas as pd
import requests
from io import StringIO

INPUT_EXCEL = "TUM_ILLER_TUM_ATLAR_EKSIKSIZ_ILK4_SON6AY_ERTESI_GUN.xlsx"
UA = {"User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                     "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"),
      "Referer": "https://www.tjk.org/TR/yarissever/Info/Page/GunlukYarisProgrami"}


def incele(url, etiket):
    try:
        r = requests.get(url, headers=UA, timeout=30)
        html = r.text
    except Exception as e:
        print(f"  [{etiket}] HATA: {e}")
        return None
    tarihler = re.findall(r"\b\d{2}\.\d{2}\.\d{4}\b", html)
    try:
        tablolar = pd.read_html(StringIO(html))
        n_tablo = len(tablolar)
        gecmis = tablolar[1] if n_tablo > 1 else (tablolar[0] if n_tablo else None)
        n_satir = 0 if gecmis is None else len(gecmis)
    except Exception:
        n_tablo, n_satir = 0, 0
    enes = min(tarihler) if tarihler else "-"
    yeni = max(tarihler) if tarihler else "-"
    # tarih string min/max yanlış sıralar; sadece bilgi amaçlı. Doğrusu:
    def _k(t):
        g, a, y = t.split(".")
        return y + a + g
    if tarihler:
        enes = min(tarihler, key=_k)
        yeni = max(tarihler, key=_k)
    print(f"  [{etiket}] HTTP={r.status_code} boyut={len(html)} | tablo={n_tablo} "
          f"| geçmiş satır={n_satir} | tarih adedi={len(tarihler)} | en eski={enes} | en yeni={yeni}")
    return html, n_satir


def ipuclari(html):
    print("\n  --- 'daha fazla' / sayfalama ipuçları ---")
    bulundu = False
    for kalip in ["daha fazla", "dahafazla", "load", "more", "page", "sayfa",
                  "xcount", "count", "offset", "scroll", "lazy"]:
        for m in re.finditer(kalip, html, re.I):
            s = max(0, m.start() - 80)
            parca = html[s:m.end() + 120].replace("\n", " ")
            parca = re.sub(r"\s+", " ", parca)
            if any(x in parca.lower() for x in ["href", "onclick", "data-", "button", "ajax", "url", "query"]):
                print(f"   [{kalip}] ...{parca[:200]}...")
                bulundu = True
                break  # her kalıptan 1 örnek yeter
    if not bulundu:
        print("   (belirgin bir buton/ajax izi yok)")


def main():
    df = pd.read_excel(INPUT_EXCEL)
    if len(sys.argv) > 1:
        ad = sys.argv[1].strip().upper()
        sec = df[df["At Adı"].astype(str).str.upper().str.strip() == ad]
        if len(sec) == 0:
            raise SystemExit(f"At bulunamadı: {ad}")
    else:
        # en çok geçmiş koşusu olan at (satır sayısına göre)
        say = df.groupby("At Adı").size().sort_values(ascending=False)
        ad = say.index[0]
        sec = df[df["At Adı"] == ad]
        print(f"Seçilen at (en çok koşulu): {ad} ({say.iloc[0]} satır derece'de)")
    url = str(sec.iloc[0]["Kaynak URL"])
    print(f"At: {ad}\nURL: {url}\n")

    print("1) Farklı Era değerleriyle satır sayısı karşılaştırması:")
    base = url.split("&Era=")[0]
    sonuc = {}
    ana_html = None
    for era in ["today", "tomorrow", "lastYear", "last6Month", ""]:
        u = base + (f"&Era={era}" if era else "")
        cev = incele(u, f"Era={era or 'YOK'}")
        if cev:
            sonuc[era] = cev[1]
            if era == "today":
                ana_html = cev[0]
    print("\n2) Satır sayıları:", sonuc)
    print("   -> Eğer sayılar AYNI ve atın gerçek koşu sayısından AZ ise sayfa KESİK geliyor demektir.")
    print("   -> Tarayıcıda aynı sayfayı açıp koşu satırlarını sayarak karşılaştır.")
    if ana_html:
        ipuclari(ana_html)


if __name__ == "__main__":
    main()
