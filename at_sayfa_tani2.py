#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AT SAYFASI TANI v2 — 'Daha Fazla Sonuç Göster' formunu çözer ve DEVAMINI ÇEKMEYİ DENER
Kullanım:  python3 at_sayfa_tani2.py "FORLORN"
"""
import re
import sys
import pandas as pd
import requests
from io import StringIO
from urllib.parse import urlencode

INPUT_EXCEL = "TUM_ILLER_TUM_ATLAR_EKSIKSIZ_ILK4_SON6AY_ERTESI_GUN.xlsx"
UA = {"User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                     "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"),
      "Referer": "https://www.tjk.org/TR/yarissever/Info/Page/GunlukYarisProgrami"}
S = requests.Session()


def tarih_say(html):
    ts = re.findall(r"\b\d{2}\.\d{2}\.\d{4}\b", html)
    def _k(t):
        g, a, y = t.split(".")
        return y + a + g
    return len(ts), (min(ts, key=_k) if ts else "-"), (max(ts, key=_k) if ts else "-")


def form_coz(html):
    """show-more butonunu içeren formu bul: action + tüm input name/value'ları."""
    for m in re.finditer(r"<form\b(.*?)</form>", html, re.S | re.I):
        blok = m.group(0)
        if "show-more" not in blok:
            continue
        act = re.search(r'action="([^"]*)"', blok)
        action = (act.group(1) if act else "").replace("&amp;", "&")
        inputs = {}
        for im in re.finditer(r"<input\b[^>]*>", blok, re.I):
            tag = im.group(0)
            n = re.search(r'name="([^"]*)"', tag)
            v = re.search(r'value="([^"]*)"', tag)
            if n:
                inputs[n.group(1)] = (v.group(1) if v else "")
        return action, inputs, blok[:1200]
    return None, {}, ""


def main():
    ad = sys.argv[1].strip().upper() if len(sys.argv) > 1 else None
    df = pd.read_excel(INPUT_EXCEL)
    sec = df[df["At Adı"].astype(str).str.upper().str.strip() == ad] if ad else df
    if len(sec) == 0:
        raise SystemExit(f"At bulunamadı: {ad}")
    url = str(sec.iloc[0]["Kaynak URL"])
    print(f"At: {ad}\nURL: {url}\n")

    r = S.get(url, headers=UA, timeout=30)
    html = r.text
    n, eski, yeni = tarih_say(html)
    top = re.search(r"Toplam\s+(\d+)\s+sonuçtan\s+(\d+)", html)
    print(f"[SAYFA 1] HTTP={r.status_code} | tarih={n} | en eski={eski} | en yeni={yeni}"
          + (f" | TOPLAM={top.group(1)} GÖSTERİLEN={top.group(2)}" if top else " | 'Toplam X' ibaresi yok (tam liste olabilir)"))

    action, inputs, ham = form_coz(html)
    if not action and not inputs:
        print("show-more formu bulunamadı — sayfa tam olabilir.")
        return
    print(f"\nFORM action: {action}")
    print(f"FORM inputlar: {inputs}")
    print(f"FORM ham (ilk 1200): {ham}\n")

    # DEVAMINI ÇEKMEYİ DENE: action + inputlar GET
    if action.startswith("/"):
        action = "https://www.tjk.org" + action
    devam_url = action + ("&" if "?" in action else "?") + urlencode(inputs)
    print(f"[DENEME] devam URL: {devam_url[:160]}...")
    r2 = S.get(devam_url, headers=UA, timeout=30)
    n2, eski2, yeni2 = tarih_say(r2.text)
    print(f"[SAYFA 2] HTTP={r2.status_code} | boyut={len(r2.text)} | tarih={n2} | en eski={eski2} | en yeni={yeni2}")
    # 2. sayfada da form var mı (3. sayfa)?
    a3, i3, _ = form_coz(r2.text)
    print(f"[SAYFA 2] show-more devamı: {'VAR -> ' + str(i3) if i3 else 'yok (liste bitti)'}")


if __name__ == "__main__":
    main()
