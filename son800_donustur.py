#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
son800 Veri -> '800' dönüştürücü
================================

son800 scraper'ının ürettiği "Veri" sayfasını ana çalışma kitabındaki '800'
sayfasının formatına otomatik çevirir (sütun sırası + Anahtar/rep).

Elle yaptığın "tamamen kopyalayıp bu hale getirip yapıştırma" işini bitirir.
Anahtar, Sayfa2 ile birebir uyumlu olsun diye aynı clean_horse_name kullanır.

Standalone:
    python3 son800_donustur.py  son800_ciktisi.xlsx  [cikti_800.xlsx]
"""

import sys
import pandas as pd
from openpyxl import Workbook

from tjk_donustur import clean_horse_name


# '800' sayfasının C..M kolon sırası (A=Anahtar, B=rep öne eklenir)
SEKIZYUZ_KOLONLARI = [
    # 1. kolon artık YIL değil TAM TARİH (ham 800 verisinde 'Tarih' = 19.05.2026).
    "Tarih", "Şehir", "Pist", "Pist Durumu", "Mesafe", "Kilo",
    "At Adı", "Irk", "Son800Baş", "Son800Son", "Koşu Cinsi",
]

# 14. sütun (N) = Stil. yeni yer'e yeni sütun olarak (800 tarafı) eklenir.
# son800 scraper bunu "Stil" (foto/seyir sırası) adıyla üretir.
STIL_KOLON_ADAYLARI = ["Stil", "%45 Seyir Sırası", "Seyir Sırası", "Foto Sırası"]

SEKIZYUZ_BASLIK = ["Anahtar", "rep"] + SEKIZYUZ_KOLONLARI + ["Stil"]


def _get(d, ad):
    if ad in d:
        return d[ad]
    hedef = str(ad).strip().upper()
    for k in d.keys():
        if str(k).strip().upper() == hedef:
            return d[k]
    return ""


def veri_to_800_rows(veri_kayitlari):
    """veri_kayitlari: list[dict] (son800 Veri satırları).
    Döndürür: list[list] -> '800' satırları (A..M = 13 kolon), başlıksız.
    Her at için rep 1,2,3... atanır (satır sırası korunur)."""
    sayac = {}
    out = []
    for d in veri_kayitlari:
        atadi = _get(d, "At Adı")
        if atadi is None or str(atadi).strip() == "":
            continue
        clean = clean_horse_name(str(atadi))
        if clean == "":
            continue
        sayac[clean] = sayac.get(clean, 0) + 1
        rep = sayac[clean]
        satir = [f"{clean}{rep}", rep]
        for kol in SEKIZYUZ_KOLONLARI:
            v = _get(d, kol)
            satir.append("" if v is None else v)
        # N (14.) = Stil
        stil = ""
        for ad in STIL_KOLON_ADAYLARI:
            v = _get(d, ad)
            if v is not None and str(v).strip() != "":
                stil = v
                break
        satir.append(stil)
        out.append(satir)
    return out


def veri_dosyasindan_800(path):
    xls = pd.ExcelFile(path)
    sheet = "Veri" if "Veri" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet)
    return veri_to_800_rows(df.to_dict(orient="records"))


def main():
    if len(sys.argv) < 2:
        print("Kullanım: python3 son800_donustur.py son800_ciktisi.xlsx [cikti.xlsx]")
        raise SystemExit(1)
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) >= 3 else "800_cikti.xlsx"
    rows = veri_dosyasindan_800(inp)
    wb = Workbook()
    ws = wb.active
    ws.title = "800"
    ws.append(SEKIZYUZ_BASLIK)
    for r in rows:
        ws.append(["" if v is None else v for v in r])
    wb.save(out)
    print(f"800 satırı: {len(rows)} -> {out}")


if __name__ == "__main__":
    main()
