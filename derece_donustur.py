#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Veri -> derece dönüştürücü
==========================

Geçmiş scraper'ının (ERTESI_GUN / TUM_ILLER) ürettiği Excel'in "Veri" sayfasını,
ana çalışma kitabındaki 'derece' sayfasının formatına otomatik çevirir.

Elle yaptığın iş buydu:
  - sütunları yeniden sıralamak,
  - her at için 1,2,3... "say" (rep) eklemek,
  - başına Anahtar (At adı + rep) koymak.

Artık otomatik. Anahtar, Sayfa2 ile birebir uyumlu olsun diye aynı clean_horse_name
ile üretilir.

Standalone kullanım:
    python3 derece_donustur.py  scraper_ciktisi.xlsx  [cikti_derece.xlsx]
"""

import sys
import re
import pandas as pd
from openpyxl import load_workbook, Workbook

from tjk_donustur import clean_horse_name


# Stil (%45 Seyir Sırası) sayısını GÖRSEL üçgene çevirir.
# Dolu = ▶ (kaçak/önde en sağda), boş = ▷. 4 üçgen.
#   1 (kaçak) -> ▷▷▷▶   |   2-3 -> ▷▷▶▷   |   4-6 -> ▷▶▷▷   |   7+ -> ▶▷▷▷
def stil_ucgen(stil):
    m = re.search(r"\d+", str(stil or "").strip())
    if not m:
        return ""                      # sayı yoksa boş bırak
    n = int(m.group())
    if n <= 0:
        return ""
    if n == 1:
        idx = 3                         # en sağ
    elif n <= 3:
        idx = 2
    elif n <= 6:
        idx = 1
    else:
        idx = 0                         # en sol
    return "".join("▶" if i == idx else "▷" for i in range(4))


# 'derece' sayfasının C..AA kolon sırası (A=Anahtar, B=rep öne eklenir)
DERECE_KOLONLARI = [
    "At Adı", "Sıklet", "Irk", "Cinsiyet", "Tarih", "Şehir", "Pist",
    "Pist Durumu", "Mesafe", "Derece", "Koşu Cinsi", "Sıra", "HP",
    "Atın İlk 3 HP", "Atın İlk 3 HP Koşu Cinsleri", "İlk 3 HP Ham",
    "Koşu HP Medyan", "Koşu HP Dolu Atlar Kilo Medyan",
    "Medyan HP Listesi", "Medyan Kilo Listesi", "HP-Kilo Eşleşme Listesi",
    "Koşu Detay URL", "Kaynak URL", "Satır Durumu", "Hata",
]

# 28. sütun (AB) = Stil. yeni yer BY bunu çeker (VLOOKUP derece!28).
# Video fazı (Veri_Stil_Ekli) bunu "%45 Seyir Sırası" adıyla üretir.
STIL_KOLON_ADAYLARI = ["%45 Seyir Sırası", "Seyir Sırası", "Stil", "%45 Seyir Sirasi"]

# Çıktı başlığı (A..AB = 28 kolon) + AC (29.) = Stil Üçgen (görsel)
DERECE_BASLIK = ["Anahtar", "rep"] + DERECE_KOLONLARI + ["Stil", "Stil Üçgen"]


def _get(d, ad):
    """Dict/Series'ten kolonu büyük/küçük harf ve boşluk toleranslı al."""
    if ad in d:
        return d[ad]
    hedef = str(ad).strip().upper()
    for k in d.keys():
        if str(k).strip().upper() == hedef:
            return d[k]
    return ""


def veri_to_derece_rows(veri_kayitlari):
    """veri_kayitlari: list[dict]  (Veri sayfasının satırları, kolon adı -> değer).
    Döndürür: list[list]  -> 'derece' satırları (A..AA = 27 kolon), başlıksız.
    Satır sırası korunur; her at için rep 1,2,3... atanır."""
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
        for kol in DERECE_KOLONLARI:
            v = _get(d, kol)
            satir.append("" if v is None else v)
        # AB (28.) = Stil
        stil = ""
        for ad in STIL_KOLON_ADAYLARI:
            v = _get(d, ad)
            if v is not None and str(v).strip() != "":
                stil = v
                break
        satir.append(stil)
        satir.append(stil_ucgen(stil))          # AC (29.) = görsel üçgen
        out.append(satir)
    return out


def _veri_df_oku(path):
    xls = pd.ExcelFile(path)
    # Stil'i içeren "Veri_Stil_Ekli" varsa onu tercih et; yoksa "Veri".
    if "Veri_Stil_Ekli" in xls.sheet_names:
        sheet = "Veri_Stil_Ekli"
    elif "Veri" in xls.sheet_names:
        sheet = "Veri"
    else:
        sheet = xls.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet)
    return df


def veri_dosyasindan_derece(path):
    """Scraper çıktısı Excel -> derece satırları (list[list], başlıksız)."""
    df = _veri_df_oku(path)
    kayitlar = df.to_dict(orient="records")
    return veri_to_derece_rows(kayitlar)


def main():
    if len(sys.argv) < 2:
        print("Kullanım: python3 derece_donustur.py scraper_ciktisi.xlsx [cikti.xlsx]")
        raise SystemExit(1)
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) >= 3 else "derece_cikti.xlsx"
    rows = veri_dosyasindan_derece(inp)
    wb = Workbook()
    ws = wb.active
    ws.title = "derece"
    ws.append(DERECE_BASLIK)
    for r in rows:
        ws.append(["" if v is None else v for v in r])
    wb.save(out)
    print(f"derece satırı: {len(rows)} -> {out}")


if __name__ == "__main__":
    main()
