#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FARK EŞLEŞTİRİCİ  (Son800 fark -> Koşu Sorgulama orijin satırlarına)
==================================================================

Son800Ist'te baba/ana YOK; Koşu Sorgulama'da 800-fark YOK. İkisi de her satırda
BİR KAZANANI (1.) gösterir. Ortak alanlardan eşleştirip fark'ı orijin satırına
ekleriz:  anahtar = (At | Şehir | Mesafe | Zemin)  [+ Kilo tiebreak, sıra hizası].

Çıktı: Koşu Sorgulama excel'ine 'fark' kolonu eklenmiş yeni dosya.
Kütüphane yükleyici (ham_excel_yukle) 'fark' kolonunu otomatik tanır.

Kullanım:
  python3 tjk_fark_esle.py kosu_2026_tam.xlsx son800_2026.xlsx [cikti.xlsx]
"""

import sys
import re
from collections import defaultdict
import pandas as pd


def _norm(s):
    t = str(s or "").strip().upper()
    t = t.replace("İ", "I").replace("İ", "I")
    return re.sub(r"\s+", " ", t)


def _mesafe(s):
    """Metre olarak normalize eder. son800 KM verir (2.1 -> 2100),
    koşu METRE verir (1200 -> 1200)."""
    t = str(s or "").replace(",", ".").strip()
    m = re.search(r"\d+(?:\.\d+)?", t)
    if not m:
        return ""
    v = float(m.group(0))
    if v < 100:            # km biçimi (ör. 2.1) -> metre
        v *= 1000
    return str(int(round(v)))


def _zemin(s):
    t = str(s or "").lower()
    if "kum" in t:
        return "Kum"
    if "çim" in t or "cim" in t:
        return "Çim"
    if "sentetik" in t:
        return "Sentetik"
    return ""


def _cins(s):
    """Koşu cinsinin KISA hali: '/' öncesi (kosu uzun, son800 kısa yazıyor).
    'Handikap 15 /H2' -> 'HANDIKAP 15' ; 'Maiden /Dişi' -> 'MAIDEN'."""
    base = str(s or "").split("/")[0].strip().upper().replace("İ", "I")
    return re.sub(r"\s+", " ", base)


def _kilo_yukari(s):
    """Kiloyu ÜSTE yuvarlar (son800 buçukları yukarı yuvarladığı için).
    '55,5' -> 56 ; '57' -> 57. Ayrım (tie-break) için; ham eşleşmede kullanılmaz."""
    import math
    t = str(s or "").replace(",", ".").strip()
    m = re.search(r"\d+(?:\.\d+)?", t)
    if not m:
        return None
    return int(math.ceil(float(m.group(0))))


def _anahtar(at, sehir, mesafe, zemin):
    return (_norm(at), _norm(sehir), _mesafe(mesafe), _zemin(zemin))


def fark_ekle(kosu_xlsx, son800_xlsx, cikti_xlsx=None):
    cikti_xlsx = cikti_xlsx or kosu_xlsx
    kdf = pd.read_excel(kosu_xlsx)
    sdf = pd.read_excel(son800_xlsx)

    # Son800 kolonları
    c_at = "At İsmi" if "At İsmi" in sdf.columns else _bul(sdf, "at")
    c_seh = "Şehir Adı" if "Şehir Adı" in sdf.columns else _bul(sdf, "şehir")
    c_mes = "Mesafe" if "Mesafe" in sdf.columns else _bul(sdf, "mesafe")
    c_pist = "Pist" if "Pist" in sdf.columns else _bul(sdf, "pist")
    c_cins = "Koşu Cinsi" if "Koşu Cinsi" in sdf.columns else _bul(sdf, "cins")
    c_kilo = "Kilo" if "Kilo" in sdf.columns else _bul(sdf, "kilo")
    c_fark = "fark" if "fark" in sdf.columns else _bul(sdf, "fark")

    # Ana anahtar: at|şehir|mesafe|zemin. Aynı anahtarda birden çok son800 varsa
    # (72/1071) CİNS(kısa) ile, o da tutmazsa KİLO(üste yuvarlı) ile ayrılır.
    # KİLO ham kullanılmaz — son800 buçukları YUKARI yuvarlıyor; kosu'yu da yukarı
    # yuvarlayınca tutuyor.
    gruplar = defaultdict(list)   # base -> [{fark, cins, kilo}, ...]
    for _, r in sdf.iterrows():
        if pd.isna(r[c_fark]):
            continue
        base = _anahtar(r[c_at], r[c_seh], r[c_mes], r[c_pist])
        gruplar[base].append({"fark": float(r[c_fark]),
                              "cins": _cins(r[c_cins]),
                              "kilo": _kilo_yukari(r[c_kilo])})

    # Koşu Sorgulama kolonları
    k_at = "at adı" if "at adı" in kdf.columns else _bul(kdf, "at")
    k_seh = "şehir" if "şehir" in kdf.columns else _bul(kdf, "şehir")
    k_mes = "mesafe" if "mesafe" in kdf.columns else _bul(kdf, "mesafe")
    k_zem = "zemin" if "zemin" in kdf.columns else _bul(kdf, "zemin")
    k_cins = "cins" if "cins" in kdf.columns else _bul(kdf, "cins")
    k_kilo = "kilo" if "kilo" in kdf.columns else _bul(kdf, "kilo")

    farklar = []
    es = es_cins = es_kilo = es_tek = 0
    for _, r in kdf.iterrows():
        base = _anahtar(r[k_at], r[k_seh], r[k_mes], r[k_zem])
        aday = gruplar.get(base)
        f = None
        if aday:
            if len(aday) == 1:
                sec = 0; es_tek += 1
            else:
                c = _cins(r[k_cins]); kl = _kilo_yukari(r[k_kilo])
                # ÖNCE KİLO (üste yuvarlı) — senin manuel eşleştirme yöntemin.
                # Kilo ayıramazsa CİNS (kısa) yedek.
                sec = next((i for i, a in enumerate(aday) if a["kilo"] == kl), None)
                if sec is not None:
                    es_kilo += 1
                else:
                    sec = next((i for i, a in enumerate(aday) if a["cins"] == c), None)
                    if sec is not None:
                        es_cins += 1
                    else:
                        sec = 0
            f = aday.pop(sec)["fark"]
            es += 1
        farklar.append(f)

    kdf["fark"] = farklar
    kdf.to_excel(cikti_xlsx, index=False)
    oran = es / len(kdf) if len(kdf) else 0
    print(f"[FARK EŞLE] koşu satırı: {len(kdf)} | fark eşleşen: {es} ({oran:.0%}) | "
          f"eşleşmeyen: {len(kdf)-es}")
    print(f"  ayrım: tekil={es_tek}, kilo(üste) ile={es_kilo}, cins ile={es_cins}")
    print(f"  -> {cikti_xlsx}")
    return kdf


def _bul(df, ad):
    for c in df.columns:
        if ad.lower() in str(c).strip().lower():
            return c
    return df.columns[0]


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Kullanım: python3 tjk_fark_esle.py kosu_2026_tam.xlsx son800_2026.xlsx [cikti.xlsx]")
        raise SystemExit(1)
    kosu = sys.argv[1]
    son8 = sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else kosu
    fark_ekle(kosu, son8, out)
