#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
800 DOMİNANS — 'yeni yer 800' üretici
=====================================

Amaç: derece'nin "yeni yer" satırlarını temel alıp, dominansın SADECE "eski"
(referans) tarafını 375 günlük 800 referansıyla değiştirmek. Böylece aynı
yapılacak yer makinesine (YYZ.hesapla + orjin/dede + galop/son galop) verince
derece raporunun BİREBİR formatında ama 800 dominanslı ayrı bir sayfa çıkar.

Neyi değiştiriyoruz (1-tabanlı kolon -> 0-tabanlı indeks):
    col 55  Eski HP            -> row[54]
    col 30  Siklet_derece (eski kilo) -> row[29]
    col 66  Medyan HP Listesi  -> row[65]
    col 67  Medyan Kilo Listesi-> row[66]
Yeni HP (col 54) ve yeni kilo (col 29) DEĞİŞMEZ -> "yeni dominans" derece ile
aynı; sadece "eski dominans" 800 referansından gelir.

Veri kaynağı: son800 scraper'ının Veri sayfası (genişletilmiş kolonlar):
    "At Adı", "Eski HP", "Kilo" (eski kilo), "Medyan HP Listesi",
    "Medyan Kilo Listesi", "Durum".
Her at için ilk OK satır (= en yeni 375 günlük 1.'lik = rep 1) alınır.
"""

import pandas as pd
from tjk_donustur import clean_horse_name


# yeni yer satırında değiştirilecek 0-tabanlı indeksler
_IDX_ESKI_HP = 54     # col 55
_IDX_ESKI_KILO = 29   # col 30
_IDX_HP_LISTE = 65    # col 66
_IDX_KILO_LISTE = 66  # col 67
_IDX_AT_ADI = 2       # col 3  (At_Adi_Temiz)


def _get(d, ad, alt=""):
    if ad in d and d[ad] is not None:
        return d[ad]
    hedef = str(ad).strip().upper()
    for k in d.keys():
        if str(k).strip().upper() == hedef:
            v = d[k]
            return alt if v is None else v
    return alt


def _s(v):
    """Hücreyi güvenli metne çevir (NaN/None -> '')."""
    if v is None:
        return ""
    if isinstance(v, float) and v != v:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "<na>") else s


def son800_ref_map(veri_dosya, sheet="Veri"):
    """son800 Veri dosyasından KOŞU BAZINDA referans haritası döndürür.
    Anahtar = kosu_key(at+tarih+şehir+zemin+mesafe) -> {eski_hp,eski_kilo,hp_liste,kilo_liste}.
    Ayrıca at bazında (rep-1) yedek anahtar da eklenir ('AT::<ad>') — koşu eşleşmezse
    kullanılır. Böylece her geçmiş koşu KENDİ son800'üne göre dominans alır."""
    ref = {}
    try:
        xls = pd.ExcelFile(veri_dosya)
        sh = sheet if sheet in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(veri_dosya, sheet_name=sh)
    except Exception as e:
        print(f"[800 DOMİNANS] son800 Veri okunamadı (800 sayfası atlanır): {e}")
        return ref

    kosu_say = 0
    for d in df.to_dict(orient="records"):
        durum = _s(_get(d, "Durum")).upper()
        if durum and durum != "OK":
            continue
        atadi = _s(_get(d, "At Adı"))
        if atadi == "":
            continue
        ad = clean_horse_name(atadi)
        if ad == "":
            continue
        kayit = {
            "eski_hp": _s(_get(d, "Eski HP")),
            "eski_kilo": _s(_get(d, "Kilo")),
            "hp_liste": _s(_get(d, "Medyan HP Listesi")),
            "kilo_liste": _s(_get(d, "Medyan Kilo Listesi")),
        }
        # KOŞU anahtarı (son800 dosyasındaki kolonlar: Şehir/Pist/Mesafe/Tarih)
        kk = kosu_key(atadi, _get(d, "Tarih"), _get(d, "Şehir"),
                      _get(d, "Pist"), _get(d, "Mesafe"))
        if kk not in ref:
            ref[kk] = kayit
            kosu_say += 1
        # at bazında yedek (ilk OK satır)
        atkey = "AT::" + ad
        if atkey not in ref:
            ref[atkey] = kayit
    print(f"[800 DOMİNANS] son800 referans haritası: {kosu_say} koşu (+ at-yedek)")
    return ref


def _dt_norm(x):
    """Tarih/hücreyi anahtar için kararlı metne çevir (datetime -> dd.mm.yyyy)."""
    if x is None:
        return ""
    try:
        import datetime as _dt
        if isinstance(x, (_dt.datetime, _dt.date)):
            return x.strftime("%d.%m.%Y")
    except Exception:
        pass
    s = str(x).strip()
    return s.upper()


def kosu_key(at, tarih, sehir, zemin, mesafe):
    """Bir koşuyu benzersiz tanımlayan anahtar: at+tarih+şehir+zemin+mesafe.
    Aynı at farklı koşularda farklı dominans alsın diye KOŞU bazında kullanılır."""
    from tjk_donustur import clean_horse_name
    a = clean_horse_name(str(at or "")).upper()
    mes = _dt_norm(mesafe)
    # mesafe '1400' / 1400 / 1400.0 -> '1400'
    try:
        mes = str(int(float(str(mesafe).replace(",", "."))))
    except Exception:
        pass
    return "|".join([a, _dt_norm(tarih), _dt_norm(sehir), _dt_norm(zemin), mes])


def dominans_haritasi(ana_rows_with_header):
    """800-transformed ana satırlardan (başlık DAHİL, satır0=başlık) HER KOŞU için
    6'lı dominansı DOĞRUDAN hesaplar -> {kosu_key: (d50hp,d50k,d66hp,d66k,d75hp,d75k)}.
    Dominans KOŞU bazındadır (yeni HP/kilo her koşuda farklı; eski taraf 800 referansı
    at bazında sabit). Bu yüzden aynı atın farklı koşuları farklı dominans alır.
    Anahtar = at+tarih+şehir+zemin+mesafe. Grid'in kaydırma hatasına düşmeden doğru
    değer yerleştirmek için yazarken aynı anahtar grid satırından kurulur."""
    import yapilacak_yer as YZ
    from tjk_donustur import clean_horse_name
    M = YZ.AnaTablo(ana_rows_with_header)
    harita = {}
    for i in range(1, len(ana_rows_with_header)):
        row = ana_rows_with_header[i]
        at = clean_horse_name(str(YZ._ham(row, 3)))
        if not at:
            continue
        # ana satır: at=col3, tarih=col7, şehir=col8, zemin=col9, mesafe=col11
        key = kosu_key(YZ._ham(row, 3), YZ._ham(row, 7), YZ._ham(row, 8),
                       YZ._ham(row, 9), YZ._ham(row, 11))
        if key in harita:
            continue
        harita[key] = (
            M._hp_dominans(row, 50), M._kilo_dominans(row, 50),
            M._hp_dominans(row, 66.66), M._kilo_dominans(row, 66.66),
            M._hp_dominans(row, 75), M._kilo_dominans(row, 75),
        )
    return harita


# derece 'yeni yer' satırındaki GEÇMİŞ KOŞU kimliği (0-tabanlı):
#   At = col3(idx2) | Tarih_derece = col34(idx33) | Sehir_derece = col35(idx34)
#   Pist_derece = col36(idx35) | Mesafe_derece = col38(idx37)
_IDX_D_TARIH = 33
_IDX_D_SEHIR = 34
_IDX_D_PIST = 35
_IDX_D_MESAFE = 37


_IDX_STIL = 77   # col 77 = BY_derece_Stil (%45 Seyir Sırası)


def stil_haritasi(ana_rows_with_header):
    """Yeni yer satırlarından her KOŞU için %45 seyir sırasını (stil) döndürür ->
    {kosu_key: stil}. Anahtar dominans haritasıyla aynı (at+tarih+şehir+zemin+mesafe)."""
    import yapilacak_yer as YZ
    from tjk_donustur import clean_horse_name
    harita = {}
    for i in range(1, len(ana_rows_with_header)):
        row = ana_rows_with_header[i]
        at = clean_horse_name(str(YZ._ham(row, 3)))
        if not at:
            continue
        key = kosu_key(YZ._ham(row, 3), YZ._ham(row, 7), YZ._ham(row, 8),
                       YZ._ham(row, 9), YZ._ham(row, 11))
        if key in harita:
            continue
        harita[key] = _s(YZ._ham(row, _IDX_STIL))
    return harita


def stil_ucgen(stil):
    """%45 seyir sırası -> 4 kademeli üçgen (derecedeki mantığın aynısı).
    1 (kaçak) -> ▷▷▷▶ | 2-3 -> ▷▷▶▷ | 4-6 -> ▷▶▷▷ | 7+ -> ▶▷▷▷. Boşsa ''."""
    try:
        s = int(float(str(stil).replace(",", ".")))
    except Exception:
        return ""
    if s <= 1:
        idx = 3
    elif s <= 3:
        idx = 2
    elif s <= 6:
        idx = 1
    else:
        idx = 0
    return "".join("▶" if i == idx else "▷" for i in range(4))


def yeni_yer_800_uret(yeniyer_rows, ref_map):
    """derece 'yeni yer' satırlarını 800 referansıyla dönüştürür — KOŞU BAZINDA.
    Her satır (geçmiş koşu) KENDİ son800'üyle eşleştirilir (at+tarih+şehir+zemin+mesafe);
    böylece eskiHP ve saha listeleri her koşuda farklı -> dominans koşu bazında.
    Koşu eşleşmezse at-yedeğine (AT::ad) düşer; o da yoksa eski taraf boşlanır."""
    out = []
    _kosu_es = _at_es = _yok = 0
    for r in yeniyer_rows:
        r2 = list(r)
        while len(r2) <= _IDX_KILO_LISTE:
            r2.append("")
        ad = clean_horse_name(_s(r2[_IDX_AT_ADI])) if len(r2) > _IDX_AT_ADI else ""
        # önce KOŞU anahtarıyla dene
        kk = None
        if len(r2) > _IDX_D_MESAFE:
            kk = kosu_key(r2[_IDX_AT_ADI], r2[_IDX_D_TARIH], r2[_IDX_D_SEHIR],
                          r2[_IDX_D_PIST], r2[_IDX_D_MESAFE])
        ref = ref_map.get(kk) if kk else None
        if ref:
            _kosu_es += 1
        else:
            ref = ref_map.get("AT::" + ad)  # at-yedek
            if ref:
                _at_es += 1
        if ref:
            r2[_IDX_ESKI_HP] = ref.get("eski_hp", "")
            r2[_IDX_ESKI_KILO] = ref.get("eski_kilo", "")
            r2[_IDX_HP_LISTE] = ref.get("hp_liste", "")
            r2[_IDX_KILO_LISTE] = ref.get("kilo_liste", "")
        else:
            _yok += 1
            r2[_IDX_ESKI_HP] = ""
            r2[_IDX_ESKI_KILO] = ""
            r2[_IDX_HP_LISTE] = ""
            r2[_IDX_KILO_LISTE] = ""
        out.append(r2)
    print(f"[800 DOMİNANS] eşleşme: koşu={_kosu_es} at-yedek={_at_es} yok={_yok}")
    return out


def son800_stil_haritasi(veri_dosya, sheet="Veri_Stil_Ekli"):
    """son800 %45 VİDEO çıktısından ('..._STIL_EKLI.xlsx') her KOŞU için %45 seyir
    sırasını döndürür -> {kosu_key: stil}. (OPTION B: derece 183 gün + AYRI son800
    375 gün video fazı.)  Anahtar = kosu_key(At Adı, Tarih, Şehir, Pist, Mesafe) —
    yapılacak yer 800 AT satırlarının anahtarıyla AYNI. Dosya yoksa ya da stil kolonu
    boşsa boş harita döner; o durumda yapılacak yer 800 stili derece %45'inden gelir."""
    import os as _os
    harita = {}
    if not _os.path.exists(veri_dosya):
        print(f"[800 STİL] son800 %45 dosyası yok ({veri_dosya}) -> derece %45'i kullanılır")
        return harita
    try:
        xls = pd.ExcelFile(veri_dosya)
        sh = sheet if sheet in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(veri_dosya, sheet_name=sh)
    except Exception as e:
        print(f"[800 STİL] son800 %45 okunamadı ({e}) -> derece %45'i kullanılır")
        return harita
    # %45 seyir sırası kolonu; yoksa temel 'Stil' kolonuna düş
    stil_kol = None
    for c in df.columns:
        cs = str(c).strip().lower().replace("i̇", "i")
        if "45" in cs and "seyir" in cs:
            stil_kol = c
            break
    if stil_kol is None:
        for c in df.columns:
            if str(c).strip().lower() == "stil":
                stil_kol = c
                break
    if stil_kol is None:
        print("[800 STİL] son800 %45/Stil kolonu bulunamadı -> derece %45'i kullanılır")
        return harita
    n = 0
    for d in df.to_dict(orient="records"):
        stil = _s(d.get(stil_kol))
        if stil == "":
            continue
        kk = kosu_key(_get(d, "At Adı"), _get(d, "Tarih"), _get(d, "Şehir"),
                      _get(d, "Pist"), _get(d, "Mesafe"))
        if kk and kk not in harita:
            harita[kk] = stil
            n += 1
    print(f"[800 STİL] son800 %45 stil (ayrı video fazı): {n} koşu")
    return harita
