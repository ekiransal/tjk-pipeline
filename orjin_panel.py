#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ORJIN PANEL ÜRETİCİ  (AnalizCalistir_Sayfa7 karşılığı)
======================================================
Girdi : Sayfa5 (model: sehir+zemin+irk+mesafe -> final+yorum)
        Sayfa6 (at başına ORAN/MESAFE/YAPISI/... verisi; formülle dolu)
        sehir (elle girilen il)
Çıktı : Sayfa7 paneli (= mudur 'orjin'/'dede' sayfası) -> grid {(r,c):val}

Renk/merge uygulanmaz; amaç doğru değer + düzen.
"""

import re


def _trim_text(v):
    if v is None:
        return ""
    s = str(v).strip()
    s = (s.replace("İ", "i").replace("I", "ı").replace("Ç", "ç")
          .replace("Ğ", "ğ").replace("Ö", "ö").replace("Ş", "ş").replace("Ü", "ü"))
    return s.lower()


def _num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip()
    if t == "" or t.upper() in ("#YOK", "#N/A", "#REF!"):
        return None
    t = t.replace("\xa0", "").replace(" ", "")
    if "." in t and "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except Exception:
        return None


def _usable(v):
    return _num(v) is not None


def _sort_parallel(pairs, ascending):
    """VBA SortParallel birebir (stabil DEĞİL; tie sırası makroyla aynı olsun diye)."""
    a = list(pairs)
    n = len(a)
    for i in range(n - 1):
        for j in range(i + 1, n):
            if ascending:
                if a[j][1] < a[i][1]:
                    a[i], a[j] = a[j], a[i]
            else:
                if a[j][1] > a[i][1]:
                    a[i], a[j] = a[j], a[i]
    return a


def _o(rows, r, c):
    if 1 <= r <= len(rows):
        row = rows[r - 1]
        if 1 <= c <= len(row):
            return row[c - 1]
    return None


# ---- race title parse ----
def _parse_race_no(s):
    s = str(s or "")
    p = s.lower().find(". koşu")
    if p < 1:
        p = s.lower().find(". kosu")
    if p > 0:
        m = re.match(r"\s*(\d+)", s[:p][::-1][::-1])
        mm = re.search(r"(\d+)\s*$", s[:p])
        return int(mm.group(1)) if mm else 0
    return 0


def _parse_mesafe(s):
    for tok in re.split(r"[()\-\s]+", str(s or "")):
        if tok.isdigit():
            n = int(tok)
            if 800 <= n <= 4000:
                return n
    return 0


def _parse_zemin(s):
    t = _trim_text(s)
    if "kum" in t:
        return "Kum"
    if "çim" in t or "cim" in t:
        return "Çim"
    if "sentetik" in t:
        return "Sentetik"
    return ""


def _parse_irk(s):
    t = _trim_text(s)
    if "arap" in t:
        return "Arap"
    if "ingiliz" in t:
        return "İngiliz"
    return ""


def _lookup_model(s5, sehir, zemin, irk, mesafe):
    for r in range(2, len(s5) + 1):
        if (_trim_text(_o(s5, r, 1)) == _trim_text(sehir)
                and _trim_text(_o(s5, r, 2)) == _trim_text(zemin)
                and _trim_text(_o(s5, r, 3)) == _trim_text(irk)):
            mv = _num(_o(s5, r, 4))
            if mv is not None and int(mv) == mesafe:
                fv = _o(s5, r, 5)
                yv = _o(s5, r, 6)
                final = "" if (isinstance(fv, str) and fv.strip().upper() in ("#YOK", "#N/A", "#REF!")) else fv
                yorum = "" if yv is None else str(yv)
                return final, yorum
    return "", ""


def _yorum_tipi(yorum):
    t = _trim_text(yorum)
    if t == _trim_text("Kaçak/Stalker"):
        return "KACAK"
    if t == _trim_text("Sprinter/Stalker"):
        return "SPRINTER"
    if t in (_trim_text("Nötr"), _trim_text("KNSZ"), _trim_text("Kazançsız")):
        return "NOTR"
    return ""


def _yorum_tipi_from_final(final):
    f = _num(final)
    if f is None:
        return ""
    if f < -10:
        return "SPRINTER"
    if f > 10:
        return "KACAK"
    return "NOTR"


def _ana_mesafe_tipi(irk, mesafe):
    if _trim_text(irk) == _trim_text("Arap"):
        if mesafe <= 1500:
            return "kısa"
        if 1600 <= mesafe <= 1800:
            return "orta"
        if mesafe >= 1900:
            return "uzun"
    elif _trim_text(irk) == _trim_text("İngiliz"):
        if mesafe <= 1400:
            return "kısa"
        if 1500 <= mesafe <= 1700:
            return "orta"
        if mesafe >= 1800:
            return "uzun"
    return ""


def _fark_sprint_tipi(mesafe):
    if mesafe <= 1600:
        return "kısa"
    if mesafe >= 1700:
        return "uzun"
    return ""


def _mesafe_col(tip, zemin):
    return {
        ("kısa", "kum"): 10, ("kısa", "sentetik"): 11, ("kısa", "çim"): 12,
        ("orta", "kum"): 13, ("orta", "sentetik"): 14, ("orta", "çim"): 15,
        ("uzun", "kum"): 16, ("uzun", "sentetik"): 17, ("uzun", "çim"): 18,
    }.get((_trim_text(tip), _trim_text(zemin)), 0)


def _yapi_col(zemin):
    return {"kum": 21, "çim": 22, "sentetik": 23}.get(_trim_text(zemin), 0)


def _sitil_col(tip):
    return {"kısa": 26, "orta": 27, "uzun": 28}.get(_trim_text(tip), 0)


def _fark_col(tip, zemin):
    return {
        ("kısa", "kum"): 36, ("kısa", "çim"): 37, ("kısa", "sentetik"): 38,
        ("uzun", "kum"): 39, ("uzun", "çim"): 40, ("uzun", "sentetik"): 41,
    }.get((_trim_text(tip), _trim_text(zemin)), 0)


def _oran_sprint_col(tip, zemin):
    return {
        ("kısa", "kum"): 44, ("uzun", "kum"): 46, ("kısa", "sentetik"): 47,
        ("uzun", "sentetik"): 49, ("kısa", "çim"): 50, ("uzun", "çim"): 52,
    }.get((_trim_text(tip), _trim_text(zemin)), 0)


def _s(v):
    return "" if v is None else str(v).strip()


def hesapla(sayfa5, sayfa6, sehir):
    """Sayfa7 paneli grid'i {(r,c):val} döndürür."""
    G = {}
    G[(1, 1)] = "KOŞU ANALİZ"
    out_row = 3
    last = len(sayfa6)
    r_start = 2
    while r_start <= last:
        if _s(_o(sayfa6, r_start, 1)) != "":
            # race end
            first = _s(_o(sayfa6, r_start, 1))
            r_end = last
            for rr in range(r_start + 1, last + 1):
                if _s(_o(sayfa6, rr, 1)) != first:
                    r_end = rr - 1
                    break
            out_row = _panel(G, sayfa5, sayfa6, sehir, r_start, r_end, out_row)
            r_start = r_end + 1
        else:
            r_start += 1
    return G


def _sirali_blok(G, s6, r_start, r_end, top_row, start_col, title,
                 no_col, val_col, ascending, write_value=True,
                 write_final=False, final_val=""):
    pairs = []
    for r in range(r_start, r_end + 1):
        v = _num(_o(s6, r, val_col))
        if v is not None:
            pairs.append((_o(s6, r, no_col), v))
    if not pairs:
        return start_col
    pairs = _sort_parallel(pairs, ascending)
    bw = 2 + (1 if write_value else 0) + (1 if write_final else 0)
    G[(top_row, start_col)] = title
    G[(top_row + 1, start_col)] = "Sıra"
    G[(top_row + 1, start_col + 1)] = "No"
    c = start_col + 2
    val_c = 0
    fin_c = 0
    if write_value:
        G[(top_row + 1, c)] = "Değer"
        val_c = c
        c += 1
    if write_final:
        G[(top_row + 1, c)] = "Final"
        fin_c = c
    fnum = _num(final_val)
    for i, (no, val) in enumerate(pairs, 1):
        G[(top_row + 1 + i, start_col)] = i
        if no is not None and str(no).strip() != "":
            G[(top_row + 1 + i, start_col + 1)] = no
        if val_c:
            G[(top_row + 1 + i, val_c)] = val
        if fin_c and fnum is not None:
            G[(top_row + 1 + i, fin_c)] = fnum
    return start_col + bw + 1


def _ozet_blok(G, s6, r_start, r_end, top_row, start_col, title,
               no_col, val_col, ascending, write_final=False, final_val=""):
    pairs = []
    for r in range(r_start, r_end + 1):
        v = _num(_o(s6, r, val_col))
        if v is not None:
            pairs.append((_o(s6, r, no_col), v))
    if not pairs:
        return start_col
    pairs = _sort_parallel(pairs, ascending)
    G[(top_row, start_col)] = title
    if write_final:
        G[(top_row, start_col + 1)] = "Final"
    fnum = _num(final_val)
    for i, (no, val) in enumerate(pairs, 1):
        if no is not None and str(no).strip() != "":
            G[(top_row + i, start_col)] = no
        if write_final and fnum is not None:
            G[(top_row + i, start_col + 1)] = fnum
    return start_col + (3 if write_final else 2)


def _panel(G, s5, s6, sehir, r_start, r_end, out_row):
    a_text = str(_o(s6, r_start, 1) or "")
    kosu_no = _parse_race_no(a_text)
    mesafe = _parse_mesafe(a_text)
    zemin = _parse_zemin(a_text)
    irk = _parse_irk(a_text)
    final_val, yorum = _lookup_model(s5, sehir, zemin, irk, mesafe)
    yt = _yorum_tipi(yorum)
    if yt == "":
        yt = _yorum_tipi_from_final(final_val)
        if yt == "SPRINTER":
            yorum = "Sprinter/Stalker"
        elif yt == "KACAK":
            yorum = "Kaçak/Stalker"
        elif yt == "NOTR":
            yorum = "KNSZ"
    ana_tipi = _ana_mesafe_tipi(irk, mesafe)
    fark_tipi = _fark_sprint_tipi(mesafe)

    pt = out_row
    G[(pt, 1)] = "Koşu No"; G[(pt, 2)] = kosu_no
    G[(pt, 3)] = "İl"; G[(pt, 4)] = sehir
    G[(pt, 5)] = "Mesafe"; G[(pt, 6)] = mesafe
    G[(pt, 7)] = "Zemin"; G[(pt, 8)] = zemin
    G[(pt, 9)] = "Irk"; G[(pt, 10)] = irk
    G[(pt, 11)] = "Yorum"; G[(pt, 12)] = yorum
    G[(pt, 13)] = "Final"
    fnum = _num(final_val)
    if fnum is not None:
        G[(pt, 14)] = fnum

    block_top = pt + 2
    col = [1]

    def sb(*a, **k):
        col[0] = _sirali_blok(G, s6, r_start, r_end, block_top, col[0], *a, **k)

    sb("ORAN", 2, 7, False)
    mc = _mesafe_col(ana_tipi, zemin)
    if mc > 0:
        sb("MESAFE", 8, mc, False)
    yc = _yapi_col(zemin)
    if yc > 0:
        sb("YAPISI", 19, yc, False)
    sc = _sitil_col(ana_tipi)
    if sc > 0:
        sb("SİTİLİ", 24, sc, False)
    if yt == "KACAK":
        sb("GENEL KAÇAK", 29, 33, False, write_final=True, final_val=final_val)
    elif yt == "SPRINTER":
        sb("GENEL KAÇAK", 29, 33, True, write_final=True, final_val=final_val)
    fc = _fark_col(fark_tipi, zemin)
    if fc > 0:
        if yt == "KACAK":
            sb("FARK SPRINT", 34, fc, True)
        elif yt == "SPRINTER":
            sb("FARK SPRINT", 34, fc, False)
    osc = _oran_sprint_col(fark_tipi, zemin)
    if osc > 0:
        if yt == "KACAK":
            sb("ORAN SPRINT", 42, osc, False)
        elif yt == "SPRINTER":
            sb("ORAN SPRINT", 42, osc, True)
    sb("BABA DOĞU", 53, 57, False)
    sb("EXTREM BABA KUM", 58, 60, False)
    sb("EXTREM BABA ÇİM", 58, 61, False)

    # GENEL ÖZET (sağda, dikey)
    summary_col = col[0] + 2
    # başlık (merge yok) - 'GENEL ÖZET'
    G[(block_top, summary_col)] = "GENEL ÖZET"
    sc2 = [summary_col]

    def ob(*a, **k):
        sc2[0] = _ozet_blok(G, s6, r_start, r_end, block_top + 2, sc2[0], *a, **k)

    ob("ORAN", 2, 7, False)
    if mc > 0:
        ob("MESAFE", 8, mc, False)
    if yc > 0:
        ob("YAPISI", 19, yc, False)
    if sc > 0:
        ob("SİTİLİ", 24, sc, False)
    if yt == "KACAK":
        ob("GENEL KAÇAK", 29, 33, False, write_final=True, final_val=final_val)
    elif yt == "SPRINTER":
        ob("GENEL KAÇAK", 29, 33, True, write_final=True, final_val=final_val)
    if fc > 0:
        if yt == "KACAK":
            ob("FARK SPRINT", 34, fc, True)
        elif yt == "SPRINTER":
            ob("FARK SPRINT", 34, fc, False)
    if osc > 0:
        if yt == "KACAK":
            ob("ORAN SPRINT", 42, osc, False)
        elif yt == "SPRINTER":
            ob("ORAN SPRINT", 42, osc, True)
    ob("BABA DOĞU", 53, 57, False)
    ob("EXTREM BABA KUM", 58, 60, False)
    ob("EXTREM BABA ÇİM", 58, 61, False)

    max_horse = r_end - r_start + 1
    return block_top + max_horse + 6
