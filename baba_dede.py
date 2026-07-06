#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baba / Dede ayıklama  (Sayfa1 -> Sayfa2)
========================================
TJK_Hamdan_Makroya_Tek_Guzel (Baba)  ve  ..._Dede  karşılığı.

Çıktı kolonları: Kosu_No, At_No, At_Adi_Temiz, Baba|Dede, Kosu_Basligi

mode="baba" -> orijindeki ilk "-" öncesi (sire)
mode="dede" -> orijindeki son "/" sonrası (anne tarafı dede)
"""

from tjk_donustur import (_RowsWS, is_race_title, race_no_from_title,
                          is_horse_row, clean_horse_name)

SAYFA2_BASLIK = ["Kosu_No", "At_No", "At_Adi_Temiz", "Baba", "Kosu_Basligi"]


def _sire_from_origin(v):
    s = "" if v is None else str(v).strip()
    if s == "":
        return ""
    s = s.replace("–", "-").replace("—", "-")
    p = s.find("-")
    return s[:p].strip() if p > 0 else s


def _dede_from_origin(v):
    s = "" if v is None else str(v).strip()
    if s == "":
        return ""
    s = s.replace("–", "-").replace("—", "-")
    p = s.find("/")
    return s[p + 1:].strip() if p >= 0 else ""


def _build(ws, mode):
    # son dolu satır
    last_row = 0
    for r in range(ws.max_row, 0, -1):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1)):
            last_row = r
            break
    rows_out = []
    r = 1
    race_title = ""
    race_no = ""
    while r <= last_row:
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        a_s = "" if a is None else str(a)
        b_s = "" if b is None else str(b)
        if is_race_title(a):
            race_title = a_s.strip()
            race_no = race_no_from_title(race_title)
            r += 1
        elif a_s == "N" and b_s == "At İsmi":
            r += 1
            while r <= last_row:
                a2 = ws.cell(r, 1).value
                b2 = ws.cell(r, 2).value
                if is_race_title(a2):
                    break
                if (("" if a2 is None else str(a2)) == "N"
                        and ("" if b2 is None else str(b2)) == "At İsmi"):
                    break
                if is_horse_row(ws, r):
                    race_no_cell = int(race_no) if str(race_no).isdigit() else race_no
                    at_no = int(float(str(ws.cell(r, 1).value).replace(",", ".")))
                    at_adi = clean_horse_name(ws.cell(r, 2).value)
                    orijin = ws.cell(r, 4).value
                    sire = _sire_from_origin(orijin) if mode == "baba" else _dede_from_origin(orijin)
                    rows_out.append([race_no_cell, at_no, at_adi, sire, race_title])
                r += 1
        else:
            r += 1
    return rows_out


def sayfa1_to_sayfa2(sayfa1_rows, mode="baba"):
    """sayfa1_rows (list[list]) -> Sayfa2 satırları (başlıksız). mode: 'baba'|'dede'."""
    return _build(_RowsWS(sayfa1_rows), mode)
